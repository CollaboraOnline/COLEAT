/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <cassert>
#include <cstdio>
#include <cstdlib>
#include <cwchar>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>

#include <Windows.h>

// DbgHelp.h has even more sloppier code than <Windows.h>
#pragma warning(push)
#pragma warning(disable : 4091)
#include <DbgHelp.h>
#pragma warning(pop)

#pragma warning(pop)

#include "exewrapper.hpp"
#include "utils.hpp"

#include "CProxiedCoclass.hpp"
#include "CProxiedDispatch.hpp"

#include "InterfaceMapping.hxx"

static ThreadProcParam* pGlobalParamPtr;

static int nHookedFunctions = 0;

static bool hook(ThreadProcParam* pParam, HMODULE hModule, const wchar_t* sModuleName,
                 const wchar_t* sDll, const wchar_t* sFunction, PVOID pOwnFunction);

static HMODULE WINAPI myLoadLibraryExW(LPCWSTR lpFileName, HANDLE hFile, DWORD dwFlags);

static void storeError(ThreadProcParam* pParam, const wchar_t* pPrefix,
                       const wchar_t* pPrefix2 = nullptr, const wchar_t* pPrefix3 = nullptr,
                       const wchar_t* pPrefix4 = nullptr)
{
    LPWSTR pMsgBuf;
    wcscpy_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION, pPrefix);
    if (pPrefix2)
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), pPrefix2);
    if (pPrefix3)
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), pPrefix3);
    if (pPrefix4)
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), pPrefix4);

    if (GetWindowsErrorString(GetLastError(), &pMsgBuf))
    {
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), L": ");
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), pMsgBuf);
        HeapFree(GetProcessHeap(), 0, pMsgBuf);
    }
}

static HRESULT WINAPI myCoCreateInstance(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext,
                                         REFIID riid, LPVOID* ppv)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "myCoCreateInstance(" << rclsid << ", " << riid << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << std::endl;

    for (int i = 0; i < sizeof(aInterfaceMap) / sizeof(aInterfaceMap[0]); ++i)
    {
        if (IsEqualIID(rclsid, aInterfaceMap[i].maFromCoclass))
        {
            CProxiedCoclass* pCoclass = new CProxiedCoclass(aInterfaceMap[i]);
            pCoclass->AddRef();
            *ppv = pCoclass;
            return S_OK;
        }
    }

    return CoCreateInstance(rclsid, pUnkOuter, dwClsContext, riid, ppv);
}

static HRESULT WINAPI myCoCreateInstanceEx(REFCLSID clsid, LPUNKNOWN pUnkOuter, DWORD dwClsCtx,
                                           COSERVERINFO* pServerInfo, DWORD dwCount,
                                           MULTI_QI* pResults)
{
    if (pGlobalParamPtr->mbVerbose)
    {
        std::cout << "myCoCreateInstanceEx(" << clsid << ", " << dwCount << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << std::endl;
        for (DWORD j = 0; j < dwCount; ++j)
            std::cout << "   " << *pResults[j].pIID << "\n";
        std::cout << std::flush;
    }

    for (int i = 0; i < sizeof(aInterfaceMap) / sizeof(aInterfaceMap[0]); ++i)
    {
        if (IsEqualIID(clsid, aInterfaceMap[i].maFromCoclass))
        {
            CProxiedCoclass* pCoclass = new CProxiedCoclass(aInterfaceMap[i]);

            DWORD nSuccess = 0;
            for (DWORD j = 0; j < dwCount; ++j)
            {
                pResults[j].hr
                    = pCoclass->QueryInterface(*pResults[j].pIID, (void**)&pResults[j].pItf);
                if (SUCCEEDED(pResults[j].hr))
                    ++nSuccess;
            }
            if (nSuccess == dwCount)
                return S_OK;
            else if (nSuccess == 0)
                return E_NOINTERFACE;
            else
                return CO_S_NOTALLINTERFACES;
        }
    }

    return CoCreateInstanceEx(clsid, pUnkOuter, dwClsCtx, pServerInfo, dwCount, pResults);
}

static PROC WINAPI myGetProcAddress(HMODULE hModule, LPCSTR lpProcName)
{
    HMODULE hOle32 = GetModuleHandleW(L"ole32.dll");
    FunPtr pFun;

    if (hModule == hOle32 && std::strcmp(lpProcName, "CoCreateInstanceEx") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "myGetProcAddress(CoCreateInstanceEx) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myCoCreateInstanceEx;
        return pFun.pProc;
    }

    if (hModule == hOle32 && std::strcmp(lpProcName, "CoCreateInstance") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "myGetProcAddress(CoCreateInstance) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myCoCreateInstance;
        return pFun.pProc;
    }

    return GetProcAddress(hModule, lpProcName);
}

static HMODULE WINAPI myLoadLibraryW(LPCWSTR lpFileName)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "myLoadLibraryW(" << convertUTF16ToUTF8(lpFileName) << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << std::endl;

    HMODULE hModule = LoadLibraryW(lpFileName);

    if (hModule != NULL)
    {
        hook(pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"GetProcAddress",
             myGetProcAddress);
        hook(pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"LoadLibraryW",
             myLoadLibraryW);
        hook(pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"LoadLibraryExW",
             myLoadLibraryExW);
        hook(pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", L"CoCreateInstance",
             myCoCreateInstance);
        hook(pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", L"CoCreateInstanceEx",
             myCoCreateInstanceEx);
    }

    return hModule;
}

static HMODULE WINAPI myLoadLibraryExW(LPCWSTR lpFileName, HANDLE hFile, DWORD dwFlags)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "myLoadLibraryExW(" << convertUTF16ToUTF8(lpFileName) << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << std::endl;

    HMODULE hModule = LoadLibraryExW(lpFileName, hFile, dwFlags);

    if (hModule != NULL)
    {
        hook(pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"GetProcAddress",
             myGetProcAddress);
        hook(pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"LoadLibraryW",
             myLoadLibraryW);
        hook(pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"LoadLibraryExW",
             myLoadLibraryExW);
        hook(pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", L"CoCreateInstance",
             myCoCreateInstance);
        hook(pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", L"CoCreateInstanceEx",
             myCoCreateInstanceEx);
    }

    return hModule;
}

// The functions below, InjectedDllMainFunction() and the functions it calls, can not write to
// std::cout and std::cerr. They run in a thread created before those have been set up. They can,
// however, use Win32 API directly.

static bool hook(ThreadProcParam* pParam, HMODULE hModule, const wchar_t* sModuleName,
                 const wchar_t* sDll, const wchar_t* sFunction, PVOID pOwnFunction)
{
    ULONG nSize;
    PIMAGE_IMPORT_DESCRIPTOR pImportDescriptor
        = (PIMAGE_IMPORT_DESCRIPTOR)ImageDirectoryEntryToDataEx(
            hModule, TRUE, IMAGE_DIRECTORY_ENTRY_IMPORT, &nSize, NULL);
    if (pImportDescriptor == NULL)
    {
        storeError(pParam, L"Could not find import directory in ", sModuleName);
        return false;
    }

    bool bFound = false;
    while (pImportDescriptor->Characteristics && pImportDescriptor->Name)
    {
        PSTR sName = (PSTR)((PBYTE)hModule + pImportDescriptor->Name);
        if (_stricmp(sName, convertUTF16ToUTF8(sDll).data()) == 0)
        {
            bFound = true;
            break;
        }
        pImportDescriptor++;
    }

    if (!bFound)
    {
        wcscpy_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION,
                 L"Import descriptor for ");
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), sDll);
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation),
                 L" not found in ");
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), sModuleName);
        return false;
    }

    PROC pOriginalFunc
        = (PROC)GetProcAddress(GetModuleHandleW(sDll), convertUTF16ToUTF8(sFunction).data());
    if (pOriginalFunc == NULL)
    {
        storeError(pParam, L"Could not find original ", sFunction, L" in ", sDll);
        return false;
    }

    PIMAGE_THUNK_DATA pThunk = (PIMAGE_THUNK_DATA)((PBYTE)hModule + pImportDescriptor->FirstThunk);
    while (pThunk->u1.Function)
    {
        PROC* pFunc = (PROC*)&pThunk->u1.Function;
        if (*pFunc == pOriginalFunc)
        {
            MEMORY_BASIC_INFORMATION aMBI;
            if (!VirtualQuery(pFunc, &aMBI, sizeof(MEMORY_BASIC_INFORMATION)))
            {
                storeError(pParam, L"VirtualQuery failed");
                return false;
            }

            if (!VirtualProtect(aMBI.BaseAddress, aMBI.RegionSize, PAGE_READWRITE, &aMBI.Protect))
            {
                storeError(pParam, L"VirtualProtect failed");
                return false;
            }

            FunPtr pFun;
            pFun.pVoid = pOwnFunction;
            *pFunc = pFun.pProc;

            DWORD nOldProtect;
            if (!VirtualProtect(aMBI.BaseAddress, aMBI.RegionSize, aMBI.Protect, &nOldProtect))
            {
                storeError(pParam, L"VirtualProtect failed");
                return false;
            }

            nHookedFunctions++;
            return true;
        }

        pThunk++;
    }

    wcscpy_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION, L"Did not find ");
    wcscat_s(pParam->msErrorExplanation,
             ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), sFunction);
    wcscat_s(pParam->msErrorExplanation,
             ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation),
             L" in import table for ");
    wcscat_s(pParam->msErrorExplanation,
             ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), sDll);
    wcscat_s(pParam->msErrorExplanation,
             ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), L" in ");
    wcscat_s(pParam->msErrorExplanation,
             ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), sModuleName);

    return false;
}

static bool hook(ThreadProcParam* pParam, const wchar_t* sModule, const wchar_t* sDll,
                 const wchar_t* sFunction, PVOID pOwnFunction)
{
    const wchar_t* sModuleName;

    sModuleName = sModule;

    if (!sModule)
    {
        const DWORD NFILENAME = 1000;
        static wchar_t sFileName[NFILENAME];

        DWORD nSizeOut = GetModuleFileNameW(NULL, sFileName, NFILENAME);
        if (nSizeOut == 0 || nSizeOut == NFILENAME)
            sModuleName = L".exe file of the process";
        else
            sModuleName = baseName(sFileName);
    }

    HMODULE hModule = GetModuleHandleW(sModule);
    if (hModule == NULL)
    {
        storeError(pParam, sModuleName, L" is not loaded");
        return false;
    }

    return hook(pParam, hModule, sModuleName, sDll, sFunction, pOwnFunction);
}

extern "C" DWORD WINAPI InjectedDllMainFunction(ThreadProcParam* pParam)
{
// Magic to export this function using a plain undecorated name despite it being WINAPI
#pragma comment(linker, "/EXPORT:" __FUNCTION__ "=" __FUNCDNAME__)

    if (pParam->mnSize != sizeof(*pParam))
    {
        // We can't really store any error message in pParam->msErrorExplanation if the structure of
        // *pParam is not what we expect;)
        return FALSE;
    }

    pParam->mbPassedSizeCheck = true;

    // This function returns and the remotely created thread exits, and the wrapper process will
    // copy back the parameter block, but we keep a pointer to it for use by the hook functions.
    pGlobalParamPtr = pParam;

    CProxiedUnknown::setParam(pGlobalParamPtr);

    tryToEnsureStdHandlesOpen();

    // Do our IAT patching. We want to hook CoCreateInstance() and CoCreateInstanceEx().

    HMODULE hMsvbvm60 = GetModuleHandleW(L"msvbvm60.dll");
    HMODULE hOle32 = GetModuleHandleW(L"ole32.dll");

    // FIXME: We import ole32.dll ourselves thanks to the IID operator<< in utils.hpp. Would be
    // better if we didn't.
    (void)hOle32;
    if (hMsvbvm60 != NULL /* && hOle32 == NULL */)
    {
        // It is most likely an exe created by VB6. Hook the COM object creation functions in
        // msvbvm60.dll.

        // Msvbvm60.dll seems to import just CoCreateInstance() directly, it looks up
        // CoCreateInstanceEx() with GetProcAddress(). (But try to hook CoCreateInstanceEx() anyway,
        // just in case.) Thus we need to hook GetProcAddress() too.

        if (!hook(pParam, L"msvbvm60.dll", L"ole32.dll", L"CoCreateInstance", myCoCreateInstance))
            return FALSE;

        hook(pParam, L"msvbvm60.dll", L"ole32.dll", L"CoCreateInstanceEx", myCoCreateInstanceEx);

        if (!hook(pParam, L"msvbvm60.dll", L"kernel32.dll", L"GetProcAddress", myGetProcAddress))
            return FALSE;
    }
    else
    {
        // It is some other executable. We must hook LoadLibraryW().

        nHookedFunctions = 0;
        hook(pParam, nullptr, L"kernel32.dll", L"GetProcAddress", myGetProcAddress);
        hook(pParam, nullptr, L"kernel32.dll", L"LoadLibraryW", myLoadLibraryW);
        hook(pParam, nullptr, L"kernel32.dll", L"LoadLibraryExW", myLoadLibraryExW);
        hook(pParam, nullptr, L"ole32.dll", L"CoCreateInstance", myCoCreateInstance);
        hook(pParam, nullptr, L"ole32.dll", L"CoCreateInstanceEx", myCoCreateInstanceEx);

        if (nHookedFunctions == 0)
        {
            wcscpy_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION,
                     L"Could not hook a single interesting function");
            return FALSE;
        }
    }

    return TRUE;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

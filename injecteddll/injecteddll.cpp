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

static bool hook(bool bMandatory, ThreadProcParam* pParam, HMODULE hModule,
                 const wchar_t* sModuleName, const wchar_t* sDll, const wchar_t* sFunction,
                 PVOID pOwnFunction);

static HMODULE WINAPI myLoadLibraryExW(LPCWSTR lpFileName, HANDLE hFile, DWORD dwFlags);

static void showMessage(bool bIsError, ThreadProcParam* pParam)
{
    if (pParam->mbPassedInjectedThread)
    {
        if (bIsError)
            std::cerr << convertUTF16ToUTF8(pParam->msErrorExplanation) << std::endl;
        else if (pParam->mbVerbose)
            std::cout << convertUTF16ToUTF8(pParam->msErrorExplanation) << std::endl;
    }
    else
    {
        pParam->mbMessageIsError = bIsError;
        if (SetEvent(pParam->mhMessageEvent))
            WaitForSingleObject(pParam->mhMessageEvent, 1000);
    }
}

static void storeMessage(bool bIsError, ThreadProcParam* pParam, const wchar_t* pMsg,
                         const wchar_t* pMsg2 = nullptr, const wchar_t* pMsg3 = nullptr,
                         const wchar_t* pMsg4 = nullptr, const wchar_t* pMsg5 = nullptr,
                         const wchar_t* pMsg6 = nullptr)
{
    LPWSTR pMsgBuf;
    wcscpy_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION, pMsg);
    if (pMsg2)
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), pMsg2);
    if (pMsg3)
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), pMsg3);
    if (pMsg4)
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), pMsg4);
    if (pMsg5)
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), pMsg5);
    if (pMsg6)
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), pMsg6);

    if (bIsError && pParam->mbPassedInjectedThread
        && GetWindowsErrorString(GetLastError(), &pMsgBuf))
    {
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), L": ");
        wcscat_s(pParam->msErrorExplanation,
                 ThreadProcParam::NEXPLANATION - wcslen(pParam->msErrorExplanation), pMsgBuf);
        HeapFree(GetProcessHeap(), 0, pMsgBuf);
    }

    showMessage(bIsError, pParam);
}

static void printCreateInstanceResult(void* pV)
{
    HRESULT nResult;
    IDispatch* pDispatch = NULL;

    // Silence verbosity while doing the QueryInterface() etc calls here, they might go through our
    // proxies, and it is pointless to do verbose logging those proxies for caused by fetching of
    // information for verbose logging...
    bool bWasVerbose = pGlobalParamPtr->mbVerbose;
    pGlobalParamPtr->mbVerbose = false;

    nResult = ((IUnknown*)pV)->QueryInterface(IID_IDispatch, (void**)&pDispatch);

    ITypeInfo* pTI = NULL;
    if (nResult == S_OK)
        nResult = pDispatch->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTI);

    BSTR sTypeName = NULL;
    if (nResult == S_OK)
        nResult = pTI->GetDocumentation(MEMBERID_NIL, &sTypeName, NULL, NULL, NULL);

    ITypeLib* pTL = NULL;
    UINT nIndex;
    if (nResult == S_OK)
        nResult = pTI->GetContainingTypeLib(&pTL, &nIndex);

    BSTR sLibName = NULL;
    if (nResult == S_OK)
        nResult = pTL->GetDocumentation(-1, &sLibName, NULL, NULL, NULL);

    std::cout << pV << " (" << (sLibName != NULL ? convertUTF16ToUTF8(sLibName) : "?") << "."
              << (sTypeName != NULL ? convertUTF16ToUTF8(sTypeName) : "?") << ")" << std::endl;

    if (sTypeName != NULL)
        SysFreeString(sTypeName);
    if (sLibName != NULL)
        SysFreeString(sLibName);

    pGlobalParamPtr->mbVerbose = bWasVerbose;
}

static HRESULT WINAPI myCoCreateInstance(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext,
                                         REFIID riid, LPVOID* ppv)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "CoCreateInstance(" << rclsid << ", " << riid << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << std::endl;

    // Is it one of the interfaces we have generated proxies for? In that case return a proxy for it
    // to the client.
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

    HRESULT nResult = CoCreateInstance(rclsid, pUnkOuter, dwClsContext, riid, ppv);

    if (nResult == S_OK && pGlobalParamPtr->mbVerbose)
    {
        std::cout << "...CoCreateInstance(" << rclsid << ", " << riid << ")"
                  << " -> ";
        printCreateInstanceResult(*ppv);
    }

    return nResult;
}

static HRESULT WINAPI myCoCreateInstanceEx(REFCLSID clsid, LPUNKNOWN pUnkOuter, DWORD dwClsCtx,
                                           COSERVERINFO* pServerInfo, DWORD dwCount,
                                           MULTI_QI* pResults)
{
    if (pGlobalParamPtr->mbVerbose)
    {
        std::cout << "CoCreateInstanceEx(" << clsid << ", " << dwCount << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << std::endl;
        for (DWORD j = 0; j < dwCount; ++j)
            std::cout << "   " << j << "(" << dwCount << "): " << *pResults[j].pIID << std::endl;
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
                if (pResults[j].hr == S_OK)
                    ++nSuccess;
            }
            if (pGlobalParamPtr->mbVerbose)
            {
                for (DWORD j = 0; j < dwCount; ++j)
                {
                    std::cout << "...CoCreateInstanceEx(" << clsid << "): " << j << "(" << dwCount
                              << "): "
                              << ": ";
                    if (pResults[j].hr == S_OK)
                    {
                        std::cout << *pResults[j].pIID << ": ";
                        printCreateInstanceResult(pResults[j].pItf);
                    }
                    else
                        std::cout << HRESULT_to_string(pResults[j].hr) << std::endl;
                    ;
                }
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
            std::cout << "GetProcAddress(CoCreateInstanceEx) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myCoCreateInstanceEx;
        return pFun.pProc;
    }

    if (hModule == hOle32 && std::strcmp(lpProcName, "CoCreateInstance") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(CoCreateInstance) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myCoCreateInstance;
        return pFun.pProc;
    }

    return GetProcAddress(hModule, lpProcName);
}

static HMODULE WINAPI myLoadLibraryW(LPCWSTR lpFileName)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "LoadLibraryW(" << convertUTF16ToUTF8(lpFileName) << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << std::endl;

    HMODULE hModule = LoadLibraryW(lpFileName);
    DWORD nLastError = GetLastError();

    if (pGlobalParamPtr->mbVerbose)
        std::cout << "...LoadLibraryW(" << convertUTF16ToUTF8(lpFileName) << "): " << hModule;

    if (hModule == NULL)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << ": " << WindowsErrorString(nLastError) << std::endl;
    }
    else
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << std::endl;

        hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"GetProcAddress",
             myGetProcAddress);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"LoadLibraryW",
             myLoadLibraryW);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"LoadLibraryExW",
             myLoadLibraryExW);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", L"CoCreateInstance",
             myCoCreateInstance);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", L"CoCreateInstanceEx",
             myCoCreateInstanceEx);
    }

    return hModule;
}

static HMODULE WINAPI myLoadLibraryExW(LPCWSTR lpFileName, HANDLE hFile, DWORD dwFlags)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "LoadLibraryExW(" << convertUTF16ToUTF8(lpFileName) << ", " << to_uhex(dwFlags)
                  << ") from " << prettyCodeAddress(_ReturnAddress()) << std::endl;

    HMODULE hModule = LoadLibraryExW(lpFileName, hFile, dwFlags);
    DWORD nLastError = GetLastError();

    if (hModule == NULL)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "...LoadLibraryExW(" << convertUTF16ToUTF8(lpFileName) << ", "
                      << to_uhex(dwFlags) << "): " << hModule << ": "
                      << WindowsErrorString(nLastError) << std::endl;
    }
    else
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "...LoadLibraryExW(" << convertUTF16ToUTF8(lpFileName) << ", "
                      << to_uhex(dwFlags) << "): " << hModule << std::endl;

        if (!(dwFlags
              & (LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE
                 | LOAD_LIBRARY_AS_IMAGE_RESOURCE)))
        {
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"GetProcAddress",
                 myGetProcAddress);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"LoadLibraryW",
                 myLoadLibraryW);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", L"LoadLibraryExW",
                 myLoadLibraryExW);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", L"CoCreateInstance",
                 myCoCreateInstance);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", L"CoCreateInstanceEx",
                 myCoCreateInstanceEx);
        }
    }

    return hModule;
}

// The functions below, InjectedDllMainFunction() and the functions it calls, can not write to
// std::cout and std::cerr. They run in a thread created before those have been set up. They can,
// however, use Win32 API directly.

static bool hook(bool bMandatory, ThreadProcParam* pParam, HMODULE hModule,
                 const wchar_t* sModuleName, const wchar_t* sDll, const wchar_t* sFunction,
                 PVOID pOwnFunction)
{
    ULONG nSize;
    PIMAGE_IMPORT_DESCRIPTOR pImportDescriptor
        = (PIMAGE_IMPORT_DESCRIPTOR)ImageDirectoryEntryToDataEx(
            hModule, TRUE, IMAGE_DIRECTORY_ENTRY_IMPORT, &nSize, NULL);
    if (pImportDescriptor == NULL)
    {
        if (bMandatory)
            storeMessage(bMandatory, pParam, L"Could not find import directory in ", sModuleName);
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
        if (bMandatory)
            storeMessage(bMandatory, pParam, L"Import descriptor for ", sDll, L" not found in ",
                         sModuleName);

        return false;
    }

    PROC pOriginalFunc
        = (PROC)GetProcAddress(GetModuleHandleW(sDll), convertUTF16ToUTF8(sFunction).data());
    if (pOriginalFunc == NULL)
    {
        storeMessage(bMandatory, pParam, L"Could not find original ", sFunction, L" in ", sDll);
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
                storeMessage(bMandatory, pParam, L"VirtualQuery failed");
                return false;
            }

            if (!VirtualProtect(aMBI.BaseAddress, aMBI.RegionSize, PAGE_READWRITE, &aMBI.Protect))
            {
                storeMessage(bMandatory, pParam, L"VirtualProtect failed");
                return false;
            }

            FunPtr pFun;
            pFun.pVoid = pOwnFunction;
            *pFunc = pFun.pProc;

            DWORD nOldProtect;
            if (!VirtualProtect(aMBI.BaseAddress, aMBI.RegionSize, aMBI.Protect, &nOldProtect))
            {
                storeMessage(bMandatory, pParam, L"VirtualProtect failed");
                return false;
            }

            nHookedFunctions++;

            storeMessage(false, pParam, L"Hooked ", sFunction, L" in import table for ", sDll,
                         L" in ", sModuleName);
            return true;
        }

        pThunk++;
    }

    // Don't bother mentioning non-mandatory imports not found even in verbose mode.
    if (bMandatory)
        storeMessage(bMandatory, pParam, L"Did not find ", sFunction, L" in import table for ",
                     sDll, L" in ", sModuleName);

    return false;
}

static bool hook(bool bMandatory, ThreadProcParam* pParam, const wchar_t* sModule,
                 const wchar_t* sDll, const wchar_t* sFunction, PVOID pOwnFunction)
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
        storeMessage(bMandatory, pParam, sModuleName, L" is not loaded");
        return false;
    }

    return hook(bMandatory, pParam, hModule, sModuleName, sDll, sFunction, pOwnFunction);
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

        if (!hook(true, pParam, L"msvbvm60.dll", L"ole32.dll", L"CoCreateInstance",
                  myCoCreateInstance))
            return FALSE;

        hook(false, pParam, L"msvbvm60.dll", L"ole32.dll", L"CoCreateInstanceEx",
             myCoCreateInstanceEx);

        if (!hook(true, pParam, L"msvbvm60.dll", L"kernel32.dll", L"GetProcAddress",
                  myGetProcAddress))
            return FALSE;
    }
    else
    {
        // It is some other executable. We must hook LoadLibraryW().

        nHookedFunctions = 0;
        hook(false, pParam, nullptr, L"kernel32.dll", L"GetProcAddress", myGetProcAddress);
        hook(false, pParam, nullptr, L"kernel32.dll", L"LoadLibraryW", myLoadLibraryW);
        hook(false, pParam, nullptr, L"kernel32.dll", L"LoadLibraryExW", myLoadLibraryExW);
        hook(false, pParam, nullptr, L"ole32.dll", L"CoCreateInstance", myCoCreateInstance);
        hook(false, pParam, nullptr, L"ole32.dll", L"CoCreateInstanceEx", myCoCreateInstanceEx);

        if (nHookedFunctions == 0)
        {
            storeMessage(true, pParam, L"Could not hook a single interesting function to hook");
            return FALSE;
        }
    }

    pParam->mbPassedInjectedThread = true;
    return TRUE;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

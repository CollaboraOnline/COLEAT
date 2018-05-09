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

#include "CProxiedClassFactory.hpp"
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
            std::cout << convertUTF16ToUTF8(pParam->msErrorExplanation) << std::endl;
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
        wcscat_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION, pMsg2);
    if (pMsg3)
        wcscat_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION, pMsg3);
    if (pMsg4)
        wcscat_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION, pMsg4);
    if (pMsg5)
        wcscat_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION, pMsg5);
    if (pMsg6)
        wcscat_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION, pMsg6);

    if (bIsError && pParam->mbPassedInjectedThread
        && GetWindowsErrorString(GetLastError(), &pMsgBuf))
    {
        wcscat_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION, L": ");
        wcscat_s(pParam->msErrorExplanation, ThreadProcParam::NEXPLANATION, pMsgBuf);
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

static void* generateTrampoline(void* pFunction, uintptr_t nId, short nArguments)
{
    // We must allocate a fresh page for each trampoline because of multi-thread concerns: Otherwise
    // we would need to change the protection of the page back to RW for a moment when creating
    // another trampoline on it, and if another thread was just executing an existing trampoline,
    // that would be a problem.
    unsigned char* pPage = (unsigned char*)VirtualAlloc(NULL, 100, MEM_COMMIT, PAGE_READWRITE);
    if (pPage == NULL)
    {
        std::cout << "VirtualAlloc failed\n";
        std::exit(1);
    }
    unsigned char* pCode = pPage;

#ifndef _WIN64
    // Normal __stdcall prologue

    // push ebp
    *pCode++ = 0x55;

    // mov evp, esp
    *pCode++ = 0x8B;
    *pCode++ = 0xEC;

    // sub esp, 64
    *pCode++ = 0x83;
    *pCode++ = 0xEC;
    *pCode++ = 0x40;

    // push ebx
    *pCode++ = 0x53;

    // push esi
    *pCode++ = 0x56;

    // push edi
    *pCode++ = 0x57;

    // Push our parameters
    for (short i = 0; i < nArguments; ++i)
    {
        if ((i % 3) == 0)
        {
            // mov eax, dword ptr arg[ebp]
            *pCode++ = 0x8B;
            *pCode++ = 0x45;
            *pCode++ = (unsigned char)(8 + (nArguments - i - 1) * 4);

            // push eax
            *pCode++ = 0x50;
        }
        else if ((i % 3) == 1)
        {
            // mov ecx, dword ptr arg[ebp]
            *pCode++ = 0x8B;
            *pCode++ = 0x4D;
            *pCode++ = (unsigned char)(8 + (nArguments - i - 1) * 4);

            // push ecx
            *pCode++ = 0x51;
        }
        else if ((i % 3) == 2)
        {
            // mov edx, dword ptr arg[ebp]
            *pCode++ = 0x8B;
            *pCode++ = 0x55;
            *pCode++ = (unsigned char)(8 + (nArguments - i - 1) * 4);

            // push edx
            *pCode++ = 0x52;
        }
        else
            abort();
    }

    // call <relative 32-bit offset>
    *pCode++ = 0xE8;
    intptr_t nDiff = ((unsigned char*)pFunction - pCode - 4);
    *pCode++ = ((unsigned char*)&nDiff)[0];
    *pCode++ = ((unsigned char*)&nDiff)[1];
    *pCode++ = ((unsigned char*)&nDiff)[2];
    *pCode++ = ((unsigned char*)&nDiff)[3];

    // Normal __stdcall epilogue

    // pop edi
    *pCode++ = 0x5F;

    // pop esi
    *pCode++ = 0x5E;

    // pop ebx
    *pCode++ = 0x5B;

    // mov esp, ebp
    *pCode++ = 0x8B;
    *pCode++ = 0xE5;

    // pop ebp
    *pCode++ = 0x5D;

    // ret <nArguments*4>
    *pCode++ = 0xC2;
    short n = nArguments * 4;
    *pCode++ = ((unsigned char*)&n)[0];
    *pCode++ = ((unsigned char*)&n)[1];

    // the unique id is stored after the ret <n>
    *pCode++ = ((unsigned char*)&nId)[0];
    *pCode++ = ((unsigned char*)&nId)[1];
    *pCode++ = ((unsigned char*)&nId)[2];
    *pCode++ = ((unsigned char*)&nId)[3];
#else
    // Normal prologue

    if (nArguments > 3)
    {
        // mov qword ptr [rsp+32], r9
        *pCode++ = 0x4C;
        *pCode++ = 0x89;
        *pCode++ = 0x4C;
        *pCode++ = 0x24;
        *pCode++ = 0x20;
    }

    if (nArguments > 2)
    {
        // mov qword ptr [rsp+24], r8
        *pCode++ = 0x4C;
        *pCode++ = 0x89;
        *pCode++ = 0x44;
        *pCode++ = 0x24;
        *pCode++ = 0x18;
    }

    if (nArguments > 1)
    {
        // mov qword ptr [rsp+16], rdx
        *pCode++ = 0x48;
        *pCode++ = 0x89;
        *pCode++ = 0x54;
        *pCode++ = 0x24;
        *pCode++ = 0x10;
    }

    if (nArguments > 0)
    {
        // mov qword ptr [rsp+8], rcx
        *pCode++ = 0x48;
        *pCode++ = 0x89;
        *pCode++ = 0x4C;
        *pCode++ = 0x24;
        *pCode++ = 0x08;
    }

    // push rbp
    *pCode++ = 0x40;
    *pCode++ = 0x55;

    // sub rsp, <x>
    *pCode++ = 0x48;
    *pCode++ = 0x83;
    *pCode++ = 0xEC;
    if (nArguments <= 4)
        *pCode++ = 0x60;
    else
    {
        assert(0x70 + (nArguments - 5) / 2 * 0x10 <= 0xFF);
        *pCode++ = (unsigned char)(0x70 + (nArguments - 5) / 2 * 0x10);
    }

    // lea rbp, qword ptr [rsp+<x>]
    *pCode++ = 0x48;
    *pCode++ = 0x8D;
    *pCode++ = 0x6C;
    *pCode++ = 0x24;
    if (nArguments <= 4)
        *pCode++ = 0x20;
    else
    {
        assert(0x30 + (nArguments - 5) / 2 * 0x10 <= 0xFF);
        *pCode++ = (unsigned char)(0x30 + (nArguments - 5) / 2 * 0x10);
    }

    // Parameters
    for (short i = 0; i < nArguments; ++i)
    {
        if (i == 0)
        {
            // mov rcx, qword ptr arg0[rbp]
            *pCode++ = 0x48;
            *pCode++ = 0x8B;
            *pCode++ = 0x4D;
            *pCode++ = 0x50;
        }
        else if (i == 1)
        {
            // mov rdx, qword ptr arg1[rbp]
            *pCode++ = 0x48;
            *pCode++ = 0x8B;
            *pCode++ = 0x55;
            *pCode++ = 0x58;
        }
        else if (i == 2)
        {
            // mov r8, qword ptr arg2[rbp]
            *pCode++ = 0x4C;
            *pCode++ = 0x8B;
            *pCode++ = 0x45;
            *pCode++ = 0x60;
        }
        else if (i == 3)
        {
            // mov r9, qword ptr arg2[rbp]
            *pCode++ = 0x4C;
            *pCode++ = 0x8B;
            *pCode++ = 0x4D;
            *pCode++ = 0x68;
        }
        else
        {
            if (i <= 5)
            {
                // mov rax, qword ptr (112+(i-4)*8)[rbp]
                *pCode++ = 0x48;
                *pCode++ = 0x8B;
                *pCode++ = 0x45;
                *pCode++ = (unsigned char)(112 + (i - 4) * 8);
            }
            else
            {
                // mov rax, qword ptr (112+(i-4)*8)[rbp]
                *pCode++ = 0x48;
                *pCode++ = 0x8B;
                *pCode++ = 0x85;
                int n = 112 + (i - 4) * 8;
                *pCode++ = ((unsigned char*)&n)[0];
                *pCode++ = ((unsigned char*)&n)[1];
                *pCode++ = ((unsigned char*)&n)[2];
                *pCode++ = ((unsigned char*)&n)[3];
            }

            // mov qword ptr [rsp+32+(i-4)*8], rax
            *pCode++ = 0x48;
            *pCode++ = 0x89;
            *pCode++ = 0x44;
            *pCode++ = 0x24;
            *pCode++ = (unsigned char)(32 + (i - 4) * 8);
        }
    }

    // mov rax, qword ptr pFunction (or something like that)
    *pCode++ = 0x48;
    *pCode++ = 0xB8;
    *pCode++ = ((unsigned char*)&pFunction)[0];
    *pCode++ = ((unsigned char*)&pFunction)[1];
    *pCode++ = ((unsigned char*)&pFunction)[2];
    *pCode++ = ((unsigned char*)&pFunction)[3];
    *pCode++ = ((unsigned char*)&pFunction)[4];
    *pCode++ = ((unsigned char*)&pFunction)[5];
    *pCode++ = ((unsigned char*)&pFunction)[6];
    *pCode++ = ((unsigned char*)&pFunction)[7];

    // call rax
    *pCode++ = 0xFF;
    *pCode++ = 0xD0;

    // Normal epilogue

    // lea rsp, qword ptr [rbp+64]
    *pCode++ = 0x48;
    *pCode++ = 0x8D;
    *pCode++ = 0x65;
    *pCode++ = 0x40;

    // pop rbp
    *pCode++ = 0x5D;

    // ret
    *pCode++ = 0xC3;

    // the unique number is stored after the ret
    *pCode++ = ((unsigned char*)&nId)[0];
    *pCode++ = ((unsigned char*)&nId)[1];
    *pCode++ = ((unsigned char*)&nId)[2];
    *pCode++ = ((unsigned char*)&nId)[3];
    *pCode++ = ((unsigned char*)&nId)[4];
    *pCode++ = ((unsigned char*)&nId)[5];
    *pCode++ = ((unsigned char*)&nId)[6];
    *pCode++ = ((unsigned char*)&nId)[7];
#endif

    DWORD nOldProtection;
    if (!VirtualProtect(pPage, 100, PAGE_EXECUTE, &nOldProtection))
    {
        std::cout << "VirtualProtect failed\n";
        std::exit(1);
    }

    return pPage;
}

static HRESULT WINAPI myCoCreateInstance(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext,
                                         REFIID riid, LPVOID* ppv)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "CoCreateInstance(" << rclsid << ", " << riid << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << "..." << std::endl;

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
                  << prettyCodeAddress(_ReturnAddress()) << "..." << std::endl;
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

static HRESULT __stdcall myCoGetClassObject(REFCLSID rclsid, DWORD dwClsContext,
                                            COSERVERINFO* pServerInfo, REFIID riid, LPVOID* ppv)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "CoGetClassObject(" << rclsid << ", " << riid << ")..." << std::endl;

    HRESULT nRetval = CoGetClassObject(rclsid, dwClsContext, pServerInfo, riid, ppv);

    if (pGlobalParamPtr->mbVerbose)
    {
        std::cout << "...CoGetClassObject(" << rclsid << ", " << riid << "): unhandled: ";
        if (nRetval == S_OK)
            std::cout << *ppv << std::endl;
        else
            std::cout << HRESULT_to_string(nRetval) << std::endl;
    }

    return nRetval;
}

static HRESULT __stdcall myDllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
#ifndef _WIN64
    // 9 = the length of the __stdcall epilogue
    unsigned char* pHModule = (unsigned char*)_ReturnAddress() + 9;
#else
    // 6 = the length of the epilogue
    unsigned char* pHModule = (unsigned char*)_ReturnAddress() + 6;
#endif
    HMODULE hModule;
    std::memmove(&hModule, pHModule, sizeof(HMODULE));

    if (pGlobalParamPtr->mbVerbose)
        std::cout << "DllGetClassObject(" << rclsid << ", " << riid << ") in "
                  << moduleName(hModule) << "..." << std::endl;

    FunPtr aProc;
    aProc.pProc = GetProcAddress(hModule, "DllGetClassObject");
    assert(aProc.pProc != NULL);
    HRESULT nResult = aProc.pDllGetClassObject(rclsid, riid, ppv);

    if (pGlobalParamPtr->mbVerbose)
    {
        std::cout << "...DllGetClassObject(" << rclsid << ", " << riid << "): ";
        if (nResult == S_OK)
            printCreateInstanceResult(*ppv);
        else
            std::cout << HRESULT_to_string(nResult) << std::endl;
    }

    if (nResult == S_OK && IsEqualIID(riid, IID_IClassFactory))
    {
        IClassFactory* pCF;
        if (!FAILED(
                reinterpret_cast<IUnknown*>(*ppv)->QueryInterface(IID_IClassFactory, (void**)&pCF)))
        {
            *ppv = new CProxiedClassFactory(pCF,
                                            _strdup(withoutExtension(moduleName(hModule)).c_str()));
        }
    }

    return nResult;
}

static PROC WINAPI myGetProcAddress(HMODULE hModule, LPCSTR lpProcName)
{
    HMODULE hOle32 = GetModuleHandleW(L"ole32.dll");
    FunPtr pFun;

    if (hModule == hOle32 && std::strcmp(lpProcName, "CoCreateInstanceEx") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(ole32.dll, CoCreateInstanceEx) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myCoCreateInstanceEx;
        return pFun.pProc;
    }

    if (hModule == hOle32 && std::strcmp(lpProcName, "CoCreateInstance") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(ole32.dll, CoCreateInstance) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myCoCreateInstance;
        return pFun.pProc;
    }

    if ((uintptr_t)lpProcName < 10000)
    {
        // It is most likely an ordinal, sigh
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(" << moduleName(hModule) << ", " << (uintptr_t)lpProcName
                      << ") from " << prettyCodeAddress(_ReturnAddress()) << std::endl;

        return GetProcAddress(hModule, lpProcName);
    }

    if (std::strcmp(lpProcName, "DllGetClassObject") == 0)
    {
        PROC pProc = GetProcAddress(hModule, lpProcName);
        if (pProc == NULL)
        {
            if (pGlobalParamPtr->mbVerbose)
                std::cout << "GetProcAddress(" << moduleName(hModule) << ", "
                          << "DllGetClassObject) from " << prettyCodeAddress(_ReturnAddress())
                          << ": NULL" << std::endl;

            return NULL;
        }

        // Interesting case. We must generate a unique trampoline for it.
        FunPtr aTrampoline;
        aTrampoline.pVoid = generateTrampoline(myDllGetClassObject, (uintptr_t)hModule, 3);

        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(" << moduleName(hModule) << ", "
                      << "DllGetClassObject) from " << prettyCodeAddress(_ReturnAddress())
                      << ": Generated trampoline" << std::endl;

        return aTrampoline.pProc;
    }

    PROC pRetval = GetProcAddress(hModule, lpProcName);

    if (pGlobalParamPtr->mbVerbose)
        std::cout << "GetProcAddress(" << moduleName(hModule) << ", " << lpProcName << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << ": unhandled: " << pRetval << std::endl;

    return pRetval;
}

static HMODULE WINAPI myLoadLibraryW(LPCWSTR lpFileName)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "LoadLibraryW(" << convertUTF16ToUTF8(lpFileName) << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << "..." << std::endl;

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
                  << ") from " << prettyCodeAddress(_ReturnAddress()) << "..." << std::endl;

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
// std::cout. They run in a thread created before those have been set up. They can, however, use
// Win32 API directly.

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

        hook(false, pParam, L"msvbvm60.dll", L"ole32.dll", L"CoGetClassObject", myCoGetClassObject);

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
        hook(false, pParam, nullptr, L"ole32.dll", L"CoGetClassObject", myCoGetClassObject);

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

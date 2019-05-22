/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#pragma warning(push)
#pragma warning(disable : 4365 4458 4571 4625 4668 4774 4820 4917 5026 5039)

#include <cassert>
#include <chrono>
#include <cstdio>
#include <cstdlib>
#include <cwchar>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>

#include <intrin.h>
#include <process.h>

#include <Windows.h>
#include <gdiplus.h>
#include <shlwapi.h>

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
#include "CProxiedMoniker.hpp"

#include "InterfaceMapping.hxx"

struct UNICODE_STRING
{
    USHORT Length;
    USHORT MaximumLength;
    PWCH Buffer;
};

static ThreadProcParam* pGlobalParamPtr;

static int nHookedFunctions = 0;

static bool bDidAllocConsole;

static IID aIID_WriterApplication{ 0x82154421, 0x0FBF, 0x11D4, 0x83, 0x13, 0x00,
                                   0x50,       0x04,   0x52,   0x6A, 0xB4 };

static DLL_DIRECTORY_COOKIE(WINAPI* pAddDllDirectory)(PCWSTR) = NULL;
static BOOL(WINAPI* pSetDllDirectoryW)(LPCWSTR) = NULL;
static BOOL(WINAPI* pSetDllDirectoryA)(LPCSTR) = NULL;

static bool hook(bool bMandatory, ThreadProcParam* pParam, HMODULE hModule,
                 const wchar_t* sModuleName, const wchar_t* sDll, const char* sFunction,
                 PVOID pOwnFunction);

static HMODULE WINAPI myLoadLibraryA(LPCSTR lpFileName);
static HMODULE WINAPI myLoadLibraryExW(LPCWSTR lpFileName, HANDLE hFile, DWORD dwFlags);
static HMODULE WINAPI myLoadLibraryExA(LPCSTR lpFileName, HANDLE hFile, DWORD dwFlags);
static NTSTATUS NTAPI myLdrLoadDll(PWCHAR PathToFile, ULONG Flags, UNICODE_STRING* ModuleFileName,
                                   PHANDLE ModuleHandle);

static void printCreateInstanceResult(void* pV)
{
    std::string sName = prettyNameOfType(pGlobalParamPtr, (IUnknown*)pV);

    std::cout << pV;

    if (sName != "")
        std::cout << " (" << sName << ")";

    std::cout << std::endl;
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
            if (pGlobalParamPtr->mbVerbose)
            {
                std::cout << "...CoCreateInstance(" << rclsid << ", " << riid << ")"
                          << " -> ";
                printCreateInstanceResult(*ppv);
            }

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
                              << "): ";
                    if (pResults[j].hr == S_OK)
                    {
                        std::cout << *pResults[j].pIID << ": ";
                        printCreateInstanceResult(pResults[j].pItf);
                    }
                    else
                        std::cout << HRESULT_to_string(pResults[j].hr) << std::endl;
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

    HRESULT nRetval
        = CoCreateInstanceEx(clsid, pUnkOuter, dwClsCtx, pServerInfo, dwCount, pResults);

    if (pGlobalParamPtr->mbVerbose)
    {
        for (DWORD j = 0; j < dwCount; ++j)
        {
            std::cout << "...CoCreateInstanceEx(" << clsid << "): " << j << "(" << dwCount << "): ";
            if (pResults[j].hr == S_OK)
            {
                std::cout << *pResults[j].pIID << ": ";
                printCreateInstanceResult(pResults[j].pItf);
            }
            else
                std::cout << HRESULT_to_string(pResults[j].hr) << std::endl;
        }
    }

    return nRetval;
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

#ifdef HARDCODE_MSO_TO_CO

class myEnumOLEVERB : public IEnumOLEVERB
{
private:
    ULONG mnRefCount;
    ULONG mnIndex;

public:
    myEnumOLEVERB()
        : mnRefCount(1)
        , mnIndex(0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myEnumOLEVERB::CTOR()" << std::endl;
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myEnumOLEVERB::QueryInterface(" << riid << ")" << std::endl;
        if (IsEqualIID(riid, IID_IUnknown))
            *ppvObject = this;
        else if (IsEqualIID(riid, IID_IEnumOLEVERB))
            *ppvObject = this;
        else
            return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        mnRefCount++;
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myEnumOLEVERB::AddRef(): " << mnRefCount << std::endl;
        return mnRefCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        mnRefCount--;
        ULONG nRetval = mnRefCount;
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myEnumOLEVERB::Release(): " << mnRefCount << std::endl;
        if (nRetval == 0)
            delete this;
        return nRetval;
    }

    // IEnumOLEVERB
    HRESULT STDMETHODCALLTYPE Next(ULONG celt, LPOLEVERB rgelt, ULONG* pceltFetched) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myEnumOLEVERB::Next(" << celt << ")" << std::endl;
        if (mnIndex > 1)
            return S_FALSE;
        ULONG nIndexFirst = mnIndex;
        ULONG nLeft = celt;
        while (nLeft > 0)
        {
            if (mnIndex == 0)
            {
                rgelt->lVerb = 0;
                rgelt->lpszVerbName = L"&Edit";
                rgelt->fuFlags = 0;
                rgelt->grfAttribs = OLEVERBATTRIB_ONCONTAINERMENU;
                rgelt++;
                mnIndex++;
            }
            else if (mnIndex == 1)
            {
                rgelt->lVerb = 1;
                rgelt->lpszVerbName = L"&Open";
                rgelt->fuFlags = 0;
                rgelt->grfAttribs = OLEVERBATTRIB_ONCONTAINERMENU;
                rgelt++;
                mnIndex++;
            }
            nLeft--;
        }
        *pceltFetched = mnIndex - nIndexFirst;
        if (*pceltFetched == celt)
            return S_OK;
        else
            return S_FALSE;
    }

    HRESULT STDMETHODCALLTYPE Skip(ULONG celt) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myEnumOLEVERB::Skip(" << celt << ")" << std::endl;
        mnIndex += celt;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Reset() override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myEnumOLEVERB::Reset()" << std::endl;
        mnIndex = 0;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Clone(IEnumOLEVERB** ppenum) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myEnumOLEVERB::Clone()" << std::endl;
        *ppenum = new myEnumOLEVERB();
        return S_OK;
    }
};

class myOleLink : IOleLink
{
private:
    // Pointer to the myOleObject that manages this object
    IUnknown* mpUnk;
    DWORD mnUpdateOptions;
    LPMONIKER mpLinkSrc;
    IMalloc* mpMalloc;
    wchar_t* msDisplayName;

public:
    myOleLink(LPMONIKER pmkLinkSrc)
        : mpUnk(NULL)
        , mnUpdateOptions(OLEUPDATE_ALWAYS)
        , mpLinkSrc(pmkLinkSrc)
        , mpMalloc(NULL)
    {
        HRESULT nResult;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::CTOR(" << pmkLinkSrc << ")" << std::endl;

        nResult = CoGetMalloc(1, &mpMalloc);
        if (nResult != S_OK)
        {
            std::cout << "CoGetMalloc failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            return;
        }

        msDisplayName = (wchar_t*)mpMalloc->Alloc(2);
        msDisplayName[0] = L'\0';

        if (mpLinkSrc)
        {
            mpLinkSrc->AddRef();

            IBindCtx* pBindContext;
            nResult = CreateBindCtx(0, &pBindContext);

            if (nResult != S_OK)
            {
                std::cout << "CreateBindCtx failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
                return;
            }
            mpMalloc->Free(msDisplayName);
            mpLinkSrc->GetDisplayName(pBindContext, NULL, &msDisplayName);

            // FIXME: Or should we keep it around until deleteThis()? Or use the one created over in
            // tryRenderDrawInCollaboraOffice()? What *is* a bind context anyway?
            pBindContext->Release();
        }
    }

    // Can't pass the 'this' of myOleObject when constructing the myOleLink, so have to set it
    // separately, hmm.
    void setUnk(IUnknown* pUnk) { mpUnk = pUnk; }

    void deleteThis()
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::deleteThis()" << std::endl;

        if (mpLinkSrc)
            mpLinkSrc->Release();

        if (mpMalloc)
            mpMalloc->Free(msDisplayName);

        delete this;
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        return mpUnk->QueryInterface(riid, ppvObject);
    }

    ULONG STDMETHODCALLTYPE AddRef() override { return mpUnk->AddRef(); }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG nRetval = mpUnk->Release();
        return nRetval;
    }

    // IOleLink
    HRESULT STDMETHODCALLTYPE SetUpdateOptions(DWORD dwUpdateOpt) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::SetUpdateOptions(" << std::hex << dwUpdateOpt << ")"
                      << std::endl;
        mnUpdateOptions = dwUpdateOpt;

        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE GetUpdateOptions(DWORD* pdwUpdateOpt) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::GetUpdateOptions(): " << std::hex << mnUpdateOptions
                      << std::endl;
        *pdwUpdateOpt = mnUpdateOptions;

        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE SetSourceMoniker(IMoniker* pmk, REFCLSID rclsid) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::SetSourceMoniker(" << pmk << "," << rclsid << ")"
                      << std::endl;
        if (mpLinkSrc)
            mpLinkSrc->Release();
        mpLinkSrc = pmk;
        if (mpLinkSrc)
            mpLinkSrc->AddRef();

        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE GetSourceMoniker(IMoniker** ppmk) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::GetSourceMoniker()" << std::endl;

        *ppmk = mpLinkSrc;

        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE SetSourceDisplayName(LPCOLESTR pszStatusText) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::SetSourceDisplayName(" << pszStatusText << ")"
                      << std::endl;

        if (mpMalloc)
        {
            const size_t nBytes = (wcslen(pszStatusText) + 1) * 2;
            msDisplayName = (wchar_t*)mpMalloc->Alloc(nBytes);
            memcpy(msDisplayName, pszStatusText, nBytes);
        }

        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE GetSourceDisplayName(LPOLESTR* ppszDisplayName) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::GetSourceDisplayName()" << std::endl;

        if (!mpMalloc)
            return E_NOTIMPL;

        const size_t nBytes = (wcslen(msDisplayName) + 1) * 2;
        *ppszDisplayName = (LPOLESTR)mpMalloc->Alloc(nBytes);
        memcpy(*ppszDisplayName, msDisplayName, nBytes);

        if (pGlobalParamPtr->mbVerbose)
            std::cout << "..." << this << "@myOleLink::GetSourceDisplayName(): '" << convertUTF16ToUTF8(*ppszDisplayName) << "'" << std::endl;

        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE BindToSource(DWORD bindflags, IBindCtx* pbc) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::BindToSource(" << std::hex << bindflags << "," << pbc
                      << ")" << std::endl;

        return E_NOTIMPL;
    }

    virtual HRESULT STDMETHODCALLTYPE BindIfRunning() override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::BindIfRunning()" << std::endl;

        return E_NOTIMPL;
    }

    virtual HRESULT STDMETHODCALLTYPE GetBoundSource(IUnknown** ppunk) override
    {
        (void)ppunk;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::GetBoundSource()" << std::endl;

        return E_NOTIMPL;
    }

    virtual HRESULT STDMETHODCALLTYPE UnbindSource() override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::UnbindSource()" << std::endl;

        return E_NOTIMPL;
    }

    virtual HRESULT STDMETHODCALLTYPE Update(IBindCtx* pbc) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleLink::Update(" << pbc << ")" << std::endl;

        return E_NOTIMPL;
    }
};

class myDataObject : IDataObject
{
private:
    // Pointer to the myOleObject that manages this object
    IUnknown* mpUnk;
    std::vector<IAdviseSink*>* mpAdvises;

public:
    myDataObject()
        : mpUnk(NULL)
        , mpAdvises(new std::vector<IAdviseSink*>)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::CTOR()" << std::endl;
    }

    // Can't pass the 'this' of myOleObject when constructing the myDataObject, so have to set it
    // separately, hmm.
    void setUnk(IUnknown* pUnk) { mpUnk = pUnk; }

    void deleteThis()
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::deleteThis()" << std::endl;

        delete mpAdvises;
        delete this;
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        return mpUnk->QueryInterface(riid, ppvObject);
    }

    ULONG STDMETHODCALLTYPE AddRef() override { return mpUnk->AddRef(); }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG nRetval = mpUnk->Release();
        return nRetval;
    }

    // IDataObject
    HRESULT STDMETHODCALLTYPE GetData(FORMATETC* pformatetcIn, STGMEDIUM* pmedium) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::GetData(" << pformatetcIn << "," << pmedium << ")"
                      << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC* pformatetc, STGMEDIUM* pmedium) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::GetDataHere(" << pformatetc << "," << pmedium
                      << ")" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* pformatetc) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::QueryGetData(" << pformatetc << ")" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC* pformatetcIn,
                                                    FORMATETC* pformatetcOut) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::GetCanonicalFormatEtc(" << pformatetcIn << ","
                      << pformatetcOut << ")" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE SetData(FORMATETC* pformatetc, STGMEDIUM* pmedium,
                                      BOOL fRelease) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::SetData(" << pformatetc << "," << pmedium
                      << (fRelease ? "YES" : "NO") << ")" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD dwDirection,
                                            IEnumFORMATETC** ppenumFormatEtc) override
    {
        (void)dwDirection, ppenumFormatEtc;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::EnumFormatEtc()" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC* pformatetc, DWORD advf, IAdviseSink* pAdvSink,
                                      DWORD* pdwConnection) override
    {
        mpAdvises->push_back(pAdvSink);
        *pdwConnection = (DWORD) mpAdvises->size();

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::DAdvise(" << pformatetc << "," << std::hex << advf
                      << "," << pAdvSink << "): " << *pdwConnection << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DUnadvise(DWORD dwConnection) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::DUnadvise(" << dwConnection << ")" << std::endl;

        if (dwConnection == 0 || dwConnection > mpAdvises->size()
            || (*mpAdvises)[dwConnection - 1] == nullptr)
            return OLE_E_NOCONNECTION;

        (*mpAdvises)[dwConnection - 1] = nullptr;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA** ppenumAdvise) override
    {
        (void)ppenumAdvise;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myDataObject::EnumDAdvise()" << std::endl;

        return E_NOTIMPL;
    }
};

class myViewObject : IViewObject
{
private:
    // Pointer to the myOleObject that manages this object
    IUnknown* mpUnk;

    HBITMAP mhBitmap;
    DWORD mnAdviseAspects;
    DWORD mnAdviseFlags;
    IAdviseSink* mpAdviseSink;

public:
    myViewObject(HBITMAP hBitmap)
        : mpUnk(NULL)
        , mhBitmap(hBitmap)
        , mnAdviseAspects(0)
        , mnAdviseFlags(0)
        , mpAdviseSink(NULL)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myViewObject::CTOR()" << std::endl;
    }

    // Can't pass the 'this' of myOleObject when constructing the myViewObject, so have to set it
    // separately, hmm.
    void setUnk(IUnknown* pUnk) { mpUnk = pUnk; }

    void deleteThis()
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myViewObject::deleteThis()" << std::endl;

        DeleteObject(mhBitmap);
        delete this;
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        return mpUnk->QueryInterface(riid, ppvObject);
    }

    ULONG STDMETHODCALLTYPE AddRef() override { return mpUnk->AddRef(); }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG nRetval = mpUnk->Release();
        return nRetval;
    }

    // IViewObject
    HRESULT STDMETHODCALLTYPE Draw(DWORD dwDrawAspect, LONG lindex, void* pvAspect,
                                   DVTARGETDEVICE* ptd, HDC hdcTargetDev, HDC hdcDraw,
                                   LPCRECTL lprcBounds, LPCRECTL lprcWBounds,
                                   BOOL(STDMETHODCALLTYPE* pfnContinue)(ULONG_PTR dwContinue),
                                   ULONG_PTR dwContinue) override
    {
        (void)dwDrawAspect, lindex, pvAspect, ptd, hdcTargetDev, lprcWBounds, pfnContinue,
            dwContinue;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myViewObject::Draw()" << std::endl;

        HDC hdcMem = CreateCompatibleDC(hdcDraw);
        HGDIOBJ hBitmapOld = SelectObject(hdcMem, mhBitmap);

        BITMAP aBitmap;
        GetObject(mhBitmap, sizeof(aBitmap), &aBitmap);

        BitBlt(hdcDraw, lprcBounds->left, lprcBounds->top, aBitmap.bmWidth, aBitmap.bmHeight,
               hdcMem, 0, 0, SRCCOPY);

        SelectObject(hdcMem, hBitmapOld);
        DeleteDC(hdcMem);

#if 0
        // Debugging code. Just of interest, check these in the debugger.
        int nW = GetDeviceCaps(hdcDraw, HORZRES);
        int nH = GetDeviceCaps(hdcDraw, VERTRES);

        (void) nW, nH;

        // Draw some trivial line graphics.
        SelectObject(hdcDraw, GetStockObject(BLACK_PEN));
        MoveToEx(hdcDraw, lprcBounds->left, lprcBounds->top, NULL);
        LineTo(hdcDraw, lprcBounds->right, lprcBounds->bottom);
        MoveToEx(hdcDraw, lprcBounds->left, lprcBounds->bottom, NULL);
        LineTo(hdcDraw, lprcBounds->right, lprcBounds->top);
        MoveToEx(hdcDraw, lprcBounds->left + 5, lprcBounds->top + 5, NULL);
        LineTo(hdcDraw, lprcBounds->right - 5, lprcBounds->top + 5);
        LineTo(hdcDraw, lprcBounds->right - 5, lprcBounds->bottom - 5);
        LineTo(hdcDraw, lprcBounds->left + 5, lprcBounds->bottom - 5);
        LineTo(hdcDraw, lprcBounds->left + 5, lprcBounds->top + 5);
#endif

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetColorSet(DWORD dwDrawAspect, LONG lindex, void* pvAspect,
                                          DVTARGETDEVICE* ptd, HDC hicTargetDev,
                                          LOGPALETTE** ppColorSet) override
    {
        (void)dwDrawAspect, lindex, pvAspect, ptd, hicTargetDev, ppColorSet;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myViewObject::GetColorSet()" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE Freeze(DWORD dwDrawAspect, LONG lindex, void* pvAspect,
                                     DWORD* pdwFreeze) override
    {
        (void)dwDrawAspect, lindex, pvAspect, pdwFreeze;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myViewObject::Freeze()" << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Unfreeze(DWORD dwFreeze) override
    {
        (void)dwFreeze;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myViewObject::Unfreeze()" << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetAdvise(DWORD aspects, DWORD advf, IAdviseSink* pAdvSink) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myViewObject::SetAdvise()" << std::endl;

        mnAdviseAspects = aspects;
        mnAdviseFlags = advf;
        mpAdviseSink = pAdvSink;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetAdvise(DWORD* pAspects, DWORD* pAdvf,
                                        IAdviseSink** ppAdvSink) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myViewObject::GetAdvise()" << std::endl;

        if (pAspects)
            *pAspects = mnAdviseAspects;
        if (pAdvf)
            *pAdvf = mnAdviseFlags;
        if (ppAdvSink)
            *ppAdvSink = mpAdviseSink;

        return S_OK;
    }
};

class myRunnableObject : IRunnableObject
{
private:
    // Pointer to the myOleObject that manages this object
    IUnknown* mpUnk;

public:
    myRunnableObject()
        : mpUnk(NULL)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myRunnableObject::CTOR()" << std::endl;
    }

    // Can't pass the 'this' of myOleObject when constructing the myViewObject, so have to set it
    // separately, hmm.
    void setUnk(IUnknown* pUnk) { mpUnk = pUnk; }

    void deleteThis()
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myRunnableObject::deleteThis()" << std::endl;

        delete this;
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        return mpUnk->QueryInterface(riid, ppvObject);
    }

    ULONG STDMETHODCALLTYPE AddRef() override { return mpUnk->AddRef(); }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG nRetval = mpUnk->Release();

        return nRetval;
    }

    // IRunnableObject
    HRESULT STDMETHODCALLTYPE GetRunningClass(LPCLSID lpClsid) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myRunnableObject::GetRunningClass()" << std::endl;

        *lpClsid = aIID_WriterApplication;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Run(LPBINDCTX pbc) override
    {
        (void)pbc;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myRunnableObject::Run()" << std::endl;

        return S_OK;
    }

    BOOL STDMETHODCALLTYPE IsRunning() override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myRunnableObject::IsRunning()" << std::endl;

        return TRUE;
    }

    HRESULT STDMETHODCALLTYPE LockRunning(BOOL fLock, BOOL fLastUnlockCloses) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myRunnableObject::LockRunning(" << (fLock ? "YES" : "NO") << ","
                      << (fLastUnlockCloses ? "YES" : "NO") << ")" << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetContainedObject(BOOL fContained) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myRunnableObject::SetContainedObject("
                      << (fContained ? "YES" : "NO") << ")" << std::endl;

        return S_OK;
    }
};

class myOleObject : public IOleObject
{
private:
    ULONG mnRefCount;
    SIZEL maExtent;
    std::wstring* mpDocumentPathname;
    myOleLink* mpOleLink;
    myDataObject* mpDataObject;
    myViewObject* mpViewObject;
    myRunnableObject* mpRunnableObject;
    std::vector<IAdviseSink*>* mpAdvises;

public:
    myOleObject(LPMONIKER pmkLinkSrc, HBITMAP hBitmap, const std::wstring& sDocumentPathname)
        : mnRefCount(1)
        , mpDocumentPathname(new std::wstring(sDocumentPathname))
        , mpOleLink(new myOleLink(pmkLinkSrc))
        , mpDataObject(new myDataObject())
        , mpViewObject(new myViewObject(hBitmap))
        , mpRunnableObject(new myRunnableObject())
        , mpAdvises(new std::vector<IAdviseSink*>)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::CTOR()" << std::endl;

        mpOleLink->setUnk(this);
        mpDataObject->setUnk(this);
        mpViewObject->setUnk(this);
        mpRunnableObject->setUnk(this);
        BITMAP aBitmap;
        GetObject(hBitmap, sizeof(aBitmap), &aBitmap);
        maExtent.cx = aBitmap.bmWidth;
        maExtent.cy = aBitmap.bmHeight;
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::QueryInterface(" << riid << ")" << std::endl;
        if (IsEqualIID(riid, IID_IUnknown))
            *ppvObject = this;
        else if (IsEqualIID(riid, IID_IOleObject))
            *ppvObject = this;
        else if (IsEqualIID(riid, IID_IOleLink))
            *ppvObject = mpOleLink;
        else if (IsEqualIID(riid, IID_IDataObject))
            *ppvObject = mpDataObject;
        else if (IsEqualIID(riid, IID_IViewObject))
            *ppvObject = mpViewObject;
        else if (IsEqualIID(riid, IID_IRunnableObject))
            *ppvObject = mpRunnableObject;
        else
            return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        mnRefCount++;
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::AddRef(): " << mnRefCount << std::endl;
        return mnRefCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        mnRefCount--;
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::Release(): " << mnRefCount << std::endl;
        ULONG nRetval = mnRefCount;
        if (nRetval == 0)
        {
            mpOleLink->deleteThis();
            mpDataObject->deleteThis();
            mpViewObject->deleteThis();
            mpRunnableObject->deleteThis();
            delete mpDocumentPathname;
            delete mpAdvises;
            delete this;
        }
        return nRetval;
    }

    // IOleObject
    HRESULT STDMETHODCALLTYPE SetClientSite(IOleClientSite* pClientSite) override
    {
        (void)pClientSite;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::SetClientSize()" << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetClientSite(IOleClientSite** ppClientSite) override
    {
        (void)ppClientSite;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::GetClientSize()" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE SetHostNames(LPCOLESTR szContainerApp,
                                           LPCOLESTR szContainerObj) override
    {
        (void)szContainerApp, szContainerObj;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::SetHostNames()" << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Close(DWORD dwSaveOption) override
    {
        (void)dwSaveOption;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::Close(" << dwSaveOption << ")" << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetMoniker(DWORD dwWhichMoniker, IMoniker* pmk) override
    {
        (void)dwWhichMoniker, pmk;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::SetMoniker()" << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker,
                                         IMoniker** ppmk) override
    {
        (void)dwAssign, dwWhichMoniker, ppmk;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::GetMoniker()" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE InitFromData(IDataObject* pDataObject, BOOL fCreation,
                                           DWORD dwReserved) override
    {
        (void)pDataObject, fCreation, dwReserved;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::InitFromData()" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE GetClipboardData(DWORD dwReserved,
                                               IDataObject** ppDataObject) override
    {
        (void)dwReserved, ppDataObject;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::GetClipboardData()" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE DoVerb(LONG iVerb, LPMSG lpmsg, IOleClientSite* pActiveSite,
                                     LONG lindex, HWND hwndParent, LPCRECT lprcPosRect) override
    {
        (void)iVerb, lpmsg, pActiveSite, lindex, hwndParent, lprcPosRect;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::DoVerb(" << iVerb << ")" << std::endl;

        // We have two verbs, Edit(0) and Open(1), but both do the same thing.

        if (iVerb != 0 && iVerb != 1)
            return E_NOTIMPL;

        // Create a Collabora Office Writer.Application editing the document

        IUnknown* pWriter;
        HRESULT nResult;

        nResult = CoCreateInstance(aIID_WriterApplication, NULL, CLSCTX_LOCAL_SERVER, IID_IUnknown,
                                   (LPVOID*)&pWriter);
        if (nResult != S_OK)
        {
            std::cout << "Could not create Writer.Application object: "
                      << WindowsErrorStringFromHRESULT(nResult) << "\n";
            return nResult;
        }

        IDispatch* pApplication;
        nResult = pWriter->QueryInterface(IID_IDispatch, (void**)&pApplication);
        if (nResult != S_OK)
        {
            std::cout << "Could not get IDispatch from Writer.Application object: "
                      << WindowsErrorStringFromHRESULT(nResult) << "\n";
            pWriter->Release();
            return nResult;
        }

        wchar_t* sDocuments = L"Documents";
        DISPID nDocuments;
        nResult = pApplication->GetIDsOfNames(IID_NULL, &sDocuments, 1, GetUserDefaultLCID(),
                                              &nDocuments);
        if (nResult != S_OK)
        {
            std::cout << "Could not get DISPID of 'Documents' from Writer.Application object: "
                      << WindowsErrorStringFromHRESULT(nResult) << "\n";
            pApplication->Release();
            pWriter->Release();
            return nResult;
        }

        DISPPARAMS aDocumentsArguments[] = { NULL, NULL, 0, 0 };
        VARIANT aDocumentsResult;
        pApplication->Invoke(nDocuments, IID_NULL, GetUserDefaultLCID(), DISPATCH_METHOD,
                             aDocumentsArguments, &aDocumentsResult, NULL, NULL);
        if (nResult != S_OK)
        {
            std::cout << "Could not create invoke 'Documents' of Writer.Application object: "
                      << WindowsErrorStringFromHRESULT(nResult) << "\n";
            pApplication->Release();
            pWriter->Release();
            return nResult;
        }

        if (aDocumentsResult.vt != VT_DISPATCH)
        {
            std::cout << "The 'Documents' function of Writer.Application object did not return an "
                         "IDispatch object\n";
            pApplication->Release();
            pWriter->Release();
            return E_NOTIMPL;
        }

        wchar_t* sOpen = L"Open";
        DISPID nOpen;
        nResult = aDocumentsResult.pdispVal->GetIDsOfNames(IID_NULL, &sOpen, 1,
                                                           GetUserDefaultLCID(), &nOpen);
        if (nResult != S_OK)
        {
            std::cout << "Could not create get DISPID of 'Open' from Writer.Documents object: "
                      << WindowsErrorStringFromHRESULT(nResult) << "\n";
            aDocumentsResult.pdispVal->Release();
            pApplication->Release();
            pWriter->Release();
            return E_NOTIMPL;
        }

        VARIANTARG aFileName;
        aFileName.vt = VT_BSTR;
        aFileName.bstrVal = SysAllocString(mpDocumentPathname->data());
        DISPPARAMS aOpenArguments[] = { &aFileName, NULL, 1, 0 };
        VARIANT aOpenResult;
        aDocumentsResult.pdispVal->Invoke(nOpen, IID_NULL, GetUserDefaultLCID(), DISPATCH_METHOD,
                                          aOpenArguments, &aOpenResult, NULL, NULL);
        if (nResult != S_OK)
        {
            std::cout << "Could not invoke 'Open' of Writer.Documents object: "
                      << WindowsErrorStringFromHRESULT(nResult) << "\n";
            aDocumentsResult.pdispVal->Release();
            pApplication->Release();
            pWriter->Release();
            return nResult;
        }

        aDocumentsResult.pdispVal->Release();
        pApplication->Release();
        pWriter->Release();

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** ppEnumOleVerb) override
    {
        (void)ppEnumOleVerb;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::EnumVerbs()" << std::endl;

        *ppEnumOleVerb = new myEnumOLEVERB();

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Update(void) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::Update()" << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE IsUpToDate(void) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::IsUpToDate()" << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetUserClassID(CLSID* pClsid) override
    {
        (void)pClsid;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::GetUserClassID()" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE GetUserType(DWORD dwFormOfType, LPOLESTR* pszUserType) override
    {
        (void)dwFormOfType, pszUserType;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::GetUserType()" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE SetExtent(DWORD dwDrawAspect, SIZEL* psizel) override
    {
        (void)dwDrawAspect;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::SetExtent()" << std::endl;

        maExtent = *psizel;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetExtent(DWORD dwDrawAspect, SIZEL* psizel) override
    {
        (void)dwDrawAspect;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::GetExtent()" << std::endl;

        *psizel = maExtent;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Advise(IAdviseSink* pAdvSink, DWORD* pdwConnection) override
    {
        mpAdvises->push_back(pAdvSink);
        *pdwConnection = (DWORD) mpAdvises->size();

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::Advise(): " << *pdwConnection << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Unadvise(DWORD dwConnection) override
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::Unadvise(" << dwConnection << ")" << std::endl;

        if (dwConnection == 0 || dwConnection > mpAdvises->size()
            || (*mpAdvises)[dwConnection - 1] == nullptr)
            return OLE_E_NOCONNECTION;

        (*mpAdvises)[dwConnection - 1] = nullptr;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE EnumAdvise(IEnumSTATDATA** ppenumAdvise) override
    {
        (void)ppenumAdvise;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::EnumAdvise()" << std::endl;

        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE GetMiscStatus(DWORD dwAspect, DWORD* pdwStatus) override
    {
        (void)dwAspect;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::GetMiscStatus(" << dwAspect << ")" << std::endl;

        *pdwStatus = OLEMISC_STATIC;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetColorScheme(LOGPALETTE* pLogpal) override
    {
        (void)pLogpal;

        if (pGlobalParamPtr->mbVerbose)
            std::cout << this << "@myOleObject::SetColorScheme()" << std::endl;

        return S_OK;
    }
};

static HRESULT tryRenderDrawInCollaboraOffice(LPMONIKER pmkLinkSrc, REFIID riid, DWORD renderopt,
                                              LPFORMATETC lpFormatEtc, LPOLECLIENTSITE pClientSite,
                                              LPSTORAGE pStg, LPVOID* ppvObj)
{
    // Sanity check. Only attempt for the case we are interested in.
    if (!IsEqualIID(riid, IID_IOleObject) || renderopt != OLERENDER_DRAW || lpFormatEtc != NULL
        || pClientSite == NULL || pStg == NULL)
    {
        std::cout << "Not the kind of call we want to replace\n";
        return S_FALSE;
    }

    HRESULT nResult;

    // Also check that it is a file moniker.
    DWORD dwMksys;
    nResult = pmkLinkSrc->IsSystemMoniker(&dwMksys);
    if (nResult != S_OK || dwMksys != MKSYS_FILEMONIKER)
    {
        std::cout << "Not a file moniker\n";
        return S_FALSE;
    }

    IMalloc* pMalloc;

    nResult = CoGetMalloc(1, &pMalloc);
    if (nResult != S_OK)
    {
        std::cout << "CoGetMalloc failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
        return S_FALSE;
    }

    // We need a bit context. Sadly apparently there is no way to get the one that the calling code
    // must have used to create the pmkLinkSrc IMoniker?
    IBindCtx* pBindContext;
    nResult = CreateBindCtx(0, &pBindContext);

    if (nResult != S_OK)
    {
        std::cout << "CreateBindCtx failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
        return S_FALSE;
    }

    // The display name of the IMoniker is the file pathname.
    wchar_t* sDisplayName;
    nResult = pmkLinkSrc->GetDisplayName(pBindContext, NULL, &sDisplayName);

    if (nResult != S_OK)
    {
        std::cout << "IMoniker::GetDisplayName failed: " << WindowsErrorStringFromHRESULT(nResult)
                  << "\n";
        pBindContext->Release();
        return S_FALSE;
    }

    // Check if the display name is a pathname of a file with an extension we "know" that Collabora
    // Office handles. Only check a few extensions, and in the actual customer case, it is .rtf.
    wchar_t* pBasenameEnd = wcsrchr(sDisplayName, L'.');
    if (pBasenameEnd == nullptr)
    {
        // No file name extension, let's not even try.
        std::cout << "No file name extension in '" << convertUTF16ToUTF8(sDisplayName) << "'\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    const wchar_t* pLastSlash = wcsrchr(sDisplayName, L'\\');
    const std::wstring sBasename(pLastSlash + 1,
                                 static_cast<std::size_t>(pBasenameEnd - pLastSlash - 1));
    const std::wstring sExtension(pBasenameEnd + 1);

    if (!(sExtension == L"rtf" || sExtension == L"doc" || sExtension == L"docx"
          || sExtension == L"odt"))
    {
        std::cout << "Not a known file name extension in '" << convertUTF16ToUTF8(sDisplayName)
                  << "'\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

#if 0
    // Debugging code. Get the display name of the client side's moniker, if there, just out of interest.

    IMoniker *pClientSiteMoniker;
    nResult = pClientSite->GetMoniker(OLEGETMONIKER_ONLYIFTHERE, OLEWHICHMK_OBJFULL, &pClientSiteMoniker);

    if (nResult == S_OK)
    {
        wchar_t *sClientSiteDisplayName;
        nResult = pClientSiteMoniker->GetDisplayName(pBindContext, NULL, &sClientSiteDisplayName);

        if (nResult != S_OK)
        {
            pMalloc->Free(sDisplayName);
            pBindContext->Release();
            return S_FALSE;
        }

        pMalloc->Free(sClientSiteDisplayName);
    }
#endif

    // Use Collabora Office to create a png from the document.

    wchar_t sTempPath[MAX_PATH + 1];
    if (GetTempPathW(MAX_PATH + 1, sTempPath) == 0)
    {
        std::cout << "GetTempPathW failed!\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }
    if (sTempPath[wcslen(sTempPath) - 1] == L'\\')
        sTempPath[wcslen(sTempPath) - 1] = L'\0';

    // Where is CO installed?

    HKEY hUno;
    nResult = RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\LibreOffice\\UNO\\InstallPath", 0,
                            KEY_READ | KEY_WOW64_32KEY, &hUno);
    if (nResult == ERROR_FILE_NOT_FOUND)
        nResult = RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\LibreOffice\\UNO\\InstallPath", 0,
                                KEY_READ | KEY_WOW64_64KEY, &hUno);

    if (nResult != ERROR_SUCCESS)
    {
        std::cout << "Can not find Collabora Office installation, RegOpenKeyExW failed: "
                  << WindowsErrorStringFromHRESULT(nResult) << "\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    wchar_t sLOPath[MAX_PATH + 1];
    LONG nLOPathSize = sizeof(sLOPath);
    nResult = RegQueryValueW(hUno, NULL, sLOPath, &nLOPathSize);
    if (nResult != ERROR_SUCCESS)
    {
        std::cout << "Can not find Collabora Office installation, RegQueryValueW failed: "
                  << WindowsErrorStringFromHRESULT(nResult) << "\n";
        RegCloseKey(hUno);
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    RegCloseKey(hUno);

    // Create a fresh directory to use as the UserInstallation. We don't want this run of Collabora
    // Office to interfere with any potentially already running Collabora Office.
    std::error_code aError;
    std::clock_t n = std::clock();
    std::wstring sUserInstallation;
    std::experimental::filesystem::path aUserInstallation;
    while (true)
    {
        sUserInstallation = std::wstring(sTempPath) + L"\\coleat-convert-to." + std::to_wstring(n);
        aUserInstallation = std::experimental::filesystem::path(sUserInstallation);
        if (std::experimental::filesystem::create_directory(aUserInstallation, aError))
            break;
        n++;
    }

    wchar_t sUrl[200];
    DWORD nUrl = sizeof(sUrl) / sizeof(sUrl[0]);
    if (UrlCreateFromPathW(sUserInstallation.data(), sUrl, &nUrl, NULL) != S_OK)
    {
        std::cout << "Can not turn '" << convertUTF16ToUTF8(sUserInstallation.data())
                  << "' into a URL: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    std::wstring sCommandLine = L"\"" + std::wstring(sLOPath) + L"\\soffice.exe\""
                                + L" -env:UserInstallation=" + std::wstring(sUrl)
                                + L" --convert-to png --outdir \"" + sUserInstallation + L"\" \""
                                + std::wstring(sDisplayName) + L"\"";

    STARTUPINFOW aStartupInfo;
    PROCESS_INFORMATION aProcessInformation;

    memset(&aStartupInfo, 0, sizeof(aStartupInfo));
    aStartupInfo.cb = sizeof(aStartupInfo);

    wchar_t* pCommandLine = _wcsdup(sCommandLine.data());
    nResult = CreateProcessW(NULL, pCommandLine, NULL, NULL, TRUE, 0, NULL, NULL, &aStartupInfo,
                             &aProcessInformation);
    free(pCommandLine);

    if (nResult == 0)
    {
        std::cout << "Can not run soffice, CreateProcessW failed: "
                  << WindowsErrorStringFromHRESULT(nResult) << "\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    WaitForSingleObject(aProcessInformation.hProcess, INFINITE);

    std::wstring sImageFile = sUserInstallation + L"\\" + sBasename + L".png";

    Gdiplus::GdiplusStartupInput gdiplusStartupInput;
    ULONG_PTR gdiplusToken;
    Gdiplus::GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

    Gdiplus::Bitmap* pBitmap = Gdiplus::Bitmap::FromFile(sImageFile.data());

    HBITMAP hBitmap;
    pBitmap->GetHBITMAP(Gdiplus::Color(), &hBitmap);
    delete pBitmap;

    Gdiplus::GdiplusShutdown(gdiplusToken);

    std::experimental::filesystem::remove_all(aUserInstallation);

    *ppvObj = new myOleObject(pmkLinkSrc, hBitmap, sDisplayName);

    pMalloc->Free(sDisplayName);
    pBindContext->Release();

    return S_OK;
}

#endif

static HRESULT WINAPI myOleCreateLink(LPMONIKER pmkLinkSrc, REFIID riid, DWORD renderopt,
                                      LPFORMATETC lpFormatEtc, LPOLECLIENTSITE pClientSite,
                                      LPSTORAGE pStg, LPVOID* ppvObj)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "OleCreateLink(" << pmkLinkSrc << "," << riid << "," << renderopt << ","
                  << lpFormatEtc << "," << pClientSite << "," << pStg << "," << ppvObj << ")..."
                  << std::endl;

    HRESULT nRetval;

#ifdef HARDCODE_MSO_TO_CO
    nRetval = tryRenderDrawInCollaboraOffice(pmkLinkSrc, riid, renderopt, lpFormatEtc, pClientSite,
                                             pStg, ppvObj);

    if (nRetval == S_OK)
    {
        if (pGlobalParamPtr->mbVerbose)
        {
            std::cout << "...OleCreateLink(" << pmkLinkSrc << "," << riid
                      << ",...): " << WindowsErrorStringFromHRESULT(nRetval) << std::endl;
        }
        return S_OK;
    }
#endif

    nRetval = OleCreateLink(pmkLinkSrc, riid, renderopt, lpFormatEtc, pClientSite, pStg, ppvObj);

    if (pGlobalParamPtr->mbVerbose)
    {
        std::cout << "...OleCreateLink(" << pmkLinkSrc << "," << riid
                  << ",...): " << WindowsErrorStringFromHRESULT(nRetval) << std::endl;
    }

    return nRetval;
}

static void myOutputDebugStringA(char* lpOutputString)
{
    if (pGlobalParamPtr->mbVerbose)
    {
        std::wstring s = convertACPToUTF16(lpOutputString);
        if (s.length() > 1 && s[s.length() - 1] == L'\n')
        {
            s[s.length() - 1] = L'\0';
            if (s[s.length() - 2] == L'\r')
                s[s.length() - 2] = L'\0';
        }

        std::cout << "OutputDebugStringA(" << convertUTF16ToUTF8(s.data()) << ")" << std::endl;
    }
    OutputDebugStringA(lpOutputString);
}

static void myOutputDebugStringW(wchar_t* lpOutputString)
{
    if (pGlobalParamPtr->mbVerbose)
    {
        std::wstring s = std::wstring(lpOutputString);
        if (s.length() > 1 && s[s.length() - 1] == L'\n')
        {
            s[s.length() - 1] = L'\0';
            if (s[s.length() - 2] == L'\r')
                s[s.length() - 2] = L'\0';
        }

        std::cout << "OutputDebugStringW(" << convertUTF16ToUTF8(s.data()) << ")" << std::endl;
    }
    OutputDebugStringW(lpOutputString);
}

static HINSTANCE myShellExecuteA(HWND hwnd, LPCSTR lpOperation, LPCSTR lpFile, LPCSTR lpParameters,
                                 LPCSTR lpDirectory, INT nShowCmd)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "ShellExecuteA(" << hwnd << "," << (lpOperation ? lpOperation : "NULL") << ","
                  << convertUTF16ToUTF8(convertACPToUTF16(lpFile).data()) << ","
                  << (lpParameters ? convertUTF16ToUTF8(convertACPToUTF16(lpParameters).data())
                                   : "NULL")
                  << ","
                  << (lpDirectory ? convertUTF16ToUTF8(convertACPToUTF16(lpDirectory).data())
                                  : "NULL")
                  << "," << nShowCmd << ")..." << std::endl;

    HINSTANCE hInstance
        = ShellExecuteA(hwnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd);
    DWORD nLastError = GetLastError();

    if (pGlobalParamPtr->mbVerbose)
    {
        std::cout << "...ShellExecuteA(): " << (intptr_t)hInstance;
        if ((intptr_t)hInstance <= 32)
            std::cout << ": " << WindowsErrorString(nLastError);

        std::cout << std::endl;
    }

    SetLastError(nLastError);

    return hInstance;
}

static HINSTANCE myShellExecuteW(HWND hwnd, LPCWSTR lpOperation, LPCWSTR lpFile,
                                 LPCWSTR lpParameters, LPCWSTR lpDirectory, INT nShowCmd)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "ShellExecuteW(" << hwnd << ","
                  << (lpOperation ? convertUTF16ToUTF8(lpOperation) : "NULL") << ","
                  << convertUTF16ToUTF8(lpFile) << ","
                  << (lpParameters ? convertUTF16ToUTF8(lpParameters) : "NULL") << ","
                  << (lpDirectory ? convertUTF16ToUTF8(lpDirectory) : "NULL") << "," << nShowCmd
                  << ")..." << std::endl;

    HINSTANCE hInstance
        = ShellExecuteW(hwnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd);
    DWORD nLastError = GetLastError();

    if (pGlobalParamPtr->mbVerbose)
    {
        std::cout << "...ShellExecuteW(): " << (intptr_t)hInstance;
        if ((intptr_t)hInstance <= 32)
            std::cout << ": " << WindowsErrorString(nLastError);

        std::cout << std::endl;
    }

    SetLastError(nLastError);

    return hInstance;
}

static BOOL myShellExecuteExA(SHELLEXECUTEINFOA* pExecInfo)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "ShellExecuteExA({hwnd=" << pExecInfo->hwnd << ",verb="
                  << (pExecInfo->lpVerb
                          ? convertUTF16ToUTF8(convertACPToUTF16(pExecInfo->lpVerb).data())
                          : "NULL")
                  << ",file=" << convertUTF16ToUTF8(convertACPToUTF16(pExecInfo->lpFile).data())
                  << ",parameters="
                  << (pExecInfo->lpParameters
                          ? convertUTF16ToUTF8(convertACPToUTF16(pExecInfo->lpParameters).data())
                          : "NULL")
                  << ",directory="
                  << (pExecInfo->lpDirectory
                          ? convertUTF16ToUTF8(convertACPToUTF16(pExecInfo->lpDirectory).data())
                          : "NULL")
                  << ",show=" << pExecInfo->nShow << "})..." << std::endl;

    BOOL nResult = ShellExecuteExA(pExecInfo);
    DWORD nLastError = GetLastError();

    if (pGlobalParamPtr->mbVerbose)
    {
        std::cout << "...ShellExecuteExA(): " << (nResult ? "TRUE" : "FALSE");
        if (!nResult)
            std::cout << ": " << WindowsErrorString(nLastError);

        std::cout << std::endl;
    }

    SetLastError(nLastError);

    return nResult;
}

static BOOL myShellExecuteExW(SHELLEXECUTEINFOW* pExecInfo)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "ShellExecuteExW({hwnd=" << pExecInfo->hwnd << ",verb="
                  << (pExecInfo->lpVerb ? convertUTF16ToUTF8(pExecInfo->lpVerb) : "NULL")
                  << ",file=" << convertUTF16ToUTF8(pExecInfo->lpFile) << ",parameters="
                  << (pExecInfo->lpParameters ? convertUTF16ToUTF8(pExecInfo->lpParameters)
                                              : "NULL")
                  << ",directory="
                  << (pExecInfo->lpDirectory ? convertUTF16ToUTF8(pExecInfo->lpDirectory) : "NULL")
                  << ",show=" << pExecInfo->nShow << "})..." << std::endl;

    BOOL nResult = ShellExecuteExW(pExecInfo);
    DWORD nLastError = GetLastError();

    if (pGlobalParamPtr->mbVerbose)
    {
        std::cout << "...ShellExecuteExW(): " << (nResult ? "TRUE" : "FALSE");
        if (!nResult)
            std::cout << ": " << WindowsErrorString(nLastError);

        std::cout << std::endl;
    }

    SetLastError(nLastError);

    return nResult;
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
    HMODULE hShell32 = GetModuleHandleW(L"shell32.dll");
    HMODULE hKernel32 = GetModuleHandleW(L"kernel32.dll");
    FunPtr pFun;

    if (hModule == hKernel32 && std::strcmp(lpProcName, "OutputDebugStringA") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(kernel32.dll, OutputDebugStringA) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myOutputDebugStringA;
        return pFun.pProc;
    }

    if (hModule == hKernel32 && std::strcmp(lpProcName, "OutputDebugStringW") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(kernel32.dll, OutputDebugStringW) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myOutputDebugStringW;
        return pFun.pProc;
    }

    if (hModule == hShell32 && std::strcmp(lpProcName, "ShellExecuteA") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(shell32.dll, ShellExecuteA) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myShellExecuteA;
        return pFun.pProc;
    }

    if (hModule == hShell32 && std::strcmp(lpProcName, "ShellExecuteW") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(shell32.dll, ShellExecuteW) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myShellExecuteW;
        return pFun.pProc;
    }

    if (hModule == hShell32 && std::strcmp(lpProcName, "ShellExecuteExA") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(shell32.dll, ShellExecuteExA) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myShellExecuteExA;
        return pFun.pProc;
    }

    if (hModule == hShell32 && std::strcmp(lpProcName, "ShellExecuteExW") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(shell32.dll, ShellExecuteExW) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myShellExecuteExW;
        return pFun.pProc;
    }

    if (hModule == hOle32 && std::strcmp(lpProcName, "OleCreateLink") == 0)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "GetProcAddress(ole32.dll, OleCreateLink) from "
                      << prettyCodeAddress(_ReturnAddress()) << std::endl;
        pFun.pVoid = myOleCreateLink;
        return pFun.pProc;
    }

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
                  << prettyCodeAddress(_ReturnAddress()) << ": " << pRetval << std::endl;

    return pRetval;
}

static DLL_DIRECTORY_COOKIE WINAPI myAddDllDirectory(PCWSTR NewDirectory)
{
    DLL_DIRECTORY_COOKIE nResult = (*pAddDllDirectory)(NewDirectory);

    if (pGlobalParamPtr->mbVerbose)
        std::cout << "AddDllDirectory(" << convertUTF16ToUTF8(NewDirectory) << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << ": " << nResult << std::endl;

    return nResult;
}

static BOOL WINAPI mySetDllDirectoryW(LPCWSTR lpPathName)
{
    BOOL bResult = (*pSetDllDirectoryW)(lpPathName);

    if (pGlobalParamPtr->mbVerbose)
        std::cout << "SetDllDirectoryW(" << convertUTF16ToUTF8(lpPathName) << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << ": " << bResult << std::endl;

    return bResult;
}

static BOOL WINAPI mySetDllDirectoryA(LPCSTR lpPathName)
{
    BOOL bResult = (*pSetDllDirectoryA)(lpPathName);

    if (pGlobalParamPtr->mbVerbose)
        std::cout << "SetDllDirectoryA(" << convertUTF16ToUTF8(convertACPToUTF16(lpPathName).data())
                  << ") from " << prettyCodeAddress(_ReturnAddress()) << ": " << bResult
                  << std::endl;

    return bResult;
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

        hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "GetProcAddress",
             myGetProcAddress);
        if (pAddDllDirectory)
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "AddDllDirectory",
                 myAddDllDirectory);
        if (pSetDllDirectoryW)
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "SetDllDirectoryW",
                 mySetDllDirectoryW);
        if (pSetDllDirectoryA)
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "SetDllDirectoryA",
                 mySetDllDirectoryA);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "LoadLibraryW",
             myLoadLibraryW);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "LoadLibraryA",
             myLoadLibraryA);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "LoadLibraryExW",
             myLoadLibraryExW);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "LoadLibraryExA",
             myLoadLibraryExA);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"ntdll.dll", "LdrLoadDll", myLdrLoadDll);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", "CoCreateInstance",
             myCoCreateInstance);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", "CoCreateInstanceEx",
             myCoCreateInstanceEx);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", "OleCreateLink",
             myOleCreateLink);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "OutputDebugStringA",
             myOutputDebugStringA);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "OutputDebugStringW",
             myOutputDebugStringW);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"shell32.dll", "ShellExecuteA",
             myShellExecuteA);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"shell32.dll", "ShellExecuteW",
             myShellExecuteW);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"shell32.dll", "ShellExecuteExA",
             myShellExecuteExA);
        hook(false, pGlobalParamPtr, hModule, lpFileName, L"shell32.dll", "ShellExecuteExW",
             myShellExecuteExW);
    }

    return hModule;
}

static HMODULE WINAPI myLoadLibraryA(LPCSTR lpFileName)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "LoadLibraryA(" << convertUTF16ToUTF8(convertACPToUTF16(lpFileName).data())
                  << ") from " << prettyCodeAddress(_ReturnAddress()) << "..." << std::endl;

    HMODULE hModule = LoadLibraryA(lpFileName);
    DWORD nLastError = GetLastError();

    if (pGlobalParamPtr->mbVerbose)
        std::cout << "...LoadLibraryA(" << convertUTF16ToUTF8(convertACPToUTF16(lpFileName).data())
                  << "): " << hModule;

    if (hModule == NULL)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << ": " << WindowsErrorString(nLastError) << std::endl;
    }
    else
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << std::endl;

        std::wstring sWFileName = convertACPToUTF16(lpFileName);

        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll", "GetProcAddress",
             myGetProcAddress);
        if (pAddDllDirectory)
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                 "AddDllDirectory", myAddDllDirectory);
        if (pSetDllDirectoryW)
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                 "SetDllDirectoryW", mySetDllDirectoryW);
        if (pSetDllDirectoryA)
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                 "SetDllDirectoryA", mySetDllDirectoryA);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll", "LoadLibraryW",
             myLoadLibraryW);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll", "LoadLibraryA",
             myLoadLibraryA);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll", "LoadLibraryExW",
             myLoadLibraryExW);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll", "LoadLibraryExA",
             myLoadLibraryExA);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"ntdll.dll", "LdrLoadDll",
             myLdrLoadDll);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"ole32.dll", "CoCreateInstance",
             myCoCreateInstance);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"ole32.dll", "CoCreateInstanceEx",
             myCoCreateInstanceEx);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"ole32.dll", "OleCreateLink",
             myOleCreateLink);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
             "OutputDebugStringA", myOutputDebugStringA);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
             "OutputDebugStringW", myOutputDebugStringW);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll", "ShellExecuteA",
             myShellExecuteA);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll", "ShellExecuteW",
             myShellExecuteW);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll", "ShellExecuteExA",
             myShellExecuteExA);
        hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll", "ShellExecuteExW",
             myShellExecuteExW);
    }

    return hModule;
}

static HMODULE WINAPI innerMyLoadLibraryExW(const std::string& caller, LPCWSTR lpFileName,
                                            HANDLE hFile, DWORD dwFlags)
{
    HMODULE hModule = LoadLibraryExW(lpFileName, hFile, dwFlags);
    DWORD nLastError = GetLastError();

    if (hModule == NULL)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "..." << caller << "(" << convertUTF16ToUTF8(lpFileName) << ", 0x"
                      << to_uhex(dwFlags) << "): " << hModule << ": "
                      << WindowsErrorString(nLastError) << std::endl;
    }
    else
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "...LoadLibraryExW(" << convertUTF16ToUTF8(lpFileName) << ", 0x"
                      << to_uhex(dwFlags) << "): " << hModule << std::endl;

        if (!(dwFlags
              & (LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE
                 | LOAD_LIBRARY_AS_IMAGE_RESOURCE)))
        {
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "GetProcAddress",
                 myGetProcAddress);
            if (pAddDllDirectory)
                hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll",
                     "AddDllDirectory", myAddDllDirectory);
            if (pSetDllDirectoryW)
                hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll",
                     "SetDllDirectoryW", mySetDllDirectoryW);
            if (pSetDllDirectoryA)
                hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll",
                     "SetDllDirectoryA", mySetDllDirectoryA);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "LoadLibraryW",
                 myLoadLibraryW);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "LoadLibraryA",
                 myLoadLibraryA);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "LoadLibraryExW",
                 myLoadLibraryExW);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"ntdll.dll", "LdrLoadDll",
                 myLdrLoadDll);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", "CoCreateInstance",
                 myCoCreateInstance);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", "CoCreateInstanceEx",
                 myCoCreateInstanceEx);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"ole32.dll", "OleCreateLink",
                 myOleCreateLink);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "OutputDebugStringA",
                 myOutputDebugStringA);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"kernel32.dll", "OutputDebugStringW",
                 myOutputDebugStringW);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"shell32.dll", "ShellExecuteA",
                 myShellExecuteA);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"shell32.dll", "ShellExecuteW",
                 myShellExecuteW);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"shell32.dll", "ShellExecuteExA",
                 myShellExecuteExA);
            hook(false, pGlobalParamPtr, hModule, lpFileName, L"shell32.dll", "ShellExecuteExW",
                 myShellExecuteExW);
        }
    }

    return hModule;
}

static HMODULE WINAPI myLoadLibraryExW(LPCWSTR lpFileName, HANDLE hFile, DWORD dwFlags)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "LoadLibraryExW(" << convertUTF16ToUTF8(lpFileName) << ", 0x"
                  << to_uhex(dwFlags) << ") from " << prettyCodeAddress(_ReturnAddress()) << "..."
                  << std::endl;

    return innerMyLoadLibraryExW("LoadLibraryExW", lpFileName, hFile, dwFlags);
}

static HMODULE WINAPI myLoadLibraryExA(LPCSTR lpFileName, HANDLE hFile, DWORD dwFlags)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "LoadLibraryExA(" << convertUTF16ToUTF8(convertACPToUTF16(lpFileName).data())
                  << ", 0x" << to_uhex(dwFlags) << ") from " << prettyCodeAddress(_ReturnAddress())
                  << "..." << std::endl;

    HMODULE hModule = LoadLibraryExA(lpFileName, hFile, dwFlags);
    DWORD nLastError = GetLastError();

    if (hModule == NULL)
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "...LoadLibraryExA("
                      << convertUTF16ToUTF8(convertACPToUTF16(lpFileName).data()) << ", 0x"
                      << to_uhex(dwFlags) << "): " << hModule << ": "
                      << WindowsErrorString(nLastError) << std::endl;
    }
    else
    {
        if (pGlobalParamPtr->mbVerbose)
            std::cout << "...LoadLibraryExA("
                      << convertUTF16ToUTF8(convertACPToUTF16(lpFileName).data()) << ", 0x"
                      << to_uhex(dwFlags) << "): " << hModule << std::endl;

        if (!(dwFlags
              & (LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE
                 | LOAD_LIBRARY_AS_IMAGE_RESOURCE)))
        {
            std::wstring sWFileName = convertACPToUTF16(lpFileName);

            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                 "GetProcAddress", myGetProcAddress);
            if (pAddDllDirectory)
                hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                     "AddDllDirectory", myAddDllDirectory);
            if (pSetDllDirectoryW)
                hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                     "SetDllDirectoryW", mySetDllDirectoryW);
            if (pSetDllDirectoryA)
                hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                     "SetDllDirectoryA", mySetDllDirectoryA);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                 "LoadLibraryW", myLoadLibraryW);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                 "LoadLibraryA", myLoadLibraryA);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                 "LoadLibraryExW", myLoadLibraryExW);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                 "LoadLibraryExA", myLoadLibraryExA);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"ntdll.dll", "LdrLoadDll",
                 myLdrLoadDll);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"ole32.dll",
                 "CoCreateInstance", myCoCreateInstance);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"ole32.dll",
                 "CoCreateInstanceEx", myCoCreateInstanceEx);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"ole32.dll", "OleCreateLink",
                 myOleCreateLink);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                 "OutputDebugStringA", myOutputDebugStringA);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"kernel32.dll",
                 "OutputDebugStringW", myOutputDebugStringW);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll",
                 "ShellExecuteA", myShellExecuteA);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll",
                 "ShellExecuteW", myShellExecuteW);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll",
                 "ShellExecuteExA", myShellExecuteExA);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll",
                 "ShellExecuteExW", myShellExecuteExW);
        }
    }

    return hModule;
}

static NTSTATUS NTAPI myLdrLoadDll(PWCHAR PathToFile, ULONG Flags, UNICODE_STRING* ModuleFileName,
                                   PHANDLE ModuleHandle)
{
    std::vector<WCHAR> fileName(ModuleFileName->Length + 1u);
    memcpy(fileName.data(), ModuleFileName->Buffer, ModuleFileName->Length * 2u);
    fileName[ModuleFileName->Length] = L'\0';

    if (pGlobalParamPtr->mbVerbose)
        std::cout << "LdrLoadDll(" << (PathToFile ? convertUTF16ToUTF8(PathToFile) : "(null)")
                  << ", 0x" << to_uhex(Flags) << "," << convertUTF16ToUTF8(fileName.data())
                  << ") from " << prettyCodeAddress(_ReturnAddress()) << "..." << std::endl;

    HMODULE handle = innerMyLoadLibraryExW("LdrLoadDll", fileName.data(), NULL, Flags);

    if (handle != NULL)
        *ModuleHandle = handle;
    else
        return (NTSTATUS)GetLastError();

    return 0;
}

static bool hook(bool bMandatory, ThreadProcParam* pParam, HMODULE hModule,
                 const wchar_t* sModuleName, const wchar_t* sDll, const char* sFunction,
                 PVOID pOwnFunction)
{
    ULONG nSize;
    PIMAGE_IMPORT_DESCRIPTOR pImportDescriptor
        = (PIMAGE_IMPORT_DESCRIPTOR)ImageDirectoryEntryToDataEx(
            hModule, TRUE, IMAGE_DIRECTORY_ENTRY_IMPORT, &nSize, NULL);
    if (pImportDescriptor == NULL)
    {
        if (bMandatory)
            std::cout << "Could not find import directory in " << convertUTF16ToUTF8(sModuleName)
                      << std::endl;
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
            std::cout << "Import descriptor for " << convertUTF16ToUTF8(sDll) << " not found in "
                      << convertUTF16ToUTF8(sModuleName) << std::endl;

        return false;
    }

    PROC pOriginalFunc = (PROC)GetProcAddress(GetModuleHandleW(sDll), sFunction);
    if (pOriginalFunc == NULL)
    {
        std::cout << "Could not find original " << sFunction << " in " << convertUTF16ToUTF8(sDll)
                  << std::endl;
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
                std::cout << "VirtualQuery failed: " << WindowsErrorString(GetLastError())
                          << std::endl;
                return false;
            }

            if (!VirtualProtect(aMBI.BaseAddress, aMBI.RegionSize, PAGE_READWRITE, &aMBI.Protect))
            {
                std::cout << "VirtualProtect failed: " << WindowsErrorString(GetLastError())
                          << std::endl;
                return false;
            }

            FunPtr pFun;
            pFun.pVoid = pOwnFunction;
            *pFunc = pFun.pProc;

            DWORD nOldProtect;
            if (!VirtualProtect(aMBI.BaseAddress, aMBI.RegionSize, aMBI.Protect, &nOldProtect))
            {
                std::cout << "VirtualProtect failed: " << WindowsErrorString(GetLastError())
                          << std::endl;
                return false;
            }

            nHookedFunctions++;

            if (pParam->mbVerbose)
                std::cout << "Hooked " << sFunction << " in import table for "
                          << convertUTF16ToUTF8(sDll) << " in " << convertUTF16ToUTF8(sModuleName)
                          << std::endl;
            return true;
        }

        pThunk++;
    }

    // Don't bother mentioning non-mandatory imports not found even in verbose mode.
    if (bMandatory && pParam->mbVerbose)
        std::cout << "Did not find " << sFunction << " in import table for "
                  << convertUTF16ToUTF8(sDll) << " in " << convertUTF16ToUTF8(sModuleName)
                  << std::endl;

    return false;
}

static bool hook(bool bMandatory, ThreadProcParam* pParam, const wchar_t* sModule,
                 const wchar_t* sDll, const char* sFunction, PVOID pOwnFunction)
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
        if (pParam->mbVerbose)
            std::cout << convertUTF16ToUTF8(sModuleName) << " is not loaded" << std::endl;
        return false;
    }

    return hook(bMandatory, pParam, hModule, sModuleName, sDll, sFunction, pOwnFunction);
}

extern "C" DWORD WINAPI InjectedDllMainFunction(ThreadProcParam* pParam)
{
// Magic to export this function using a plain undecorated name despite it being WINAPI
#ifdef _WIN64
#pragma comment(linker, "/EXPORT:InjectedDllMainFunction=InjectedDllMainFunction")
#else
#pragma comment(linker, "/EXPORT:InjectedDllMainFunction=_InjectedDllMainFunction@4")
#endif

    if (pParam->mnSize != sizeof(*pParam))
    {
        // We can't do anything with it if the size of *pParam is not what we expect;)
        return FALSE;
    }

    pParam->mbPassedSizeCheck = true;

    // This function returns and the remotely created thread exits, and the wrapper process will
    // copy back the parameter block, but we keep a pointer to it for use by the hook functions.
    pGlobalParamPtr = pParam;

    CProxiedUnknown::setParam(pGlobalParamPtr);

    tryToEnsureStdHandlesOpen(bDidAllocConsole);

    FunPtr aFun;
    aFun.pProc = GetProcAddress(GetModuleHandleW(L"kernel32.dll"), "AddDllDirectory");
    pAddDllDirectory = aFun.pAddDllDirectory;

    aFun.pProc = GetProcAddress(GetModuleHandleW(L"kernel32.dll"), "SetDllDirectoryW");
    pSetDllDirectoryW = aFun.pSetDllDirectoryW;

    aFun.pProc = GetProcAddress(GetModuleHandleW(L"kernel32.dll"), "SetDllDirectoryA");
    pSetDllDirectoryA = aFun.pSetDllDirectoryA;

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

        if (!hook(true, pParam, L"msvbvm60.dll", L"ole32.dll", "CoCreateInstance",
                  myCoCreateInstance))
            return FALSE;

        hook(false, pParam, L"msvbvm60.dll", L"ole32.dll", "CoCreateInstanceEx",
             myCoCreateInstanceEx);

        hook(false, pParam, L"msvbvm60.dll", L"ole32.dll", "OleCreateLink", myOleCreateLink);

        hook(false, pParam, L"msvbvm60.dll", L"kernel32.dll", "OutputDebugStringA",
             myOutputDebugStringA);

        hook(false, pParam, L"msvbvm60.dll", L"kernel32.dll", "OutputDebugStringW",
             myOutputDebugStringW);

        hook(false, pParam, L"msvbvm60.dll", L"ole32.dll", "CoGetClassObject", myCoGetClassObject);

        if (!hook(true, pParam, L"msvbvm60.dll", L"kernel32.dll", "GetProcAddress",
                  myGetProcAddress))
            return FALSE;

        hook(false, pParam, L"msvbvm60.dll", L"kernel32.dll", "LoadLibraryW", myLoadLibraryW);
        hook(false, pParam, L"msvbvm60.dll", L"kernel32.dll", "LoadLibraryA", myLoadLibraryA);
        hook(false, pParam, L"msvbvm60.dll", L"kernel32.dll", "LoadLibraryExW", myLoadLibraryExW);
        hook(false, pParam, L"msvbvm60.dll", L"kernel32.dll", "LoadLibraryExA", myLoadLibraryExA);
        hook(false, pParam, L"msvbvm60.dll", L"ntdll.dll", "LdrLoadDll", myLdrLoadDll);
        hook(false, pParam, L"msvbvm60.dll", L"shell32.dll", "ShellExecuteA", myShellExecuteA);
        hook(false, pParam, L"msvbvm60.dll", L"shell32.dll", "ShellExecuteW", myShellExecuteW);
        hook(false, pParam, L"msvbvm60.dll", L"shell32.dll", "ShellExecuteExA", myShellExecuteExA);
        hook(false, pParam, L"msvbvm60.dll", L"shell32.dll", "ShellExecuteExW", myShellExecuteExW);
    }
    else
    {
        // It is some other executable. We must hook LoadLibrary*().

        nHookedFunctions = 0;
        hook(false, pParam, nullptr, L"kernel32.dll", "GetProcAddress", myGetProcAddress);
        if (pAddDllDirectory)
            hook(false, pParam, nullptr, L"kernel32.dll", "AddDllDirectory", myAddDllDirectory);
        if (pSetDllDirectoryW)
            hook(false, pParam, nullptr, L"kernel32.dll", "SetDllDirectoryW", mySetDllDirectoryW);
        if (pSetDllDirectoryA)
            hook(false, pParam, nullptr, L"kernel32.dll", "SetDllDirectoryA", mySetDllDirectoryA);
        hook(false, pParam, nullptr, L"kernel32.dll", "LoadLibraryW", myLoadLibraryW);
        hook(false, pParam, nullptr, L"kernel32.dll", "LoadLibraryA", myLoadLibraryA);
        hook(false, pParam, nullptr, L"kernel32.dll", "LoadLibraryExW", myLoadLibraryExW);
        hook(false, pParam, nullptr, L"kernel32.dll", "LoadLibraryExA", myLoadLibraryExA);
        hook(false, pParam, nullptr, L"ntdll.dll", "LdrLoadDll", myLdrLoadDll);
        hook(false, pParam, nullptr, L"shell32.dll", "ShellExecuteA", myShellExecuteA);
        hook(false, pParam, nullptr, L"shell32.dll", "ShellExecuteW", myShellExecuteW);
        hook(false, pParam, nullptr, L"shell32.dll", "ShellExecuteExA", myShellExecuteExA);
        hook(false, pParam, nullptr, L"shell32.dll", "ShellExecuteExW", myShellExecuteExW);
        hook(false, pParam, nullptr, L"ole32.dll", "CoCreateInstance", myCoCreateInstance);
        hook(false, pParam, nullptr, L"ole32.dll", "CoCreateInstanceEx", myCoCreateInstanceEx);
        hook(false, pParam, nullptr, L"ole32.dll", "OleCreateLink", myOleCreateLink);
        hook(false, pParam, nullptr, L"kernel32.dll", "OutputDebugStringA", myOutputDebugStringA);
        hook(false, pParam, nullptr, L"kernel32.dll", "OutputDebugStringW", myOutputDebugStringW);
        hook(false, pParam, nullptr, L"ole32.dll", "CoGetClassObject", myCoGetClassObject);
        if (nHookedFunctions == 0)
        {
            std::cout << "Could not hook a single interesting function to hook" << std::endl;
            return FALSE;
        }
    }

    return TRUE;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

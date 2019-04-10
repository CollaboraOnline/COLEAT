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
#include <fstream>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>

#include <intrin.h>
#include <process.h>

#include <Windows.h>
#include <gdiplus.h>

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

class myViewObjectWrapper : IViewObject
{
private:
    IViewObject* const mpWrappedViewObject;
    const std::wstring* const msImageFile;

public:
    myViewObjectWrapper(IViewObject* pViewObject, const std::wstring* const sImageFile) :
        mpWrappedViewObject(pViewObject),
        msImageFile(sImageFile)
    {
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        return mpWrappedViewObject->QueryInterface(riid, ppvObject);
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return mpWrappedViewObject->AddRef();
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG nRetval = mpWrappedViewObject->Release();
        if (nRetval == 0)
        {
            _wremove(msImageFile->data());
            delete msImageFile;
            delete this;
        }
        return nRetval;
    }

    // IViewObject
    HRESULT STDMETHODCALLTYPE Draw(DWORD dwDrawAspect, LONG lindex, void *pvAspect, DVTARGETDEVICE *ptd,
                                   HDC hdcTargetDev, HDC hdcDraw, LPCRECTL lprcBounds, LPCRECTL lprcWBounds,
                                   BOOL (STDMETHODCALLTYPE *pfnContinue)(ULONG_PTR dwContinue),
                                   ULONG_PTR dwContinue) override
    {
        (void) dwDrawAspect, lindex, pvAspect, ptd, hdcTargetDev,  lprcWBounds, pfnContinue, dwContinue;

        Gdiplus::GdiplusStartupInput gdiplusStartupInput;
        ULONG_PTR gdiplusToken;
        Gdiplus::GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

        Gdiplus::Bitmap* pBitmap = Gdiplus::Bitmap::FromFile( msImageFile->data() );

        HBITMAP hBitmap;
        pBitmap->GetHBITMAP(Gdiplus::Color(), &hBitmap);

        HDC hdcMem = CreateCompatibleDC(hdcDraw);
        HGDIOBJ hBitmapOld = SelectObject(hdcMem, hBitmap);

        BITMAP aBitmap;
        GetObject(hBitmap, sizeof(aBitmap), &aBitmap);

        BitBlt(hdcDraw, lprcBounds->left, lprcBounds->top, aBitmap.bmWidth, aBitmap.bmHeight, hdcMem, 0, 0, SRCCOPY);

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

        Gdiplus::GdiplusShutdown(gdiplusToken);

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetColorSet(DWORD dwDrawAspect, LONG lindex, void *pvAspect, DVTARGETDEVICE *ptd,
                                          HDC hicTargetDev, LOGPALETTE **ppColorSet) override
    {
        return mpWrappedViewObject->GetColorSet(dwDrawAspect, lindex, pvAspect, ptd, hicTargetDev, ppColorSet);
    }

    HRESULT STDMETHODCALLTYPE Freeze(DWORD dwDrawAspect, LONG lindex, void *pvAspect, DWORD *pdwFreeze) override
    {
        return mpWrappedViewObject->Freeze(dwDrawAspect, lindex, pvAspect, pdwFreeze);
    }

    HRESULT STDMETHODCALLTYPE Unfreeze(DWORD dwFreeze) override
    {
        return mpWrappedViewObject->Unfreeze(dwFreeze);
    }

    HRESULT STDMETHODCALLTYPE SetAdvise(DWORD aspects, DWORD advf, IAdviseSink *pAdvSink) override
    {
        return mpWrappedViewObject->SetAdvise(aspects, advf, pAdvSink);
    }

    HRESULT STDMETHODCALLTYPE GetAdvise(DWORD *pAspects, DWORD *pAdvf, IAdviseSink **ppAdvSink) override
    {
        return mpWrappedViewObject->GetAdvise(pAspects, pAdvf, ppAdvSink);
    }
};

class myOleObjectWrapper : IOleObject
{
private:
    IOleObject* const mpWrappedOleObject;
    const std::wstring* const msEmptyDocument;
    const std::wstring* const msImageFile;

public:
    myOleObjectWrapper(IOleObject* pOleObject, const std::wstring* const sEmptyDocument, const std::wstring* const sImageFile) :
        mpWrappedOleObject(pOleObject),
        msEmptyDocument(sEmptyDocument),
        msImageFile(sImageFile)
    {
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        HRESULT nResult = mpWrappedOleObject->QueryInterface(riid, ppvObject);

        if (nResult == S_OK && IsEqualIID(riid, IID_IViewObject))
        {
            *ppvObject = new myViewObjectWrapper(static_cast<IViewObject*>(*ppvObject), msImageFile);
        }
        return nResult;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return mpWrappedOleObject->AddRef();
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG nRetval = mpWrappedOleObject->Release();
        if (nRetval == 0)
        {
            _wremove(msEmptyDocument->data());
            delete msEmptyDocument;
            delete this;
        }
        return nRetval;
    }

    // IOleObject
    HRESULT STDMETHODCALLTYPE SetClientSite(IOleClientSite *pClientSite) override
    {
        return mpWrappedOleObject->SetClientSite(pClientSite);
    }

    HRESULT STDMETHODCALLTYPE GetClientSite(IOleClientSite **ppClientSite) override
    {
        return mpWrappedOleObject->GetClientSite(ppClientSite);
    }

    HRESULT STDMETHODCALLTYPE SetHostNames(LPCOLESTR szContainerApp, LPCOLESTR szContainerObj) override
    {
        return mpWrappedOleObject->SetHostNames(szContainerApp, szContainerObj);
    }

    HRESULT STDMETHODCALLTYPE Close(DWORD dwSaveOption) override
    {
        return mpWrappedOleObject->Close(dwSaveOption);
    }

    HRESULT STDMETHODCALLTYPE SetMoniker(DWORD dwWhichMoniker, IMoniker *pmk) override
    {
        return mpWrappedOleObject->SetMoniker(dwWhichMoniker, pmk);
    }

    HRESULT STDMETHODCALLTYPE GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk) override
    {
        return mpWrappedOleObject->GetMoniker(dwAssign, dwWhichMoniker, ppmk);
    }

    HRESULT STDMETHODCALLTYPE InitFromData(IDataObject *pDataObject, BOOL fCreation, DWORD dwReserved) override
    {
        return mpWrappedOleObject->InitFromData(pDataObject, fCreation, dwReserved);
    }

    HRESULT STDMETHODCALLTYPE GetClipboardData(DWORD dwReserved, IDataObject **ppDataObject) override
    {
        return mpWrappedOleObject->GetClipboardData(dwReserved, ppDataObject);
    }

    HRESULT STDMETHODCALLTYPE DoVerb(LONG iVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex,
                                     HWND hwndParent, LPCRECT lprcPosRect) override
    {
        return mpWrappedOleObject->DoVerb(iVerb, lpmsg, pActiveSite, lindex, hwndParent, lprcPosRect);
    }

    HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB **ppEnumOleVerb) override
    {
        return mpWrappedOleObject->EnumVerbs(ppEnumOleVerb);
    }

    HRESULT STDMETHODCALLTYPE Update(void) override
    {
        return mpWrappedOleObject->Update();
    }

    HRESULT STDMETHODCALLTYPE IsUpToDate(void) override
    {
        return mpWrappedOleObject->IsUpToDate();
    }

    HRESULT STDMETHODCALLTYPE GetUserClassID(CLSID *pClsid) override
    {
        return mpWrappedOleObject->GetUserClassID(pClsid);
    }

    HRESULT STDMETHODCALLTYPE GetUserType(DWORD dwFormOfType, LPOLESTR *pszUserType) override
    {
        return mpWrappedOleObject->GetUserType(dwFormOfType, pszUserType);
    }

    HRESULT STDMETHODCALLTYPE SetExtent(DWORD dwDrawAspect, SIZEL *psizel) override
    {
        return mpWrappedOleObject->SetExtent(dwDrawAspect, psizel);
    }

    HRESULT STDMETHODCALLTYPE GetExtent(DWORD dwDrawAspect, SIZEL *psizel) override
    {
        return mpWrappedOleObject->GetExtent(dwDrawAspect, psizel);
    }

    HRESULT STDMETHODCALLTYPE Advise(IAdviseSink *pAdvSink, DWORD *pdwConnection) override
    {
        return mpWrappedOleObject->Advise(pAdvSink, pdwConnection);
    }

    HRESULT STDMETHODCALLTYPE Unadvise(DWORD dwConnection) override
    {
        return mpWrappedOleObject->Unadvise(dwConnection);
    }

    HRESULT STDMETHODCALLTYPE EnumAdvise(IEnumSTATDATA **ppenumAdvise) override
    {
        return mpWrappedOleObject->EnumAdvise(ppenumAdvise);
    }

    HRESULT STDMETHODCALLTYPE GetMiscStatus(DWORD dwAspect, DWORD *pdwStatus) override
    {
        return mpWrappedOleObject->GetMiscStatus(dwAspect, pdwStatus);
    }

    HRESULT STDMETHODCALLTYPE SetColorScheme(LOGPALETTE *pLogpal) override
    {
        return mpWrappedOleObject->SetColorScheme(pLogpal);
    }
};


static HRESULT tryRenderDrawInCollaboraOffice(LPMONIKER pmkLinkSrc, REFIID riid, DWORD renderopt,
                                              LPFORMATETC lpFormatEtc, LPOLECLIENTSITE pClientSite,
                                              LPSTORAGE pStg, LPVOID *ppvObj)
{
    // Sanity check. Only attempt for the cae we are interested in.
    if (!IsEqualIID(riid, IID_IOleObject) || renderopt != OLERENDER_DRAW || pClientSite == NULL || pStg == NULL)
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

    IMalloc *pMalloc;

    nResult = CoGetMalloc(1, &pMalloc);
    if (nResult != S_OK)
    {
        std::cout << "CoGetMalloc failed: " << HRESULT_to_string(nResult) << "\n";
        return S_FALSE;
    }

    // We need a bit context. Sadly apparently there is no way to get the one that the calling code
    // must have used to create the pmkLinkSrc IMoniker?
    IBindCtx *pBindContext;
    nResult = CreateBindCtx(0, &pBindContext);

    if (nResult != S_OK)
    {
        std::cout << "CreateBindCtx failed: " << HRESULT_to_string(nResult) << "\n";
        return S_FALSE;
    }

    // The display name of the IMoniker is the file pathname.
    wchar_t *sDisplayName;
    nResult = pmkLinkSrc->GetDisplayName(pBindContext, NULL, &sDisplayName);

    if (nResult != S_OK)
    {
        std::cout << "IMoniker::GetDisplayName failed: " << HRESULT_to_string(nResult) << "\n";
        pBindContext->Release();
        return S_FALSE;
    }

    // Check if the display name is a pathname of a file with an extension we "know" that Collabora
    // Office handles. Only check a few extensions, and in the actual customer case, it is .rtf.
    wchar_t *pBasenameEnd = wcsrchr(sDisplayName, L'.');
    if (pBasenameEnd == nullptr)
    {
        // No file name extension, let's not even try.
        std::cout << "No file name extension in '" << convertUTF16ToUTF8(sDisplayName) << "'\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    const wchar_t *pLastSlash = wcsrchr(sDisplayName, L'\\');
    const std::wstring sBasename(pLastSlash + 1, static_cast<std::size_t>(pBasenameEnd - pLastSlash - 1));
    const std::wstring sExtension(pBasenameEnd + 1);

    if (!(sExtension == L"rtf" || sExtension == L"doc" || sExtension == L"docx" || sExtension == L"odt"))
    {
        std::cout << "Not a known file name extension in '" << convertUTF16ToUTF8(sDisplayName) << "'\n";
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

    wchar_t sTempPath[MAX_PATH+1];
    if (GetTempPathW(MAX_PATH+1, sTempPath) == 0)
    {
        std::cout << "GetTempPathW failed!\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }
    if (sTempPath[wcslen(sTempPath)-1] == L'\\')
        sTempPath[wcslen(sTempPath)-1] = L'\0';

    // Where is CO installed?

    HKEY hUno;
    nResult = RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\LibreOffice\\UNO\\InstallPath", 0, KEY_READ | KEY_WOW64_32KEY , &hUno);
    if (nResult == ERROR_FILE_NOT_FOUND)
        nResult = RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\LibreOffice\\UNO\\InstallPath", 0, KEY_READ | KEY_WOW64_64KEY , &hUno);

    if (nResult != ERROR_SUCCESS)
    {
        std::cout << "Can not find Collabora Office installation, RegOpenKeyExW failed: " << HRESULT_to_string(nResult) << "\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    wchar_t sLOPath[MAX_PATH+1];
    LONG nLOPathSize = sizeof(sLOPath);
    nResult = RegQueryValueW(hUno, NULL, sLOPath, &nLOPathSize);
    if (nResult != ERROR_SUCCESS)
    {
        std::cout << "Can not find Collabora Office installation, RegQueryValueW failed: " << HRESULT_to_string(nResult) << "\n";
        RegCloseKey(hUno);
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    RegCloseKey(hUno);

    std::wstring sCommandLine = L"\"" + std::wstring(sLOPath) + L"\\soffice.exe\" --convert-to png --outdir \"" + sTempPath
        + L"\" \"" + std::wstring(sDisplayName) + L"\"";

    STARTUPINFOW aStartupInfo;
    PROCESS_INFORMATION aProcessInformation;

    memset(&aStartupInfo, 0, sizeof(aStartupInfo));
    aStartupInfo.cb = sizeof(aStartupInfo);
    aStartupInfo.lpDesktop = L"";

    nResult = CreateProcessW(NULL, const_cast<wchar_t*>(sCommandLine.data()), NULL, NULL, FALSE, 0, NULL, NULL, &aStartupInfo, &aProcessInformation);
    if (nResult == 0)
    {
        std::cout << "Can not run soffice, CreateProcessW failed: " << HRESULT_to_string(nResult) << "\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    WaitForSingleObject(aProcessInformation.hProcess, INFINITE);

    std::wstring sImageFile = std::wstring(sTempPath) + L"\\" + sBasename + L".png";

    // Create an OLE link to a dummy (empty) document that we know WordPad will be able to handle
    // just fine.

    // First write the dummy RTF document.

    const std::time_t nNow = std::chrono::system_clock::to_time_t(std::chrono::system_clock::now());
    const std::wstring sEmptyDocument = std::wstring(sTempPath) + L"\\" + std::to_wstring(nNow) + L".rtf";

    std::ofstream aEmptyDocument(sEmptyDocument);

    aEmptyDocument << "{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\\deflang1033{\\fonttbl{\\f0\\fnil\\fcharset0 Calibri;}}\n"
        "{\\*\\generator Riched20 10.0.17763}\\viewkind4\\uc1\n"
        "\\pard\\sa200\\sl276\\slmult1\\f0\\fs22\\lang9\\par\n"
        "HAHAHA}\n";

    aEmptyDocument.close();

    ULONG nChEaten;
    IMoniker *pEmptyDocumentMoniker;
    nResult = MkParseDisplayName(pBindContext, sEmptyDocument.data(), &nChEaten, &pEmptyDocumentMoniker);
    if (nResult != S_OK)
    {
        std::cout << "MkParseDisplayName failed: " << HRESULT_to_string(nResult) << "\n";
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    HRESULT nRetval = OleCreateLink(pEmptyDocumentMoniker, riid, renderopt, lpFormatEtc, pClientSite, pStg, ppvObj);
    if (nRetval != S_OK)
    {
        std::cout << "OleCreateLink failed: " << HRESULT_to_string(nResult) << "\n";
        pEmptyDocumentMoniker->Release();
        pMalloc->Free(sDisplayName);
        pBindContext->Release();
        return S_FALSE;
    }

    // Replace the return value of the OleCreateLink() call for the dummy document with a wrapper
    // for IOleObject that when queried for IViewObject will return a wrapper that will draw the PNG
    // image that Collabora Office created.
    *ppvObj = new myOleObjectWrapper(static_cast<IOleObject*>(*ppvObj),
                                     new std::wstring(sEmptyDocument),
                                     new std::wstring(sImageFile));

    pEmptyDocumentMoniker->Release();
    pMalloc->Free(sDisplayName);
    pBindContext->Release();

    return S_OK;
}

#endif

static HRESULT WINAPI myOleCreateLink(LPMONIKER pmkLinkSrc, REFIID riid, DWORD renderopt,
                                      LPFORMATETC lpFormatEtc, LPOLECLIENTSITE pClientSite,
                                      LPSTORAGE pStg, LPVOID *ppvObj)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "OleCreateLink(" << pmkLinkSrc << "," << riid << ","
                  << renderopt << "," << lpFormatEtc << "," << pClientSite << ","
                  << pStg << "," << ppvObj << ")..." << std::endl;

    HRESULT nRetval;

#ifdef HARDCODE_MSO_TO_CO
    nRetval = tryRenderDrawInCollaboraOffice(pmkLinkSrc, riid, renderopt, lpFormatEtc, pClientSite, pStg, ppvObj);

    if (nRetval == S_OK)
    {
        if (pGlobalParamPtr->mbVerbose)
        {
            std::cout << "...OleCreateLink(" << pmkLinkSrc << "," << riid << ",...): "
                      << HRESULT_to_string(nRetval) << std::endl;
        }
        return S_OK;
    }
#endif

    nRetval = OleCreateLink(pmkLinkSrc, riid, renderopt, lpFormatEtc, pClientSite, pStg, ppvObj);

    if (pGlobalParamPtr->mbVerbose)
    {
        std::cout << "...OleCreateLink(" << pmkLinkSrc << "," << riid << ",...): "
                  << HRESULT_to_string(nRetval) << std::endl;
    }

    return nRetval;
}

static HINSTANCE myShellExecuteA(HWND hwnd, LPCSTR lpOperation, LPCSTR lpFile, LPCSTR lpParameters,
                                 LPCSTR lpDirectory, INT nShowCmd)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "ShellExecuteA(" << hwnd << ","
                  << (lpOperation ? lpOperation : "NULL") << ","
                  << convertUTF16ToUTF8(convertACPToUTF16(lpFile).data()) << ","
                  << (lpParameters ? convertUTF16ToUTF8(convertACPToUTF16(lpParameters).data()) : "NULL") << ","
                  << (lpDirectory ? convertUTF16ToUTF8(convertACPToUTF16(lpDirectory).data()) : "NULL") << ","
                  << nShowCmd << ")..." << std::endl;

    HINSTANCE hInstance = ShellExecuteA(hwnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd);
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

static HINSTANCE myShellExecuteW(HWND hwnd, LPCWSTR lpOperation, LPCWSTR lpFile, LPCWSTR lpParameters,
                                 LPCWSTR lpDirectory, INT nShowCmd)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "ShellExecuteW(" << hwnd << ","
                  << (lpOperation ? convertUTF16ToUTF8(lpOperation) : "NULL") << ","
                  << convertUTF16ToUTF8(lpFile) << ","
                  << (lpParameters ? convertUTF16ToUTF8(lpParameters) : "NULL") << ","
                  << (lpDirectory ? convertUTF16ToUTF8(lpDirectory) : "NULL") << ","
                  << nShowCmd << ")..." << std::endl;

    HINSTANCE hInstance = ShellExecuteW(hwnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd);
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

static BOOL myShellExecuteExA(SHELLEXECUTEINFOA *pExecInfo)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "ShellExecuteExA({hwnd=" << pExecInfo->hwnd
                  << ",verb="
                  << (pExecInfo->lpVerb ? convertUTF16ToUTF8(convertACPToUTF16(pExecInfo->lpVerb).data()) : "NULL") << ",file="
                  << convertUTF16ToUTF8(convertACPToUTF16(pExecInfo->lpFile).data()) << ",parameters="
                  << (pExecInfo->lpParameters ? convertUTF16ToUTF8(convertACPToUTF16(pExecInfo->lpParameters).data()) : "NULL") << ",directory="
                  << (pExecInfo->lpDirectory ? convertUTF16ToUTF8(convertACPToUTF16(pExecInfo->lpDirectory).data()) : "NULL") << ",show="
                  << pExecInfo->nShow << "})..." << std::endl;

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

static BOOL myShellExecuteExW(SHELLEXECUTEINFOW *pExecInfo)
{
    if (pGlobalParamPtr->mbVerbose)
        std::cout << "ShellExecuteExW({hwnd=" << pExecInfo->hwnd
                  << ",verb="
                  << (pExecInfo->lpVerb ? convertUTF16ToUTF8(pExecInfo->lpVerb) : "NULL") << ",file="
                  << convertUTF16ToUTF8(pExecInfo->lpFile) << ",parameters="
                  << (pExecInfo->lpParameters ? convertUTF16ToUTF8(pExecInfo->lpParameters) : "NULL") << ",directory="
                  << (pExecInfo->lpDirectory ? convertUTF16ToUTF8(pExecInfo->lpDirectory) : "NULL") << ",show="
                  << pExecInfo->nShow << "})..." << std::endl;

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
    FunPtr pFun;

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
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"ole32.dll",
                 "OleCreateLink", myOleCreateLink);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll", "ShellExecuteA",
                 myShellExecuteA);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll", "ShellExecuteW",
                 myShellExecuteW);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll", "ShellExecuteExA",
                 myShellExecuteExA);
            hook(false, pGlobalParamPtr, hModule, sWFileName.data(), L"shell32.dll", "ShellExecuteExW",
                 myShellExecuteExW);
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
        std::cout << "LdrLoadDll(" << convertUTF16ToUTF8(PathToFile) << ", 0x" << to_uhex(Flags)
                  << "," << convertUTF16ToUTF8(fileName.data()) << ") from "
                  << prettyCodeAddress(_ReturnAddress()) << "..." << std::endl;

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

        hook(false, pParam, L"msvbvm60.dll", L"ole32.dll", "OleCreateLink",
             myOleCreateLink);

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

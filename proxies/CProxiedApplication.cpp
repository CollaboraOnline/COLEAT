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
#define _CRT_SECURE_NO_WARNINGS

#include <cassert>
#include <cstdio>
#include <codecvt>
#include <cstdlib>
#include <cwchar>
#include <iomanip>
#include <iostream>
#include <map>
#include <sstream>
#include <string>

#include <Windows.h>

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedApplication.hpp"
#include "DefaultInterfaceCreator.hxx"

bool CProxiedApplication::mbIsActive = false;

CProxiedApplication::CProxiedApplication(const InterfaceMapping& rMapping)
    : maProxiedAppCoclassIID(rMapping.maFromCoclass)
    , maProxiedAppDefaultIID(rMapping.maFromDefault)
    , maReplacementAppCoclassIID(rMapping.maReplacementCoclass)
    , mnRefCount(0)
    , mpReplacementAppUnknown(nullptr)
    , mpReplacementAppDispatch(nullptr)
{
    std::wcout << this << L"@CProxiedApplication::CTOR" << std::endl;
    assert(!mbIsActive);
    mbIsActive = true;
}

bool CProxiedApplication::IsActive() { return mbIsActive; }

void CProxiedApplication::createReplacementAppPointers()
{
    if (mpReplacementAppUnknown == nullptr)
    {
        if (FAILED(CoCreateInstance(maReplacementAppCoclassIID, NULL, CLSCTX_LOCAL_SERVER,
                                    IID_IUnknown, (void**)&mpReplacementAppUnknown)))
        {
            std::wcerr << L"Cound not create LO instance?" << std::endl;
            std::exit(1);
        }
    }

    if (mpReplacementAppDispatch == nullptr)
    {
        if (FAILED(mpReplacementAppUnknown->QueryInterface(IID_IDispatch,
                                                           (void**)&mpReplacementAppDispatch)))
        {
            std::wcerr << L"Cound not get IDispatch from LO's interface?" << std::endl;
            std::exit(1);
        }
    }
}

// IUnknown

HRESULT STDMETHODCALLTYPE CProxiedApplication::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (IsEqualIID(riid, IID_IUnknown))
    {
        std::wcout << this << L"@CProxiedApplication::QueryInterface(IID_IUnknown): self"
                   << std::endl;
        AddRef();
        *ppvObject = this;
        return S_OK;
    }

    if (IsEqualIID(riid, IID_IClassFactory))
    {
        std::wcout << this << L"@CProxiedApplication::QueryInterface(IID_IClassFactory): self"
                   << std::endl;
        AddRef();
        *ppvObject = static_cast<IClassFactory*>(this);
        return S_OK;
    }

    createReplacementAppPointers();

    if (IsEqualIID(riid, IID_IDispatch))
    {
        std::wcout << this << L"@CProxiedApplication::QueryInterface(IID_IDispatch): self"
                   << std::endl;
        AddRef();
        *ppvObject = static_cast<IDispatch*>(this);
        return S_OK;
    }

    std::wcout << this << L"@CProxiedApplication::QueryInterface(" << riid << L")..." << std::endl;

    IDispatch* pDefault;
    std::wstring sFoundDefault;
    if (DefaultInterfaceCreator(riid, &pDefault, mpReplacementAppDispatch, sFoundDefault))
    {
        std::wcout << "..." << this << L"@CProxiedApplication::QueryInterface(" << riid
                   << L"): new " << sFoundDefault << std::endl;
        pDefault->AddRef();
        *ppvObject = pDefault;
        return S_OK;
    }

    std::wcout << "..." << this << L"@CProxiedApplication::QueryInterface(" << riid
               << L"): E_NOINTERFACE" << std::endl;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE CProxiedApplication::AddRef()
{
    std::wcout << this << L"@CProxiedApplication::AddRef: " << (mnRefCount + 1) << std::endl;
    return ++mnRefCount;
}

ULONG STDMETHODCALLTYPE CProxiedApplication::Release()
{
    std::wcout << this << L"@CProxiedApplication::Release: " << (mnRefCount - 1) << std::endl;
    ULONG nResult = --mnRefCount;
    if (nResult == 0)
    {
        mbIsActive = false;
    }
    return nResult;
}

// IClassFactory

HRESULT STDMETHODCALLTYPE CProxiedApplication::CreateInstance(IUnknown* /*pUnkOuter*/, REFIID riid,
                                                              void** ppvObject)
{
    if (IsEqualIID(riid, IID_IUnknown))
    {
        createReplacementAppPointers();

        std::wcout << this << L"@CProxiedApplication::CreateInstance(IID_IUnknown): self"
                   << std::endl;

        *ppvObject = this;
        AddRef();
        return S_OK;
    }

    std::wcout << this << L"@CProxiedApplication::CreateInstance(" << riid << L"): E_NOTIMPL"
               << std::endl;

    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CProxiedApplication::LockServer(BOOL /*fLock*/) { return S_OK; }

// IDispatch

// These return highly misleading information. We should return what we know to porxy, not what the
// replacement app says it has.

HRESULT STDMETHODCALLTYPE CProxiedApplication::GetTypeInfoCount(UINT* pctinfo)
{
    HRESULT nResult;

    std::wcout << this << L"@CProxiedApplication::GetTypeInfoCount..." << std::endl;

    nResult = mpReplacementAppDispatch->GetTypeInfoCount(pctinfo);

    std::wcout << "..." << this << L"@CProxiedApplication::GetTypeInfoCount:"
               << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedApplication::GetTypeInfo(UINT iTInfo, LCID lcid,
                                                           ITypeInfo** ppTInfo)
{
    HRESULT nResult;

    std::wcout << this << L"@CProxiedApplication::GetTypeInfo..." << std::endl;

    nResult = mpReplacementAppDispatch->GetTypeInfo(iTInfo, lcid, ppTInfo);

    std::wcout << "..." << this << L"@CProxiedApplication::GetTypeInfo: "
               << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedApplication::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames,
                                                             UINT cNames, LCID lcid,
                                                             DISPID* rgDispId)
{
    HRESULT nResult;

    std::wcout << this << L"@CProxiedApplication::GetIDsOfNames..." << std::endl;

    nResult = mpReplacementAppDispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);

    std::wcout << "..." << this << L"@CProxiedApplication::GetIDsOfNames:"
               << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedApplication::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
                                                      WORD wFlags, DISPPARAMS* pDispParams,
                                                      VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                                                      UINT* puArgErr)
{
    HRESULT nResult;

    std::wcout << this << L"@CProxiedApplication::Invoke(" << dispIdMember << L")..." << std::endl;

    nResult = mpReplacementAppDispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams,
                                               pVarResult, pExcepInfo, puArgErr);

    std::wcout << this << L"@CProxiedApplication::Invoke(" << dispIdMember << L"):"
               << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

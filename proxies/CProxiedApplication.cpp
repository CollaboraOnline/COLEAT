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

CProxiedApplication::CProxiedApplication(const InterfaceMapping& rMapping)
    : maProxiedAppCoclassIID(rMapping.maFromCoclass)
    , maProxiedAppDefaultIID(rMapping.maFromDefault)
    , maReplacementAppCoclassIID(rMapping.maTo)
    , mnRefCount(0)
    , mpReplacementAppUnknown(nullptr)
{
    std::wcout << this << L"@CProxiedApplication::CTOR" << std::endl;
    assert(!mbIsActive);
    mbIsActive = true;
}

bool CProxiedApplication::IsActive() { return mbIsActive; }

HRESULT STDMETHODCALLTYPE CProxiedApplication::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (IsEqualIID(riid, IID_IUnknown))
    {
        std::wcout << this << L"@CProxiedApplication::QueryInterface(" << riid << L"): self"
                   << std::endl;
        AddRef();
        *ppvObject = this;
        return S_OK;
    }

    if (mpReplacementAppUnknown == nullptr)
        if (FAILED(CoCreateInstance(maReplacementAppCoclassIID, NULL, CLSCTX_LOCAL_SERVER,
                                    IID_IUnknown, (void**)&mpReplacementAppUnknown)))
            return E_NOTIMPL;

    IDispatch* pReplacementAppDispatch;
    if (FAILED(mpReplacementAppUnknown->QueryInterface(IID_IDispatch,
                                                       (void**)&pReplacementAppDispatch)))
    {
        std::wcerr << L"Cound not get IDispatch from LO's interface?" << std::endl;
        std::exit(1);
    }

    std::wcout << this << L"@CProxiedApplication::QueryInterface(" << riid << L")..." << std::endl;

    IDispatch* pDefault;
    std::wstring sFoundDefault;
    if (DefaultInterfaceCreator(riid, &pDefault, pReplacementAppDispatch, sFoundDefault))
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

HRESULT STDMETHODCALLTYPE CProxiedApplication::CreateInstance(IUnknown* /*pUnkOuter*/, REFIID riid,
                                                              void** ppvObject)
{
    std::wcout << this << L"@CProxiedApplication::CreateInstance(" << riid << L")..." << std::endl;

    if (IsEqualIID(riid, IID_IUnknown))
    {
        if (mpReplacementAppUnknown == nullptr)
            if (FAILED(CoCreateInstance(maReplacementAppCoclassIID, NULL, CLSCTX_LOCAL_SERVER, riid,
                                        (void**)&mpReplacementAppUnknown)))
                return E_NOTIMPL;

        *ppvObject = this;
        AddRef();
        return S_OK;
    }

    std::wcout << L"..." << this << L"@CProxiedApplication::CreateInstance(" << riid
               << L"): E_NOTIMPL" << std::endl;

    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE CProxiedApplication::LockServer(BOOL /*fLock*/) { return S_OK; }

bool CProxiedApplication::mbIsActive = false;

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

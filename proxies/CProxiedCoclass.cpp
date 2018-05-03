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
#include <cstdlib>
#include <iostream>
#include <string>

#include <Windows.h>
#include <OCIdl.h>

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedCoclass.hpp"
#include "CProxiedDispatch.hpp"
#include "DefaultInterfaceCreator.hxx"

bool CProxiedCoclass::mbIsActive = false;

CProxiedCoclass::CProxiedCoclass(const InterfaceMapping& rMapping)
    : CProxiedDispatch(nullptr, createDispatchToProxy(rMapping), rMapping.maFromDefault)
    , maProxiedAppCoclassIID(rMapping.maFromCoclass)
    , maProxiedAppDefaultIID(rMapping.maFromDefault)
    , maReplacementAppCoclassIID(rMapping.maReplacementCoclass)
{
    // We logged in createDispatchToProxy(), which is called only above
    assert(!mbIsActive);
    mbIsActive = true;

    if (getParam()->mbTraceOnly)
        std::cout << "new " << rMapping.msFromLibName << "." << rMapping.msFromCoclassName << " -> "
                  << this << std::endl;
}

bool CProxiedCoclass::IsActive() { return mbIsActive; }

IDispatch* CProxiedCoclass::createDispatchToProxy(const InterfaceMapping& rMapping)
{
    const IID aProxiedOrReplacementIID
        = (getParam()->mbTraceOnly ? rMapping.maFromCoclass : rMapping.maReplacementCoclass);

    // If we are only tracing, there replacement app does not have to be installed, as we won't
    // actually instantiate any COM service using its CLSID. The mpReplacementAppUnknown and
    // mpReplacementAppDispatch fields in this object will actually be those of the real
    // ("proxied") application.

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedCoclass::CTOR: CoCreateInstance(" << aProxiedOrReplacementIID
                  << ")..." << std::endl;

    if (FAILED(CoCreateInstance(aProxiedOrReplacementIID, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch,
                                (void**)&mpReplacementAppDispatch)))
    {
        std::cerr << "Cound not create instance of " << aProxiedOrReplacementIID << std::endl;
        std::exit(1);
    }

#if 0
    HRESULT hr;
    ITypeInfo* pAppTI;
    if (getParam()->mbVerbose)
        std::cout << "Calling " << mpReplacementAppDispatch << "->GetTypeInfo(0)\n";
    if (FAILED((hr = mpReplacementAppDispatch->GetTypeInfo(0, LOCALE_USER_DEFAULT, &pAppTI))))
    {
        std::cerr << "GetTypeInfo failed" << std::endl;
        exit(1);
    }

    if (getParam()->mbVerbose)
        std::cout << "And calling GetTypeAttr on that\n";
    TYPEATTR* pTypeAttr;
    if (FAILED((hr = pAppTI->GetTypeAttr(&pTypeAttr))))
    {
        std::cerr << "GetTypeAttr failed" << std::endl;
        exit(1);
    }
    if (getParam()->mbVerbose)
        std::cout << "Got type attr, guid=" << pTypeAttr->guid << std::endl;
#endif
    return mpReplacementAppDispatch;
}

// IUnknown

// FIXME: I am not entirely sure why I have thought that CProxiedCoclass needs an own implementation
// of QueryInterface, and the CProxiedUnknown one wouldn't work?

HRESULT STDMETHODCALLTYPE CProxiedCoclass::QueryInterface(REFIID riid, void** ppvObject)
{
    if (!ppvObject)
        return E_POINTER;

    if (IsEqualIID(riid, IID_IUnknown))
    {
        if (getParam()->mbVerbose)
            std::cout << this << "@CProxiedCoclass::QueryInterface(IID_IUnknown): self"
                      << std::endl;
        AddRef();
        *ppvObject = this;
        return S_OK;
    }

    if (IsEqualIID(riid, IID_IDispatch))
    {
        if (getParam()->mbVerbose)
            std::cout << this << "@CProxiedCoclass::QueryInterface(IID_IDispatch): self"
                      << std::endl;
        AddRef();
        *ppvObject = reinterpret_cast<IDispatch*>(this);
        return S_OK;
    }

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedCoclass::QueryInterface(" << riid << ")..." << std::endl;

    IDispatch* pDefault;
    std::string sFoundDefault;
    if (DefaultInterfaceCreator(this, riid, &pDefault, mpReplacementAppDispatch, sFoundDefault))
    {
        if (getParam()->mbVerbose)
            std::cout << "..." << this << "@CProxiedCoclass::QueryInterface(" << riid << "): new "
                      << sFoundDefault << ": " << pDefault << std::endl;
        pDefault->AddRef();
        *ppvObject = pDefault;
        return S_OK;
    }

    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedCoclass::QueryInterface(" << riid
                  << "): E_NOINTERFACE" << std::endl;
    return E_NOINTERFACE;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

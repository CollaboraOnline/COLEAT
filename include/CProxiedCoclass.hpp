/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_CPROXIEDCOCLASS_HPP
#define INCLUDED_CPROXIEDCOCLASS_HPP

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <iostream>
#include <string>

#include <Windows.h>

#include "exewrapper.hpp"
#include "interfacemap.hpp"

#pragma warning(pop)

class CProxiedCoclass : public IClassFactory, public IDispatch
{
private:
    const IID maProxiedAppCoclassIID;
    const IID maProxiedAppDefaultIID;
    const IID maReplacementAppCoclassIID;
    static bool mbIsActive;
    ULONG mnRefCount;
    IUnknown* mpReplacementAppUnknown;
    IDispatch* mpReplacementAppDispatch;

    void createReplacementAppPointers();

public:
    CProxiedCoclass(const InterfaceMapping& rMapping);

    static bool IsActive();

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;

    ULONG STDMETHODCALLTYPE AddRef() override;

    ULONG STDMETHODCALLTYPE Release() override;

    // IClassFactory
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* pUnkOuter, REFIID riid,
                                             void** ppvObject) override;
    
    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) override;

    // IDispatch
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override;

    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override;

    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
                                            LCID lcid, DISPID* rgDispId) override;

    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                                     DISPPARAMS* pDispParams, VARIANT* pVarResult,
                                     EXCEPINFO* pExcepInfo, UINT* puArgErr) override;
};

#endif // INCLUDED_CPROXIEDCOCLASS_HPP

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

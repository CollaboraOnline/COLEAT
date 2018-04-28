/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_CProxiedSink_hpp
#define INCLUDED_CProxiedSink_hpp

#pragma warning(push)
#pragma warning(disable: 4668 4820 4917)

#include <Windows.h>

#pragma warning(pop)

#include "CProxiedUnknown.hpp"

// We have all these proxy classes inherit from CProxiedUnknown and not from the corresponding
// actual abstract interface class so that they will automatically inherit CProxiedUnknown's
// implementations of the IUnknown member functions. We must therefore declare the member functions
// in the same order as those of the actual abstract interface class as virtual (and not overridden)
// so that they go into the vtbl in the expected order.

class CProxiedSink : public CProxiedUnknown
{
private:
    IDispatch* mpDispatchToProxy;
    ITypeInfo* mpTypeInfoOfOutgoingInterface;

public:
    CProxiedSink(IDispatch* pDispatchToProxy, ITypeInfo* pTypeInfoOfOutgoingInterface, const IID& aOutgoingIID);

    // IDispatch

    virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo);

    virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);

    virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
                                                    LCID lcid, DISPID* rgDispId);

    virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                                             DISPPARAMS* pDispParams, VARIANT* pVarResult,
                                             EXCEPINFO* pExcepInfo, UINT* puArgErr);
};

#endif // INCLUDED_CProxiedSink_hpp

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

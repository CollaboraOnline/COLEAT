/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_CProxiedDispatch_hpp
#define INCLUDED_CProxiedDispatch_hpp

#pragma warning(push)
#pragma warning(disable: 4668 4820 4917)

#include <string>
#include <vector>

#include <Windows.h>
#include <OCIdl.h>

#pragma warning(pop)

#include "CProxiedUnknown.hpp"

class CProxiedDispatch: public CProxiedUnknown
{
public:
    CProxiedDispatch(IUnknown* pBaseClassUnknown,
                     IDispatch* pDispatchToProxy);

    CProxiedDispatch(IUnknown* pBaseClassUnknown,
                     IDispatch* pDispatchToProxy,
                     const IID& rIID);

    CProxiedDispatch(IUnknown* pBaseClassUnknown,
                     IDispatch* pDispatchToProxy,
                     const IID& rIID1,
                     const IID& rIID2);

private:
    IDispatch* const mpDispatchToProxy;

public:
    HRESULT genericInvoke(std::string sFuncName,
                          int nInvKind,
                          std::vector<VARIANT>& rParameters,
                          void* pRetval);

    // IDispatch
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo);

    virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);

    virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
                                                    LCID lcid, DISPID* rgDispId);

    virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                                             DISPPARAMS* pDispParams, VARIANT* pVarResult,
                                             EXCEPINFO* pExcepInfo, UINT* puArgErr);
};

#endif // INCLUDED_CProxiedDispatch_hpp

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

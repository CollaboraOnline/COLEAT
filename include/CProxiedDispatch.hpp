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
#pragma warning(disable : 4668 4820 4917)

#include <string>
#include <vector>

#include <Windows.h>
#include <OCIdl.h>

#pragma warning(pop)

#include "CProxiedUnknown.hpp"

class CProxiedDispatch : public CProxiedUnknown
{
private:
    IDispatch* const mpDispatchToProxy;

    // Cached results from GetIDsOfNames() calls, to use for logging in case no type information is
    // available.
    std::map<DISPID, std::string>* mpDispIdToName;

protected:
    CProxiedDispatch(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                     const char* sLibName);
    CProxiedDispatch(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy, const IID& rIID,
                     const char* sLibName);
    CProxiedDispatch(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy, const IID& rIID1,
                     const IID& rIID2, const char* sLibName);

#if 0
    // We can't have a non-virtual destructor because of the "class has virtual functions, but
    // destructor is not virtual" problem and we can't make the destructor virtual because our
    // vtable should have *only* the entries from IUnknown plus the ones from this class
    // (corresponding to the entries for IDispatch). So ideally we should delete mpDispIdToName in
    // Release() when the reference count reaches zero.

    ~CProxiedDispatch()
    {
        delete mpDispIdToName;
    }

#endif

public:
    static CProxiedDispatch* get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                 const char* sLibName);
    static CProxiedDispatch* get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                 const IID& rIID, const char* sLibName);
    static CProxiedDispatch* get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                 const IID& rIID1, const IID& rIID2, const char* sLibName);

    HRESULT genericInvoke(const std::string& rFuncName, int nInvKind,
                          std::vector<VARIANT>& rParameters, void* pRetval);

    // IDispatch
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo);

    virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);

    virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
                                                    LCID lcid, DISPID* rgDispId);

    virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
                                             WORD wFlags, DISPPARAMS* pDispParams,
                                             VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                                             UINT* puArgErr);
};

#endif // INCLUDED_CProxiedDispatch_hpp

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

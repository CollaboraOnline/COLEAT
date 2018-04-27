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

#include <map>
#include <string>
#include <vector>

#include <Windows.h>
#include <OCIdl.h>

#pragma warning(pop)

#include "exewrapper.hpp"

inline bool operator<(const IID& a, const IID& b) { return std::memcmp(&a, &b, sizeof(a)) < 0; }

struct ConnectionPointMapHolder
{
    std::map<IID, IConnectionPoint*> maConnectionPoints;
};

class CProxiedDispatch: public IDispatch, public IConnectionPointContainer
{
public:
    CProxiedDispatch(IDispatch* pDispatchToProxy,
                     const IID& aIID);
    CProxiedDispatch(IDispatch* pDispatchToProxy,
                     const IID& aIID,
                     const IID& aCoclassIID);

    static void setParam(ThreadProcParam* pParam);
    static ThreadProcParam* getParam();

private:
    IID maIID, maCoclassIID;
    IDispatch* mpDispatchToProxy;
    IConnectionPointContainer* mpConnectionPointContainerToProxy;
    ConnectionPointMapHolder* const mpCPMap;

public:
    HRESULT Invoke(std::string sFuncName,
                   int nInvKind,
                   std::vector<VARIANT>& rParameters,
                   void* pRetval);

    HRESULT FindTypeInfo(const IID& rIID, ITypeInfo** pTypeInfo);

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;
    // IDispatch
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override;
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override;
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
                                            LCID lcid, DISPID* rgDispId) override;
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                                     DISPPARAMS* pDispParams, VARIANT* pVarResult,
                                     EXCEPINFO* pExcepInfo, UINT* puArgErr) override;
    // IConnectionPointContainer
    HRESULT STDMETHODCALLTYPE EnumConnectionPoints(IEnumConnectionPoints **ppEnum) override;
    HRESULT STDMETHODCALLTYPE FindConnectionPoint(REFIID riid, IConnectionPoint **ppCP) override;
};

#endif // INCLUDED_CProxiedDispatch_hpp

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

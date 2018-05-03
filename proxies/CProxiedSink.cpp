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
#include <iostream>

#include <Windows.h>

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedSink.hpp"

#include "CallbackInvoker.hxx"

std::map<IUnknown*, IUnknown*> CProxiedSink::maActiveSinks;

IUnknown* CProxiedSink::existingSink(IUnknown* pUnk)
{
    auto p = maActiveSinks.find(pUnk);
    if (p == maActiveSinks.end())
        return NULL;
    return p->second;
}

void CProxiedSink::forgetExistingSink(IUnknown* pUnk) { maActiveSinks.erase(pUnk); }

CProxiedSink::CProxiedSink(IDispatch* pDispatchToProxy, ITypeInfo* pTypeInfoOfOutgoingInterface,
                           const OutgoingInterfaceMapping& rMapEntry, const IID& aOutgoingIID)
    : CProxiedUnknown(pDispatchToProxy, IID_IDispatch, aOutgoingIID)
    , mpDispatchToProxy(pDispatchToProxy)
    , mpTypeInfoOfOutgoingInterface(pTypeInfoOfOutgoingInterface)
    , maMapEntry(rMapEntry)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedSink::CTOR" << std::endl;
    maActiveSinks[this] = pDispatchToProxy;
}

HRESULT STDMETHODCALLTYPE CProxiedSink::GetTypeInfoCount(UINT* pctinfo)
{
    if (!pctinfo)
        return E_POINTER;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedSink::GetTypeInfoCount..." << std::endl;
    *pctinfo = 0;
    HRESULT nResult = mpDispatchToProxy->GetTypeInfoCount(pctinfo);
    if (getParam()->mbVerbose)
    {
        std::cout << "..." << this
                  << "@CProxiedSink::GetTypeInfoCount: " << WindowsErrorStringFromHRESULT(nResult);
        if (nResult == S_OK)
            std::cout << ": " << *pctinfo;
        std::cout << std::endl;
    }
    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedSink::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    if (!ppTInfo)
        return E_POINTER;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedSink::GetTypeInfo..." << std::endl;
    HRESULT nResult = mpDispatchToProxy->GetTypeInfo(iTInfo, lcid, ppTInfo);
    if (getParam()->mbVerbose)
        std::cout << "..." << this
                  << "@CProxiedSink::GetTypeInfo: " << WindowsErrorStringFromHRESULT(nResult)
                  << std::endl;
    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedSink::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
                                                      LCID lcid, DISPID* rgDispId)
{
    if (*rgDispId)
        return E_POINTER;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedSink::GetIDsOfNames..." << std::endl;
    HRESULT nResult = mpDispatchToProxy->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
    if (getParam()->mbVerbose)
        std::cout << "..." << this
                  << "@CProxiedSink::GetIDsOfNames: " << WindowsErrorStringFromHRESULT(nResult)
                  << std::endl;
    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedSink::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
                                               WORD wFlags, DISPPARAMS* pDispParams,
                                               VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                                               UINT* puArgErr)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedSink::Invoke(" << dispIdMember << ")..." << std::endl;

    DISPID nDispIdMemberInClient;
    if (mpTypeInfoOfOutgoingInterface != NULL)
    {
        // The "normal" mode, where the client we are tracing is connected to the replacement
        // application, whose outgoing interface is not 1:1 equivalent to that of the original
        // application that provided the type library the client was built against. So we must map
        // the member ids.
        UINT nNames;
        BSTR sName = NULL;
        nResult = mpTypeInfoOfOutgoingInterface->GetNames(dispIdMember, &sName, 1, &nNames);
        if (FAILED(nResult))
        {
            if (getParam()->mbVerbose)
                std::cout << "..." << this << "@CProxiedSink::Invoke(" << dispIdMember
                          << "): GetNames failed: " << WindowsErrorStringFromHRESULT(nResult)
                          << std::endl;
            return nResult;
        }

        nResult
            = mpDispatchToProxy->GetIDsOfNames(IID_NULL, &sName, 1, lcid, &nDispIdMemberInClient);
        if (FAILED(nResult))
        {
            if (getParam()->mbVerbose)
                std::cout << "..." << this << "@CProxiedSink::Invoke(" << dispIdMember
                          << "): GetIDsOfNames(" << convertUTF16ToUTF8(sName)
                          << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
            SysFreeString(sName);
            return nResult;
        }
        SysFreeString(sName);
    }
    else
    {
        // mpTypeInfoOfOutgoingInterface can be NULL only when we are in tracing-only mode, i.e.
        // when the client we are tracing is connected to the very application that provided the
        // type information the client was built against. Then we can use the dispIdMember parameter
        // to this function directly also the dispIdMember passed on to when invoking the client
        // callback.
        assert(getParam()->mbTraceOnly);
        nDispIdMemberInClient = dispIdMember;
    }

    // maIID1 is IID_IDispatch (see ctor above), maIID2 is the IID of the outgoing interface.
    nResult = ProxiedCallbackInvoke(maIID2, mpDispatchToProxy, nDispIdMemberInClient, riid, lcid,
                                    wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedSink::Invoke(" << dispIdMember
                  << "): maps to Invoke(" << nDispIdMemberInClient
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

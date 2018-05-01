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
#include <codecvt>
#include <iostream>

#include <Windows.h>

#include <comphelper/windowsdebugoutput.hxx>

#pragma warning(pop)

#include "CProxiedSink.hpp"
#include "utilstemp.hpp"

#include "CallbackInvoker.hxx"

std::map<IUnknown*, IUnknown*> CProxiedSink::maActiveSinks;

IUnknown* CProxiedSink::existingSink(IUnknown* pUnk)
{
    auto p = maActiveSinks.find(pUnk);
    if (p == maActiveSinks.end())
        return NULL;
    std::cout << "=== found " << pUnk << ": " << p->second << std::endl;
    return p->second;
}

void CProxiedSink::forgetExistingSink(IUnknown* pUnk)
{
    std::cout << "=== forgetting " << pUnk << std::endl;
    maActiveSinks.erase(pUnk);
}

CProxiedSink::CProxiedSink(IDispatch* pDispatchToProxy, ITypeInfo* pTypeInfoOfOutgoingInterface,
                           const OutgoingInterfaceMapping& rMapEntry, const IID& aOutgoingIID)
    : CProxiedUnknown(pDispatchToProxy, IID_IDispatch, aOutgoingIID)
    , mpDispatchToProxy(pDispatchToProxy)
    , mpTypeInfoOfOutgoingInterface(pTypeInfoOfOutgoingInterface)
    , maMapEntry(rMapEntry)
{
    std::cout << this << "@CProxiedSink::CTOR" << std::endl;
    maActiveSinks[this] = pDispatchToProxy;
    std::cout << "=== registered " << this << std::endl;
}

HRESULT STDMETHODCALLTYPE CProxiedSink::GetTypeInfoCount(UINT* pctinfo)
{
    if (!pctinfo)
        return E_POINTER;

    std::cout << this << "@CProxiedSink::GetTypeInfoCount..." << std::endl;
    *pctinfo = 0;
    HRESULT nResult = mpDispatchToProxy->GetTypeInfoCount(pctinfo);
    std::cout << "..." << this
              << "@CProxiedSink::GetTypeInfoCount: " << WindowsErrorStringFromHRESULT(nResult);
    if (nResult == S_OK)
        std::cout << ": " << *pctinfo;
    std::cout << std::endl;
    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedSink::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    if (!ppTInfo)
        return E_POINTER;

    std::cout << this << "@CProxiedSink::GetTypeInfo..." << std::endl;
    HRESULT nResult = mpDispatchToProxy->GetTypeInfo(iTInfo, lcid, ppTInfo);
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

    std::cout << this << "@CProxiedSink::GetIDsOfNames..." << std::endl;
    HRESULT nResult = mpDispatchToProxy->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
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

    std::cout << this << "@CProxiedSink::Invoke(" << dispIdMember << ")..." << std::endl;

    BSTR sName = NULL;
    UINT nNames;
    if (mpTypeInfoOfOutgoingInterface != NULL)
    {
        std::cout << "========== mpTypeInfoOfOutgoingInterface is not null\n";
        nResult = mpTypeInfoOfOutgoingInterface->GetNames(dispIdMember, &sName, 1, &nNames);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedSink::Invoke(" << dispIdMember
                      << "): GetNames failed: " << WindowsErrorStringFromHRESULT(nResult)
                      << std::endl;
            return nResult;
        }
    }
    else
    {
        std::cout << "========== mpTypeInfoOfOutgoingInterface IS NULL\n";
        bool bFound = false;
        for (int i = 0; maMapEntry.maNameToId[i].mpName != nullptr; ++i)
        {
            if (maMapEntry.maNameToId[i].mnMemberId == dispIdMember)
            {
                sName = SysAllocString(
                    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>, wchar_t>()
                        .from_bytes(maMapEntry.maNameToId[i].mpName)
                        .data());
                std::cout << "========== found: " << maMapEntry.maNameToId[i].mpName << "\n";
                bFound = true;
                break;
            }
        }
        if (!bFound)
        {
            std::cout << "..." << this << "@CProxiedSink::Invoke(" << dispIdMember
                      << "): GetIDsOfNames("
                      << std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t>().to_bytes(sName)
                      << ") failed: Not found in outgoing mapping" << std::endl;
            SysFreeString(sName);
            return E_NOTIMPL;
        }
    }

    DISPID nDispIdMemberInClient;
    nResult = mpDispatchToProxy->GetIDsOfNames(IID_NULL, &sName, 1, lcid, &nDispIdMemberInClient);
    if (FAILED(nResult))
    {
        std::cout << "..." << this << "@CProxiedSink::Invoke(" << dispIdMember
                  << "): GetIDsOfNames("
                  << std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t>().to_bytes(sName)
                  << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        SysFreeString(sName);
        return nResult;
    }

    SysFreeString(sName);

    // maIID1 is IID_IDispatch (see ctor aboce), maIID2 is the IID of the outgoing interface.
    nResult = ProxiedCallbackInvoke(maIID2, mpDispatchToProxy, nDispIdMemberInClient, riid, lcid,
                                    wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

    std::cout << "..." << this << "@CProxiedSink::Invoke(" << dispIdMember << "): maps to Invoke("
              << nDispIdMemberInClient << "): " << WindowsErrorStringFromHRESULT(nResult)
              << std::endl;

    return nResult;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

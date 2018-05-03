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

#include <cstdlib>
#include <iostream>
#include <string>
#include <vector>

#include <Windows.h>
#include <OleCtl.h>

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedDispatch.hpp"

CProxiedDispatch::CProxiedDispatch(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy)
    : CProxiedUnknown(pBaseClassUnknown, pDispatchToProxy, IID_IDispatch)
    , mpDispatchToProxy(pDispatchToProxy)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::CTOR(" << pBaseClassUnknown << ", "
                  << pDispatchToProxy << ")" << std::endl;
}

CProxiedDispatch::CProxiedDispatch(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                   const IID& rIID)
    : CProxiedDispatch(pBaseClassUnknown, pDispatchToProxy, IID_IDispatch, rIID)
{
}

CProxiedDispatch::CProxiedDispatch(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                   const IID& rIID1, const IID& rIID2)
    : CProxiedUnknown(pBaseClassUnknown, pDispatchToProxy, rIID1, rIID2)
    , mpDispatchToProxy(pDispatchToProxy)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::CTOR(" << pBaseClassUnknown << ", "
                  << pDispatchToProxy << ", " << rIID1 << ", " << rIID2 << ")" << std::endl;
}

HRESULT CProxiedDispatch::genericInvoke(std::string sFuncName, int nInvKind,
                                        std::vector<VARIANT>& rParameters, void* pRetval)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::genericInvoke(" << sFuncName << ")..."
                  << std::endl;

    HRESULT nResult = S_OK;

    std::wstring sFuncNameWstr = convertUTF8ToUTF16(sFuncName.c_str());
    LPOLESTR pFuncNameWstr = (LPOLESTR)sFuncNameWstr.data();

    MEMBERID nMemberId;

    nResult = mpDispatchToProxy->GetIDsOfNames(IID_NULL, &pFuncNameWstr, 1, LOCALE_USER_DEFAULT,
                                               &nMemberId);
    if (nResult == DISP_E_UNKNOWNNAME)
    {
        if (getParam()->mbVerbose)
            std::cerr << this << "@CProxiedDispatch::genericInvoke(" << sFuncName
                      << "): Not implemented in the replacement app" << std::endl;
        return E_NOTIMPL;
    }

    if (FAILED(nResult))
    {
        if (getParam()->mbVerbose)
            std::cerr << this << "@CProxiedDispatch::genericInvoke(" << sFuncName
                      << "): GetIDsOfNames failed: " << WindowsErrorStringFromHRESULT(nResult)
                      << std::endl;
        return nResult;
    }

    WORD nFlags;
    switch (nInvKind)
    {
        case INVOKE_FUNC:
            nFlags = DISPATCH_METHOD;
            break;
        case INVOKE_PROPERTYGET:
            nFlags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;
            break;
        case INVOKE_PROPERTYPUT:
            nFlags = DISPATCH_PROPERTYPUT;
            break;
        case INVOKE_PROPERTYPUTREF:
            nFlags = DISPATCH_PROPERTYPUTREF;
            break;
        default:
            std::cerr << this << "@CProxiedDispatch::genericInvoke(" << sFuncName
                      << "): Unhandled nInvKind: " << nInvKind << std::endl;
            std::abort();
    }

    DISPPARAMS aDispParams;
    aDispParams.rgvarg = rParameters.data();
    aDispParams.rgdispidNamedArgs = NULL;
    aDispParams.cArgs = (UINT)rParameters.size();
    aDispParams.cNamedArgs = 0;

    VARIANT aResult;
    VariantInit(&aResult);

    UINT nArgErr;

    // See https://msdn.microsoft.com/en-us/library/windows/desktop/ms221486(v=vs.85).aspx
    DISPID nDispid = DISPID_PROPERTYPUT;
    if (nFlags == DISPATCH_PROPERTYPUT)
    {
        aDispParams.rgdispidNamedArgs = &nDispid;
        aDispParams.cNamedArgs = 1;
    }

    nResult = mpDispatchToProxy->Invoke(nMemberId, IID_NULL, LOCALE_USER_DEFAULT, nFlags,
                                        &aDispParams, &aResult, NULL, &nArgErr);
    if (FAILED(nResult))
    {
        if (getParam()->mbVerbose)
            std::cout << "..." << this << "@CProxiedDispatch::genericInvoke(" << sFuncName
                      << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }

    if (aResult.vt != VT_EMPTY && aResult.vt != VT_VOID && pRetval != nullptr)
    {
        switch (aResult.vt)
        {
            case VT_BSTR:
                *(BSTR*)pRetval = aResult.bstrVal;
                break;
            case VT_DISPATCH:
                *(IDispatch**)pRetval = aResult.pdispVal;
                break;
            default:
                std::cerr << this << "@CProxiedDispatch::genericInvoke(" << sFuncName
                          << "): Unhandled vt: " << aResult.vt << std::endl;
                std::abort();
        }
    }
    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedDispatch::genericInvoke(" << sFuncName << "): S_OK"
                  << std::endl;

    return S_OK;
}

// IDispatch

HRESULT STDMETHODCALLTYPE CProxiedDispatch::GetTypeInfoCount(UINT* pctinfo)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::GetTypeInfoCount..." << std::endl;

    nResult = mpDispatchToProxy->GetTypeInfoCount(pctinfo);

    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedDispatch::GetTypeInfoCount:"
                  << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedDispatch::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::GetTypeInfo(" << iTInfo << ")..." << std::endl;

    nResult = mpDispatchToProxy->GetTypeInfo(iTInfo, lcid, ppTInfo);

    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedDispatch::GetTypeInfo(" << iTInfo
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedDispatch::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames,
                                                          UINT cNames, LCID lcid, DISPID* rgDispId)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::GetIDsOfNames..." << std::endl;

    nResult = mpDispatchToProxy->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);

    if (getParam()->mbVerbose)
        std::cout << "..." << this
                  << "@CProxiedDispatch::GetIDsOfNames:" << WindowsErrorStringFromHRESULT(nResult)
                  << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedDispatch::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
                                                   WORD wFlags, DISPPARAMS* pDispParams,
                                                   VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                                                   UINT* puArgErr)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::Invoke(" << dispIdMember << ")..." << std::endl;

    nResult = mpDispatchToProxy->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult,
                                        pExcepInfo, puArgErr);

    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedDispatch::Invoke(" << dispIdMember
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

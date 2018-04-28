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

#include <Windows.h>

#include <comphelper/windowsdebugoutput.hxx>

#pragma warning(pop)

#include "CProxiedConnectionPointContainer.hpp"
#include "CProxiedDispatch.hpp"
#include "CProxiedEnumConnectionPoints.hpp"
#include "CProxiedEnumConnections.hpp"
#include "CProxiedSink.hpp"
#include "CProxiedUnknown.hpp"
#include "utilstemp.hpp"

static ThreadProcParam* pGlobalParam;

HRESULT CProxiedUnknown::findTypeInfo(IDispatch* pDispatch, const IID& rIID, ITypeInfo** pTypeInfo)
{
    std::cout << this << "@CProxiedUnknown::findTypeInfo(" << pDispatch << ", " << rIID << ")..."
              << std::endl;
    HRESULT nResult;

    ITypeInfo* pTI;
    std::cout << "Calling " << pDispatch << "->GetTypeInfo(0)\n";
    nResult = pDispatch->GetTypeInfo(0, LOCALE_USER_DEFAULT, &pTI);
    if (FAILED(nResult))
    {
        std::cout << "..." << this << "@CProxiedUnknown::findTypeInfo(" << rIID << ") (" << __LINE__
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }

    std::cout << "And calling GetTypeAttr on that\n";
    TYPEATTR* pTA;
    nResult = pTI->GetTypeAttr(&pTA);
    if (FAILED(nResult))
    {
        std::cout << "..." << this << "@CProxiedUnknown::findTypeInfo(" << rIID << ") (" << __LINE__
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }
    std::cout << "Got type attr, guid=" << pTA->guid << std::endl;

    bool bFound = false;
    for (WORD i = 0; i < pTA->cImplTypes; ++i)
    {
        HREFTYPE nHrefType;
        nResult = pTI->GetRefTypeOfImplType(i, &nHrefType);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedUnknown::findTypeInfo(" << rIID << ") ("
                      << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
            return nResult;
        }

        nResult = pTI->GetRefTypeInfo(nHrefType, pTypeInfo);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedUnknown::findTypeInfo(" << rIID << ") ("
                      << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
            return nResult;
        }

        TYPEATTR* pRefTA;
        nResult = (*pTypeInfo)->GetTypeAttr(&pRefTA);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedUnknown::findTypeInfo(" << rIID << ") ("
                      << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
            return nResult;
        }

        std::cout << "..." << this << "@CProxiedUnknown::findTypeInfo(" << rIID << ") (" << __LINE__
                  << "): pRefTA=" << pRefTA->guid << "\n";

        if (IsEqualIID(rIID, pRefTA->guid))
        {
            (*pTypeInfo)->ReleaseTypeAttr(pRefTA);
            bFound = true;
            break;
        }
        (*pTypeInfo)->ReleaseTypeAttr(pRefTA);
    }
    pTI->ReleaseTypeAttr(pTA);

    if (!bFound)
    {
        std::cout << "..." << this << "@Cproxiedunknown::findTypeInfo(" << rIID << ") (" << __LINE__
                  << "): E_NOTIMPL" << std::endl;
        return E_NOTIMPL;
    }

    std::cout << "..." << this << "@Cproxiedunknown::findTypeInfo(" << rIID << "): S_OK"
              << std::endl;
    return S_OK;
}

CProxiedUnknown::CProxiedUnknown(IUnknown* pBaseClassUnknown, IUnknown* pUnknownToProxy,
                                 const IID& rIID)
    : CProxiedUnknown(pBaseClassUnknown, pUnknownToProxy, rIID, IID_NULL)
{
}

CProxiedUnknown::CProxiedUnknown(IUnknown* pUnknownToProxy, const IID& rIID)
    : CProxiedUnknown(nullptr, pUnknownToProxy, rIID)
{
}

CProxiedUnknown::CProxiedUnknown(IUnknown* pBaseClassUnknown, IUnknown* pUnknownToProxy,
                                 const IID& rIID1, const IID& rIID2)
    : mpBaseClassUnknown(pBaseClassUnknown)
    , maIID1(rIID1)
    , maIID2(rIID2)
    , mpUnknownToProxy(pUnknownToProxy)
    // FIXME: Must delete mpUnknownMap when Release() returns zero?
    , mpExtraInterfaces(new UnknownMapHolder())
{
    std::cout << this << "@CProxiedUnknown::CTOR(" << pBaseClassUnknown << ", " << pUnknownToProxy
              << ", " << rIID1 << ", " << rIID2 << ")" << std::endl;
}

CProxiedUnknown::CProxiedUnknown(IUnknown* pUnknownToProxy, const IID& rIID1, const IID& rIID2)
    : CProxiedUnknown(nullptr, pUnknownToProxy, rIID1, rIID2)
{
}

void CProxiedUnknown::setParam(ThreadProcParam* pParam) { pGlobalParam = pParam; }

ThreadProcParam* CProxiedUnknown::getParam() { return pGlobalParam; }

// IUnknown

HRESULT STDMETHODCALLTYPE CProxiedUnknown::QueryInterface(REFIID riid, void** ppvObject)
{
    HRESULT nResult;

    if (!ppvObject)
        return E_POINTER;

    if (IsEqualIID(riid, IID_IUnknown))
    {
        if (mpBaseClassUnknown)
        {
            std::cout << this << "@CProxiedUnknown::QueryInterface(IID_IUnknown): base: "
                      << mpBaseClassUnknown << std::endl;
            *ppvObject = mpBaseClassUnknown;
        }
        else
        {
            std::cout << this << "@CProxiedUnknown::QueryInterface(IID_IUnknown): self: " << this
                      << std::endl;
            *ppvObject = this;
        }
        AddRef();
        return S_OK;
    }

    if (!IsEqualIID(maIID1, IID_NULL) && IsEqualIID(riid, maIID1))
    {
        std::cout << this << "@CProxiedUnknown::QueryInterface(" << riid << "): self: " << this
                  << std::endl;
        *ppvObject = this;
        AddRef();
        return S_OK;
    }

    if (!IsEqualIID(maIID2, IID_NULL) && IsEqualIID(riid, maIID2))
    {
        std::cout << this << "@CProxiedUnknown::QueryInterface(" << riid << "): self: " << this
                  << std::endl;
        *ppvObject = this;
        AddRef();
        return S_OK;
    }

    std::cout << this << "@CProxiedUnknown::QueryInterface(" << riid << ")..." << std::endl;
    nResult = mpUnknownToProxy->QueryInterface(riid, ppvObject);
    std::cout << "..." << this << "@CProxiedUnknown::QueryInterface(" << riid
              << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    // Special cases

    if (nResult == S_OK && IsEqualIID(riid, IID_IDispatch))
    {
        *ppvObject = new CProxiedDispatch(mpBaseClassUnknown ? mpBaseClassUnknown : this,
                                          (IDispatch*)*ppvObject);
        return S_OK;
    }

    if (nResult == S_OK && IsEqualIID(riid, IID_IConnectionPointContainer))
    {
        // We must proxy it so that we can reverse-proxy the calls to advised sinks

        IProvideClassInfo* pPCI;
        nResult = mpUnknownToProxy->QueryInterface(IID_IProvideClassInfo, (void**)&pPCI);
        if (nResult != S_OK)
        {
            std::cerr << "Have IConnectionPointContainer but not IProvideClassInfo?" << std::endl;
            ((IUnknown*)*ppvObject)->Release();
            return E_NOINTERFACE;
        }
        *ppvObject
            = new CProxiedConnectionPointContainer(mpBaseClassUnknown ? mpBaseClassUnknown : this,
                                                   (IConnectionPointContainer*)*ppvObject, pPCI);

        return S_OK;
    }

    return nResult;
}

ULONG STDMETHODCALLTYPE CProxiedUnknown::AddRef()
{
    ULONG nRetval = mpUnknownToProxy->AddRef();
    std::cout << this << "@CProxiedUnknown::AddRef: " << nRetval << std::endl;

    return nRetval;
}

ULONG STDMETHODCALLTYPE CProxiedUnknown::Release()
{
    ULONG nRetval = mpUnknownToProxy->Release();
    std::cout << this << "@CProxiedUnknown::Release: " << nRetval << std::endl;

    return nRetval;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

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

#include <iostream>

#include <Windows.h>

#include <comphelper/windowsdebugoutput.hxx>

#pragma warning(pop)

#include "CProxiedConnectionPoint.hpp"
#include "CProxiedEnumConnections.hpp"
#include "CProxiedSink.hpp"
#include "utilstemp.hpp"

CProxiedConnectionPoint::CProxiedConnectionPoint(IUnknown* pBaseClassUnknown,
                                                 CProxiedConnectionPointContainer* pContainer,
                                                 IConnectionPoint* pCPToProxy, IID aIID,
                                                 ITypeInfo* pTypeInfoOfOutgoingInterface)
    : CProxiedUnknown(pBaseClassUnknown, pCPToProxy, IID_IConnectionPoint)
    , mpContainer(pContainer)
    , mpCPToProxy(pCPToProxy)
    , maIID(aIID)
    , mpTypeInfoOfOutgoingInterface(pTypeInfoOfOutgoingInterface)
    // FIXME: Delete this when we go inactive
    , mpAdvisedSinks(new AdvisedSinkHolder())
{
    std::cout << this << "@CProxiedConnectionPoint::CTOR" << std::endl;
}

HRESULT STDMETHODCALLTYPE CProxiedConnectionPoint::GetConnectionInterface(IID* /*pIID*/)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE
CProxiedConnectionPoint::GetConnectionPointContainer(IConnectionPointContainer** ppCPC)
{
    if (!ppCPC)
        return E_POINTER;

    *ppCPC = reinterpret_cast<IConnectionPointContainer*>(mpContainer);

    return S_OK;
}

HRESULT STDMETHODCALLTYPE CProxiedConnectionPoint::Advise(IUnknown* pUnkSink, DWORD* pdwCookie)
{
    HRESULT nResult;

    if (!pdwCookie)
        return E_POINTER;

    std::cout << this << "@CProxiedConnectionPoint::Advise(" << pUnkSink << ")..." << std::endl;

    IDispatch* pSinkAsDispatch;
    nResult = pUnkSink->QueryInterface(IID_IDispatch, (void**)&pSinkAsDispatch);
    if (FAILED(nResult))
    {
        std::cerr << "..." << this << "@CProxiedSink::Advise: Sink is not an IDispatch"
                  << std::endl;
        return E_NOTIMPL;
    }

    IDispatch* pDispatch = reinterpret_cast<IDispatch*>(
        new CProxiedSink(pSinkAsDispatch, mpTypeInfoOfOutgoingInterface, maIID));

    *pdwCookie = 0;
    nResult = mpCPToProxy->Advise(pDispatch, pdwCookie);

    if (FAILED(nResult))
    {
        std::cout << "..." << this << "@CProxiedConnectionPoint::Advise(" << pUnkSink
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        pSinkAsDispatch->Release();
        delete pDispatch;
        return nResult;
    }

    std::cout << "..." << this << "@CProxiedConnectionPoint::Advise(" << pUnkSink
              << "): " << *pdwCookie << ": S_OK" << std::endl;

    mpAdvisedSinks->maAdvisedSinks[*pdwCookie] = pDispatch;

    return S_OK;
}

HRESULT STDMETHODCALLTYPE CProxiedConnectionPoint::Unadvise(DWORD dwCookie)
{
    HRESULT nResult;

    std::cout << this << "@CProxiedConnectionPoint::Unadvise(" << dwCookie << ")..." << std::endl;

    nResult = mpCPToProxy->Unadvise(dwCookie);
    if (mpAdvisedSinks->maAdvisedSinks.count(dwCookie) == 0)
    {
        std::cout << "..." << this << "@CProxiedConnectionPoint::Unadvise(" << dwCookie
                  << "): E_POINTER" << std::endl;
        return E_POINTER;
    }
    delete mpAdvisedSinks->maAdvisedSinks[dwCookie];
    mpAdvisedSinks->maAdvisedSinks.erase(dwCookie);

    if (FAILED(nResult))
    {
        std::cout << "..." << this << "@CProxiedConnectionPoint::Unadvise(" << dwCookie
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }

    std::cout << "..." << this << "@CProxiedConnectionPoint::Unadvise(" << dwCookie << "): S_OK"
              << std::endl;

    return S_OK;
}

HRESULT STDMETHODCALLTYPE CProxiedConnectionPoint::EnumConnections(IEnumConnections** ppEnum)
{
    HRESULT nResult;

    if (!ppEnum)
        return E_POINTER;

    std::cout << this << "@CProxiedConnectionPoint::EnumConnections..." << std::endl;

    nResult = mpCPToProxy->EnumConnections(ppEnum);
    if (FAILED(nResult))
    {
        std::cout << "..." << this << "@CProxiedConnectionPoint::EnumConnections: "
                  << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }

    *ppEnum = reinterpret_cast<IEnumConnections*>(
        new CProxiedEnumConnections(mpBaseClassUnknown, *ppEnum));

    std::cout << "..." << this << "@CProxiedConnectionPoint::EnumConnections: S_OK" << std::endl;

    return S_OK;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

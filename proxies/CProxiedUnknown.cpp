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
#include <cstdlib>
#include <iostream>
#include <string>

#include <Windows.h>

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedConnectionPointContainer.hpp"
#include "CProxiedDispatch.hpp"
#include "CProxiedEnumConnectionPoints.hpp"
#include "CProxiedEnumConnections.hpp"
#include "CProxiedSink.hpp"
#include "CProxiedUnknown.hpp"

static ThreadProcParam* pGlobalParam;

unsigned CProxiedUnknown::mnIndent = 0;
bool CProxiedUnknown::mbIsAtBeginningOfLine = true;

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
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedUnknown::CTOR(" << pBaseClassUnknown << ", "
                  << pUnknownToProxy << ", " << rIID1 << ", " << rIID2 << ")" << std::endl;
}

CProxiedUnknown::CProxiedUnknown(IUnknown* pUnknownToProxy, const IID& rIID1, const IID& rIID2)
    : CProxiedUnknown(nullptr, pUnknownToProxy, rIID1, rIID2)
{
}

void CProxiedUnknown::setParam(ThreadProcParam* pParam) { pGlobalParam = pParam; }

ThreadProcParam* CProxiedUnknown::getParam() { return pGlobalParam; }

void CProxiedUnknown::increaseIndent() { mnIndent++; }

void CProxiedUnknown::decreaseIndent()
{
    assert(mnIndent > 0);
    mnIndent--;
}

std::string CProxiedUnknown::indent() { return std::string(mnIndent * 4, ' '); }

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
            if (getParam()->mbVerbose)
                std::cout << this << "@CProxiedUnknown::QueryInterface(IID_IUnknown): base: "
                          << mpBaseClassUnknown << std::endl;
            *ppvObject = mpBaseClassUnknown;
        }
        else
        {
            if (getParam()->mbVerbose)
                std::cout << this
                          << "@CProxiedUnknown::QueryInterface(IID_IUnknown): self: " << this
                          << std::endl;
            *ppvObject = this;
        }
        AddRef();
        return S_OK;
    }

    if (!IsEqualIID(maIID1, IID_NULL) && IsEqualIID(riid, maIID1))
    {
        if (getParam()->mbVerbose)
            std::cout << this << "@CProxiedUnknown::QueryInterface(" << riid << "): self: " << this
                      << std::endl;
        *ppvObject = this;
        AddRef();
        return S_OK;
    }

    if (!IsEqualIID(maIID2, IID_NULL) && IsEqualIID(riid, maIID2))
    {
        if (getParam()->mbVerbose)
            std::cout << this << "@CProxiedUnknown::QueryInterface(" << riid << "): self: " << this
                      << std::endl;
        *ppvObject = this;
        AddRef();
        return S_OK;
    }

    auto p = mpExtraInterfaces->maExtraInterfaces.find(riid);
    if (p != mpExtraInterfaces->maExtraInterfaces.end())
    {
        if (getParam()->mbVerbose)
            std::cout << this << "@CProxiedUnknown::QueryInterface(" << riid
                      << "): found: " << p->second << ": S_OK" << std::endl;
        *ppvObject = p->second;
        return S_OK;
    }

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedUnknown::QueryInterface(" << riid << ")..." << std::endl;

    // OK, at this stage we have to query the real object for it

    nResult = mpUnknownToProxy->QueryInterface(riid, ppvObject);

    // Special cases

    if (nResult == S_OK && IsEqualIID(riid, IID_IDispatch))
    {
        *ppvObject = new CProxiedDispatch(mpBaseClassUnknown ? mpBaseClassUnknown : this,
                                          (IDispatch*)*ppvObject);

        mpExtraInterfaces->maExtraInterfaces[riid] = *ppvObject;

        if (getParam()->mbVerbose)
            std::cout << "..." << this << "@CProxiedUnknown::QueryInterface(" << riid << "): S_OK"
                      << std::endl;

        return S_OK;
    }

    if (nResult == S_OK && IsEqualIID(riid, IID_IConnectionPointContainer))
    {
        // We must proxy it so that we can reverse-proxy the calls to advised sinks

        IProvideClassInfo* pPCI = nullptr;
        nResult = mpUnknownToProxy->QueryInterface(IID_IProvideClassInfo, (void**)&pPCI);
        if (nResult != S_OK)
        {
            if (!getParam()->mbTraceOnly)
            {
                if (getParam()->mbVerbose)
                    std::cout
                        << "..." << this << "@CProxiedUnknown::QueryInterface(" << riid
                        << "): IConnectionPointContainer but not IProvideClassInfo: E_NOINTERFACE"
                        << std::endl;
                ((IUnknown*)*ppvObject)->Release();
                return E_NOINTERFACE;
            }
        }
        *ppvObject
            = new CProxiedConnectionPointContainer(mpBaseClassUnknown ? mpBaseClassUnknown : this,
                                                   (IConnectionPointContainer*)*ppvObject, pPCI);

        mpExtraInterfaces->maExtraInterfaces[riid] = *ppvObject;

        if (getParam()->mbVerbose)
            std::cout << "..." << this << "@CProxiedUnknown::QueryInterface(" << riid << "): S_OK"
                      << std::endl;

        return S_OK;
    }

    if (getParam()->mbVerbose)
    {
        std::cout << "..." << this << "@CProxiedUnknown::QueryInterface(" << riid << "): ";
        if (nResult == S_OK)
            std::cout << *ppvObject << ": ";
        std::cout << WindowsErrorStringFromHRESULT(nResult) << std::endl;
    }

    if (nResult == S_OK)
        mpExtraInterfaces->maExtraInterfaces[riid] = *ppvObject;

    return nResult;
}

ULONG STDMETHODCALLTYPE CProxiedUnknown::AddRef()
{
    ULONG nRetval = mpUnknownToProxy->AddRef();
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedUnknown::AddRef: " << nRetval << std::endl;

    return nRetval;
}

ULONG STDMETHODCALLTYPE CProxiedUnknown::Release()
{
    ULONG nRetval = mpUnknownToProxy->Release();
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedUnknown::Release: " << nRetval << std::endl;

    return nRetval;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

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

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedEnumConnections.hpp"
#include "CProxiedSink.hpp"

CProxiedEnumConnections::CProxiedEnumConnections(IUnknown* pBaseClassUnknown,
                                                 IEnumConnections* pECToProxy, const char* sLibName)
    : CProxiedUnknown(pBaseClassUnknown, pECToProxy, IID_IEnumConnections, sLibName)
    , mpECToProxy(pECToProxy)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumConnections::CTOR" << std::endl;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnections::Next(ULONG cConnections, LPCONNECTDATA rgcd,
                                                        ULONG* pcFetched)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumConnections::Next(" << cConnections << ")..."
                  << std::endl;

    nResult = mpECToProxy->Next(cConnections, rgcd, pcFetched);
    if (FAILED(nResult) || nResult == S_FALSE)
    {
        if (getParam()->mbVerbose)
            std::cout << "..." << this << "@CProxiedEnumConnections::Next(" << cConnections << ") ("
                      << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }

    ULONG nFetched = (pcFetched ? *pcFetched : 1);

    for (ULONG i = 0; i < nFetched; ++i)
    {
        IUnknown* pUnk = CProxiedSink::existingSink(rgcd[i].pUnk);
        if (pUnk != NULL)
        {
            // FIXME: We don't return the pointer to our proxy of the sink, but the original sink.
            // That is hopefully not a problem?
            rgcd[i].pUnk = pUnk;
        }
        else
        {
            rgcd[i].pUnk = new CProxiedUnknown(nullptr, rgcd[i].pUnk, IID_NULL, msLibName);
        }
        if (getParam()->mbVerbose)
            std::cout << "..." << this << "@CProxiedEnumConnections::Next(" << cConnections
                      << "): " << i << ": " << rgcd[i].pUnk << std::endl;
    }

    if (getParam()->mbVerbose)
    {
        std::cout << "..." << this << "@CProxiedEnumConnections::Next(" << cConnections << ") ("
                  << __LINE__ << "): ";
        if (pcFetched)
            std::cout << *pcFetched << ": ";
        std::cout << WindowsErrorStringFromHRESULT(nResult) << std::endl;
    }

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnections::Skip(ULONG cConnections)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumConnections::Skip(" << cConnections << ")..."
                  << std::endl;

    nResult = mpECToProxy->Skip(cConnections);

    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedEnumConnections::Skip(" << cConnections
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnections::Reset()
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumConnections::Reset..." << std::endl;

    nResult = mpECToProxy->Reset();

    if (getParam()->mbVerbose)
        std::cout << "..." << this
                  << "@CProxiedEnumConnections::Reset: " << WindowsErrorStringFromHRESULT(nResult)
                  << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnections::Clone(IEnumConnections** /* ppEnum */)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumConnections::Clone: E_NOTIMPL" << std::endl;

    return E_NOTIMPL;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

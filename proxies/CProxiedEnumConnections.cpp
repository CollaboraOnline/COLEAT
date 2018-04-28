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

#include "CProxiedEnumConnections.hpp"
#include "utilstemp.hpp"

CProxiedEnumConnections::CProxiedEnumConnections(IUnknown* pBaseClassUnknown,
                                                 IEnumConnections* pECToProxy)
    : CProxiedUnknown(pBaseClassUnknown, pECToProxy, IID_IEnumConnections)
    , mpECToProxy(pECToProxy)
{
    std::cout << this << "@CProxiedEnumConnections::CTOR" << std::endl;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnections::Next(ULONG cConnections, LPCONNECTDATA rgcd,
                                                        ULONG* pcFetched)
{
    HRESULT nResult;

    std::cout << this << "@CProxiedEnumConnections::Next(" << cConnections << ")..." << std::endl;

    nResult = mpECToProxy->Next(cConnections, rgcd, pcFetched);
    if (FAILED(nResult))
    {
        std::cout << "..." << this << "@CProxiedEnumConnections::Next(" << cConnections
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }
    std::cout << "..." << this << "@CProxiedEnumConnections::Next(" << cConnections
              << "): " << WindowsErrorStringFromHRESULT(nResult);
    if (pcFetched)
        std::cout << *pcFetched;
    std::cout << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnections::Skip(ULONG cConnections)
{
    HRESULT nResult;

    std::cout << this << "@CProxiedEnumConnections::Skip(" << cConnections << ")..." << std::endl;

    nResult = mpECToProxy->Skip(cConnections);

    std::cout << "..." << this << "@CProxiedEnumConnections::Skip(" << cConnections
              << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnections::Reset()
{
    HRESULT nResult;

    std::cout << this << "@CProxiedEnumConnections::Reset..." << std::endl;

    nResult = mpECToProxy->Reset();

    std::cout << "..." << this
              << "@CProxiedEnumConnections::Reset: " << WindowsErrorStringFromHRESULT(nResult)
              << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnections::Clone(IEnumConnections** /* ppEnum */)
{
    std::cout << this << "@CProxiedEnumConnections::Clone: E_NOTIMPL" << std::endl;

    return E_NOTIMPL;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

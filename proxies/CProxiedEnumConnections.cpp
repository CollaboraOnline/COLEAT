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
#include "CProxiedSink.hpp"
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
    if (FAILED(nResult) || nResult == S_FALSE)
    {
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
#if 1 // Temporary information for debugging                                                       \
    // Be careful not to overwrite the outer nResult here as that is what we will return
            HRESULT nHr;
            std::cout << "=== unknown one " << rgcd[i].pUnk
                      << " ref count: " << rgcd[i].pUnk->AddRef() - 1 << std::endl;
            rgcd[i].pUnk->Release();
            IDispatch* pDispatch;
            if (SUCCEEDED(rgcd[i].pUnk->QueryInterface(IID_IDispatch, (void**)&pDispatch)))
            {
                std::cout << "    is IDispatch\n";
                ITypeInfo* pTI;
                nHr = pDispatch->GetTypeInfo(0, LOCALE_USER_DEFAULT, &pTI);
                if (FAILED(nHr))
                    std::cout << "    GetTypeInfo: " << WindowsErrorStringFromHRESULT(nHr)
                              << std::endl;
                else
                {
                    BSTR sName;
                    if (SUCCEEDED(pTI->GetDocumentation(MEMBERID_NIL, &sName, NULL, NULL, NULL)))
                        std::wcout << "    " << sName << "\n";
                }
            }
#endif
            rgcd[i].pUnk = new CProxiedUnknown(rgcd[i].pUnk, IID_NULL);
        }
        std::cout << "..." << this << "@CProxiedEnumConnections::Next(" << cConnections
                  << "): " << i << ": " << rgcd[i].pUnk << std::endl;
    }

    std::cout << "..." << this << "@CProxiedEnumConnections::Next(" << cConnections << ") ("
              << __LINE__ << "): ";
    if (pcFetched)
        std::cout << *pcFetched << ": ";
    std::cout << WindowsErrorStringFromHRESULT(nResult) << std::endl;

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

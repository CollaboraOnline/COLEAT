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

#include "CProxiedEnumVARIANT.hpp"

CProxiedEnumVARIANT::CProxiedEnumVARIANT(IUnknown* pUnknownToProxy, const char* sLibName)
    : CProxiedUnknown(nullptr, pUnknownToProxy, IID_IEnumVARIANT, sLibName)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumVARIANT::CTOR(" << pUnknownToProxy << ")..."
                  << std::endl;

    nResult = pUnknownToProxy->QueryInterface(IID_IEnumVARIANT, (void**)&mpEVToProxy);
    if (FAILED(nResult))
    {
        if (getParam()->mbVerbose)
            std::cout << "..." << this << "@CProxiedEnumVARIANT::CTOR(" << pUnknownToProxy
                      << "): Not an IEnumVARIANT?" << std::endl;
        mpEVToProxy = NULL;
    }
}

// IEnumVARIANT

HRESULT STDMETHODCALLTYPE CProxiedEnumVARIANT::Next(ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched)
{
    HRESULT nResult;

    if (mpEVToProxy == nullptr)
    {
        if (getParam()->mbVerbose)
            std::cout << this << "@CProxiedEnumVARIANT::Next(" << celt << "): E_ NOTIMPL"
                      << std::endl;
        return E_NOTIMPL;
    }

    nResult = mpEVToProxy->Next(celt, rgVar, pCeltFetched);

    if (getParam()->mbTrace)
        std::cout << "Enum<" << this << ">.Next(" << celt << "): " << HRESULT_to_string(nResult)
                  << std::endl;
    else if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumVARIANT::Next(" << celt
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    if (FAILED(nResult) || nResult == S_FALSE)
        return nResult;

    if (getParam()->mbTrace || getParam()->mbVerbose)
    {
        ULONG nFetched;
        if (pCeltFetched != NULL)
            nFetched = *pCeltFetched;
        else
            nFetched = 1;

        for (ULONG i = 0; i < nFetched; i++)
        {
            if (rgVar[i].vt == VT_DISPATCH)
            {
                IUnknown* pExisting = find(rgVar[i].pdispVal);
                if (pExisting != nullptr)
                {
                    std::cout << "... Already have proxy for " << rgVar[i].pdispVal << "\n";
                    rgVar[i].pdispVal = static_cast<IDispatch*>(pExisting);
                }
            }
            std::cout << "... " << i << ": " << rgVar[i] << "\n";
        }
        std::cout << std::flush;
    }

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumVARIANT::Skip(ULONG celt)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumVARIANT::Skip(" << celt << ")..." << std::endl;

    nResult = mpEVToProxy->Skip(celt);

    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedEnumVARIANT::Skip(" << celt
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumVARIANT::Reset()
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumVARIANT::Reset..." << std::endl;

    nResult = mpEVToProxy->Reset();

    if (getParam()->mbVerbose)
        std::cout << "..." << this
                  << "@CProxiedEnumVARIANT::Reset: " << WindowsErrorStringFromHRESULT(nResult)
                  << std::endl;
    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumVARIANT::Clone(IEnumVARIANT** ppEnum)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumVARIANT::Clone..." << std::endl;

    nResult = mpEVToProxy->Clone(ppEnum);

    if (getParam()->mbVerbose)
        std::cout << "..." << this
                  << "@CProxiedEnumVARIANT::Clone: " << WindowsErrorStringFromHRESULT(nResult)
                  << std::endl;

    if (!FAILED(nResult))
    {
        *ppEnum = reinterpret_cast<IEnumVARIANT*>(new CProxiedEnumVARIANT(*ppEnum, msLibName));
    }

    return nResult;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

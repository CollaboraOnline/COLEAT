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

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedClassFactory.hpp"

CProxiedClassFactory::CProxiedClassFactory(IClassFactory* pClassFactoryToProxy,
                                           const char* sLibName)
    : CProxiedUnknown(nullptr, pClassFactoryToProxy, IID_IClassFactory, sLibName)
    , mpClassFactoryToProxy(pClassFactoryToProxy)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedClassFactory::CTOR(" << pClassFactoryToProxy << ")"
                  << std::endl;
}

HRESULT STDMETHODCALLTYPE CProxiedClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid,
                                                               void** ppvObject)
{
    HRESULT nResult = mpClassFactoryToProxy->CreateInstance(pUnkOuter, riid, ppvObject);

    if (getParam()->mbVerbose)
    {
        std::cout << this << "@CProxiedClassFactory::CreateInstance(" << riid << "): ";
        if (nResult == S_OK)
            std::cout << *ppvObject << ": ";
        std::cout << HRESULT_to_string(nResult) << std::endl;
    }

    return nResult;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

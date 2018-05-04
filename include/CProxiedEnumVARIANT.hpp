/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_CProxiedEnumVARIANT_hpp
#define INCLUDED_CProxiedEnumVARIANT_hpp

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <map>

#include <Windows.h>

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedUnknown.hpp"

class CProxiedEnumVARIANT : public CProxiedUnknown
{
private:
    IEnumVARIANT* mpEVToProxy;

public:
    CProxiedEnumVARIANT(IUnknown* pUnknownToProxy, const char* sLibName);

    // IEnumVARIANT
    virtual HRESULT STDMETHODCALLTYPE Next(ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched);

    virtual HRESULT STDMETHODCALLTYPE Skip(ULONG celt);

    virtual HRESULT STDMETHODCALLTYPE Reset();

    virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnum);
};

#endif // INCLUDED_CProxiedEnumVARIANT_hpp

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

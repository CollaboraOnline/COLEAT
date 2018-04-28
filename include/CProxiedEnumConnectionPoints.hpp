/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_CProxiedEnumConnectionPoints_hpp
#define INCLUDED_CProxiedEnumConnectionPoints_hpp

#pragma warning(push)
#pragma warning(disable: 4668 4820 4917)

#include <map>
#include <string>
#include <vector>

#include <Windows.h>

#pragma warning(pop)

#include "CProxiedUnknown.hpp"

// We have all these proxy classes inherit from CProxiedUnknown and not from the corresponding
// actual abstract interface class so that they will automatically inherit CProxiedUnknown's
// implementations of the IUnknown member functions. We must therefore declare the member functions
// in the same order as those of the actual abstract interface class as virtual (and not overridden)
// so that they go into the vtbl in the expected order.

class CProxiedEnumConnectionPoints : public CProxiedUnknown
{
private:
    CProxiedConnectionPointContainer* const mpCPC;
    IEnumConnectionPoints* const mpECPToProxy;

public:
    CProxiedEnumConnectionPoints(IUnknown* pBaseClassUnknown,
                                 CProxiedConnectionPointContainer* pCPC,
                                 IEnumConnectionPoints* pECPToProxy);

    // IEnumConnectionPoints
    virtual HRESULT STDMETHODCALLTYPE Next(ULONG cConnections, LPCONNECTIONPOINT* ppCP,
                                           ULONG* pcFetched);

    virtual HRESULT STDMETHODCALLTYPE Skip(ULONG cConnections);

    virtual HRESULT STDMETHODCALLTYPE Reset();

    virtual HRESULT STDMETHODCALLTYPE Clone(IEnumConnectionPoints** ppEnum);
};

#endif // INCLUDED_CProxiedEnumConnectionPoints_hpp

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

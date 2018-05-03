/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_CProxiedConnectionPointContainer_hpp
#define INCLUDED_CProxiedConnectionPointContainer_hpp

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <map>
#include <string>
#include <vector>

#include <Windows.h>
#include <OCIdl.h>

#pragma warning(pop)

#include "CProxiedUnknown.hpp"

class CProxiedConnectionPointContainer : public CProxiedUnknown
{
private:
    struct ConnectionPointMapHolder
    {
        std::map<IID, IConnectionPoint*> maConnectionPoints;
    };

    IConnectionPointContainer* const mpCPCToProxy;
    ConnectionPointMapHolder* const mpConnectionPoints;

public:
    IProvideClassInfo* const mpProvideClassInfo;

    CProxiedConnectionPointContainer(IUnknown* pBaseClassUnknown,
                                     IConnectionPointContainer* pCPCToProxy,
                                     IProvideClassInfo* pProvideClassInfo, const char* sLibName);

    // IConnectionPointContainer
    virtual HRESULT STDMETHODCALLTYPE EnumConnectionPoints(IEnumConnectionPoints** ppEnum);

    virtual HRESULT STDMETHODCALLTYPE FindConnectionPoint(REFIID riid, IConnectionPoint** ppCP);
};

#endif // INCLUDED_CProxiedConnectionPointContainer_hpp

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

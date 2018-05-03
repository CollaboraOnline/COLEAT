/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_CProxiedConnectionPoint_hpp
#define INCLUDED_CProxiedConnectionPoint_hpp

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <map>
#include <string>
#include <vector>

#include <Windows.h>

#pragma warning(pop)

#include "outgoingmap.hpp"

#include "CProxiedUnknown.hpp"
#include "CProxiedConnectionPointContainer.hpp"

// We have all these proxy classes inherit from CProxiedUnknown and not from the corresponding
// actual abstract interface class so that they will automatically inherit CProxiedUnknown's
// implementations of the IUnknown member functions. We must therefore declare the member functions
// in the same order as those of the actual abstract interface class as virtual (and not overridden)
// so that they go into the vtbl in the expected order.

class CProxiedConnectionPoint : public CProxiedUnknown
{
private:
    struct AdvisedSinkHolder
    {
        std::map<DWORD, IDispatch*> maAdvisedSinks;
    };

    CProxiedConnectionPointContainer* mpContainer;
    IConnectionPoint* mpCPToProxy;
    IID maIID;
    ITypeInfo* mpTypeInfoOfOutgoingInterface;
    const OutgoingInterfaceMapping maMapEntry;
    AdvisedSinkHolder* const mpAdvisedSinks;

public:
    CProxiedConnectionPoint(IUnknown* pBaseClassUnknown,
                            CProxiedConnectionPointContainer* pContainer,
                            IConnectionPoint* pCPToProxy, IID aIID,
                            ITypeInfo* pTypeInfoOfOutgoingInterface,
                            const OutgoingInterfaceMapping& rMapEntry, const char* sLibName);

    // IConnectionPoint
    virtual HRESULT STDMETHODCALLTYPE GetConnectionInterface(IID* pIID);

    virtual HRESULT STDMETHODCALLTYPE
    GetConnectionPointContainer(IConnectionPointContainer** ppCPC);

    virtual HRESULT STDMETHODCALLTYPE Advise(IUnknown* pUnkSink, DWORD* pdwCookie);

    virtual HRESULT STDMETHODCALLTYPE Unadvise(DWORD dwCookie);

    virtual HRESULT STDMETHODCALLTYPE EnumConnections(IEnumConnections** ppEnum);
};

#endif // INCLUDED_CProxiedConnectionPoint_hpp

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

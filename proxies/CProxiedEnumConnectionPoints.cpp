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

#include "CProxiedConnectionPoint.hpp"
#include "CProxiedConnectionPointContainer.hpp"
#include "CProxiedEnumConnectionPoints.hpp"

#include "OutgoingInterfaceMap.hxx"

CProxiedEnumConnectionPoints::CProxiedEnumConnectionPoints(IUnknown* pBaseClassUnknown,
                                                           CProxiedConnectionPointContainer* pCPC,
                                                           IEnumConnectionPoints* pECPToProxy,
                                                           const char* sLibName)
    : CProxiedUnknown(pBaseClassUnknown, pECPToProxy, IID_IEnumConnectionPoints, sLibName)
    , mpCPC(pCPC)
    , mpECPToProxy(pECPToProxy)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumConnectionPoints::CTOR" << std::endl;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnectionPoints::Next(ULONG cConnections,
                                                             LPCONNECTIONPOINT* ppCP,
                                                             ULONG* pcFetched)
{
    if (!ppCP || !pcFetched)
        return E_POINTER;

    if (cConnections == 0)
        return E_INVALIDARG;

    // FIXME: Update below code to work like in CProxiedConnectionPointContainer::FindConnectionPoint?
    // Just return E_NOTIMPL for now.
    ;
#if 1
    return E_NOTIMPL;
#else
    std::vector<IConnectionPoint*> vCP(cConnections);
    HRESULT nResult = mpECPToProxy->Next(cConnections, vCP.data(), pcFetched);

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumConnectionPoints::Next(" << cConnections
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    if (FAILED(nResult))
        return nResult;

    for (ULONG i = 0; i < *pcFetched; ++i)
    {
        IID aIID;
        nResult = vCP[i]->GetConnectionInterface(&aIID);
        if (FAILED(nResult))
            return nResult;

        if (getParam()->mbVerbose)
            std::cout << "  " << i << ": " << aIID << std::endl;

        bool bFound = false;
        for (const auto aMapEntry : aOutgoingInterfaceMap)
        {
            const IID aProxiedOrReplacementIID
                = (CProxiedUnknown::getParam()->mbNoReplacement ? aMapEntry.maProxiedAppIID
                                                                : aMapEntry.maReplacementIID);

            if (IsEqualIID(aIID, aProxiedOrReplacementIID))
            {
                ITypeInfo* pTI;
                nResult = findTypeInfo(mpCPC->mpDispatchOfCPC, aProxiedOrReplacementIID, &pTI);
                if (FAILED(nResult))
                    return nResult;
                ppCP[i] = reinterpret_cast<IConnectionPoint*>(new CProxiedConnectionPoint(
                    nullptr, mpCPC, vCP[i], aMapEntry.maProxiedAppIID, pTI));
                bFound = true;
                break;
            }
        }
        if (!bFound)
        {
            std::cout << "Unhandled outgoing interface returned in connection point: " << aIID
                      << std::endl;
            std::abort();
        }
    }

    return S_OK;
#endif
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnectionPoints::Skip(ULONG cConnections)
{
    HRESULT nResult = mpECPToProxy->Skip(cConnections);

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumConnectionPoints::Skip(" << cConnections
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnectionPoints::Reset()
{
    HRESULT nResult = mpECPToProxy->Reset();

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumConnectionPoints::Reset: "
                  << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedEnumConnectionPoints::Clone(IEnumConnectionPoints** ppEnum)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedEnumConnectionPoints::Clone:...";

    IEnumConnectionPoints* pEnumConnectionPoints;
    HRESULT nResult = mpECPToProxy->Clone(&pEnumConnectionPoints);
    if (FAILED(nResult))
    {
        if (getParam()->mbVerbose)
            std::cout << "..." << this << "@CProxiedEnumConnectionPoints::Clone:..."
                      << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }
    *ppEnum = reinterpret_cast<IEnumConnectionPoints*>(new CProxiedEnumConnectionPoints(
        mpBaseClassUnknown, mpCPC, pEnumConnectionPoints, msLibName));

    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedEnumConnectionPoints::Clone: S_OK" << std::endl;

    return S_OK;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

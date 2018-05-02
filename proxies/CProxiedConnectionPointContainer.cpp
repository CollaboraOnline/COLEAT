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

#include <cassert>
#include <iostream>

#include <Windows.h>
#include <OleCtl.h>

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedConnectionPoint.hpp"
#include "CProxiedConnectionPointContainer.hpp"
#include "CProxiedEnumConnectionPoints.hpp"

#include "OutgoingInterfaceMap.hxx"

CProxiedConnectionPointContainer::CProxiedConnectionPointContainer(
    IUnknown* pBaseClassUnknown, IConnectionPointContainer* pCPCToProxy,
    IProvideClassInfo* pProvideClassInfo)
    : CProxiedUnknown(pBaseClassUnknown, pCPCToProxy, IID_IConnectionPointContainer)
    , mpCPCToProxy(pCPCToProxy)
    , mpProvideClassInfo(pProvideClassInfo)
    , mpConnectionPoints(new ConnectionPointMapHolder())
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedConnectionPointContainer::CTOR(" << pBaseClassUnknown << ", "
                  << pCPCToProxy << ", " << pProvideClassInfo << ")" << std::endl;
}

HRESULT STDMETHODCALLTYPE
CProxiedConnectionPointContainer::EnumConnectionPoints(IEnumConnectionPoints** ppEnum)
{
    HRESULT nResult;

    if (!ppEnum)
        return E_POINTER;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedConnectionPointContainer::EnumConnectionPoints" << std::endl;

    IEnumConnectionPoints* pEnumConnectionPoints;
    nResult = mpCPCToProxy->EnumConnectionPoints(&pEnumConnectionPoints);

    if (FAILED(nResult))
        return nResult;

    *ppEnum = reinterpret_cast<IEnumConnectionPoints*>(
        new CProxiedEnumConnectionPoints(nullptr, this, pEnumConnectionPoints));

    return S_OK;
}

HRESULT STDMETHODCALLTYPE
CProxiedConnectionPointContainer::FindConnectionPoint(REFIID riid, IConnectionPoint** ppCP)
{
    HRESULT nResult;

    if (!ppCP)
        return E_POINTER;

    if (mpConnectionPoints->maConnectionPoints.count(riid))
    {
        if (getParam()->mbVerbose)
            std::cout << this << "@CProxiedConnectionPointContainer::FindConnectionPoint(" << riid
                      << "): S_OK" << std::endl;
        *ppCP = mpConnectionPoints->maConnectionPoints[riid];
        (*ppCP)->AddRef();
        return S_OK;
    }

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedConnectionPointContainer::FindConnectionPoint(" << riid << ")..."
                  << std::endl;

    for (const auto aMapEntry : aOutgoingInterfaceMap)
    {
        if (IsEqualIID(riid, aMapEntry.maSourceInterfaceInProxiedApp))
        {
            const IID aProxiedOrReplacementIID
                = (getParam()->mbTraceOnly ? aMapEntry.maSourceInterfaceInProxiedApp
                                           : aMapEntry.maOutgoingInterfaceInReplacement);

            ITypeInfo* pTI = NULL;

            if (mpProvideClassInfo == NULL)
            {
                assert(getParam()->mbTraceOnly);
            }
            else
            {
                ITypeInfo* pCoclassTI;
                nResult = mpProvideClassInfo->GetClassInfo(&pCoclassTI);
                if (FAILED(nResult))
                {
                    if (getParam()->mbVerbose)
                        std::cout << "..." << this
                                  << "@CProxiedConnectionPointContainer::FindConnectionPoint(" << riid
                                  << "): GetClassInfo failed: "
                                  << WindowsErrorStringFromHRESULT(nResult) << std::endl;
                    return nResult;
                }

                TYPEATTR* pCoclassTA;
                nResult = pCoclassTI->GetTypeAttr(&pCoclassTA);
                if (FAILED(nResult))
                {
                    if (getParam()->mbVerbose)
                        std::cout << "..." << this
                                  << "@CProxiedConnectionPointContainer::FindConnectionPoint(" << riid
                                  << "): GetTypeAttr failed: " << WindowsErrorStringFromHRESULT(nResult)
                                  << std::endl;
                    return nResult;
                }

                for (WORD i = 0; i < pCoclassTA->cImplTypes; ++i)
                {
                    INT nImplTypeFlags;
                    nResult = pCoclassTI->GetImplTypeFlags(i, &nImplTypeFlags);
                    if (FAILED(nResult))
                    {
                        if (getParam()->mbVerbose)
                            std::cout << "..." << this
                                      << "@CProxiedConnectionPointContainer::FindConnectionPoint("
                                      << riid << "): GetImplTypeFlags failed: "
                                      << WindowsErrorStringFromHRESULT(nResult) << std::endl;
                        pCoclassTI->ReleaseTypeAttr(pCoclassTA);
                        return nResult;
                    }

                    if (!((nImplTypeFlags & IMPLTYPEFLAG_FDEFAULT)
                          && (nImplTypeFlags & IMPLTYPEFLAG_FSOURCE)))
                        continue;

                    HREFTYPE nHrefType;
                    nResult = pCoclassTI->GetRefTypeOfImplType(i, &nHrefType);
                    if (FAILED(nResult))
                    {
                        if (getParam()->mbVerbose)
                            std::cout << "..." << this
                                      << "@CProxiedConnectionPointContainer::FindConnectionPoint("
                                      << riid << "): GetRefTypeOfImplType failed: "
                                      << WindowsErrorStringFromHRESULT(nResult) << std::endl;
                        pCoclassTI->ReleaseTypeAttr(pCoclassTA);
                        return nResult;
                    }

                    nResult = pCoclassTI->GetRefTypeInfo(nHrefType, &pTI);
                    if (FAILED(nResult))
                    {
                        if (getParam()->mbVerbose)
                            std::cout << "..." << this
                                      << "@CProxiedConnectionPointContainer::FindConnectionPoint("
                                      << riid << "): GetRefTypeInfo failed: "
                                      << WindowsErrorStringFromHRESULT(nResult) << std::endl;
                        pCoclassTI->ReleaseTypeAttr(pCoclassTA);
                        return nResult;
                    }
                    break;
                }

                pCoclassTI->ReleaseTypeAttr(pCoclassTA);
            }

            IConnectionPoint* pCP;
            nResult = mpCPCToProxy->FindConnectionPoint(aProxiedOrReplacementIID, &pCP);
            if (FAILED(nResult))
            {
                if (getParam()->mbVerbose)
                    std::cout << "..." << this
                              << "@CProxiedConnectionPointContainer::FindConnectionPoint(" << riid
                              << ") (" << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult)
                              << std::endl;
                return nResult;
            }
            *ppCP = reinterpret_cast<IConnectionPoint*>(new CProxiedConnectionPoint(
                nullptr, this, pCP, aMapEntry.maSourceInterfaceInProxiedApp, pTI, aMapEntry));

            if (getParam()->mbVerbose)
                std::cout << "..." << this << "@CProxiedConnectionPointContainer::FindConnectionPoint("
                          << riid << "): S_OK" << std::endl;

            mpConnectionPoints->maConnectionPoints[riid] = *ppCP;
            (*ppCP)->AddRef();

            return S_OK;
        }
    }

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedConnectionPointContainer::FindConnectionPoint(" << riid << ") ("
                  << __LINE__ << "): CONNECT_E_NOCONNECTION" << std::endl;
    return CONNECT_E_NOCONNECTION;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

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
#include <cstdlib>
#include <iostream>
#include <string>

#include <Windows.h>

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedMoniker.hpp"

CProxiedPersist::CProxiedPersist(IUnknown* pBaseClassUnknown, IPersist* pPersistToProxy, const char* sLibName)
    : CProxiedUnknown(pBaseClassUnknown, pPersistToProxy, IID_IPersist, IID_IMoniker, sLibName)
    , mpPersistToProxy(pPersistToProxy)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedPersist::CTOR" << std::endl;
}

HRESULT STDMETHODCALLTYPE CProxiedPersist::GetClassID(CLSID *pClassID)
{
    HRESULT nResult;

    nResult = mpPersistToProxy->GetClassID(pClassID);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedPersist::GetClassId: " << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedPersist::GetClassId: " << *pClassID << std::endl;

    return nResult;
}

CProxiedPersistStream::CProxiedPersistStream(IUnknown* pBaseClassUnknown, IPersistStream* pPersistStreamToProxy, const char* sLibName)
    : CProxiedPersist(pBaseClassUnknown, pPersistStreamToProxy, sLibName)
    , mpPersistStreamToProxy(pPersistStreamToProxy)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedPersistStream::CTOR" << std::endl;
}

HRESULT STDMETHODCALLTYPE CProxiedPersistStream::IsDirty(void)
{
    HRESULT nResult;

    nResult = mpPersistStreamToProxy->IsDirty();

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedPersistStream::IsDirty: " << HRESULT_to_string(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedPersistStream::Load(IStream *pStm)
{
    HRESULT nResult;

    nResult = mpPersistStreamToProxy->Load(pStm);

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedPersistStream::Load(" << pStm << "): " << HRESULT_to_string(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedPersistStream::Save(IStream *pStm,
                                                      BOOL fClearDirty)
{
    HRESULT nResult;

    nResult = mpPersistStreamToProxy->Save(pStm, fClearDirty);

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedPersistStream::Save(" << pStm << "," << (fClearDirty?"YES":"NO") << "): "
                  << HRESULT_to_string(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedPersistStream::GetSizeMax(ULARGE_INTEGER *pcbSize)
{
    HRESULT nResult;

    nResult = mpPersistStreamToProxy->GetSizeMax(pcbSize);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedPersistStream::GetSizeMax: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedPersistStream::GetSizeMax: "
                      << pcbSize->QuadPart << std::endl;

    return nResult;
}

CProxiedMoniker::CProxiedMoniker(IUnknown* pBaseClassUnknown, IMoniker* pMonikerToProxy, const char* sLibName)
    : CProxiedPersistStream(pBaseClassUnknown, pMonikerToProxy, sLibName)
    , mpMonikerToProxy(pMonikerToProxy)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedMoniker::CTOR" << std::endl;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::BindToObject(IBindCtx *pbc,
                                                        IMoniker *pmkToLeft,
                                                        REFIID riidResult,
                                                        void **ppvResult)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->BindToObject(pbc, pmkToLeft, riidResult, ppvResult);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker::BindToObject: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::BindToObject: "
                      << *ppvResult << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::BindToStorage(IBindCtx *pbc,
                                                         IMoniker *pmkToLeft,
                                                         REFIID riid,
                                                         void **ppvObj)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->BindToStorage(pbc, pmkToLeft, riid, ppvObj);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker::BindToStorage: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::BindToStorage: "
                      << *ppvObj << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::Reduce(IBindCtx *pbc,
                                                  DWORD dwReduceHowFar,
                                                  IMoniker **ppmkToLeft,
                                                  IMoniker **ppmkReduced)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->Reduce(pbc, dwReduceHowFar, ppmkToLeft, ppmkReduced);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker:Reduce: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::Reduce: "
                      << *ppmkReduced << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::ComposeWith(IMoniker *pmkRight,
                                                       BOOL fOnlyIfNotGeneric,
                                                       IMoniker **ppmkComposite)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->ComposeWith(pmkRight, fOnlyIfNotGeneric, ppmkComposite);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker:ComposeWith: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::ComposeWith: "
                      << *ppmkComposite << std::endl;
    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::Enum(BOOL fForward,
                                                IEnumMoniker **ppenumMoniker)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->Enum(fForward, ppenumMoniker);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker:Enum: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::Enum: "
                      << *ppenumMoniker << std::endl;
    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::IsEqual(IMoniker *pmkOtherMoniker)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->IsEqual(pmkOtherMoniker);

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedMoniker:IsEqual(" << pmkOtherMoniker << "): "
                  << HRESULT_to_string(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::Hash(DWORD *pdwHash)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->Hash(pdwHash);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker:Hash: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::Hash: "
                      << *pdwHash << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::IsRunning(IBindCtx *pbc,
                                                     IMoniker *pmkToLeft,
                                                     IMoniker *pmkNewlyRunning)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->IsRunning(pbc, pmkToLeft, pmkNewlyRunning);

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedMoniker:IsRunning): "
                  << HRESULT_to_string(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::GetTimeOfLastChange(IBindCtx *pbc,
                                                               IMoniker *pmkToLeft,
                                                               FILETIME *pFileTime)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->GetTimeOfLastChange(pbc, pmkToLeft, pFileTime);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker:GetTimeOfLastChange: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::GetTimeOfLastChange: "
                      << ((uint64_t)pFileTime->dwHighDateTime * 0x100000000 + pFileTime->dwLowDateTime) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::Inverse(IMoniker **ppmk)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->Inverse(ppmk);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker:Inverse: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::Inverse: "
                      << *ppmk << std::endl;
    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::CommonPrefixWith(IMoniker *pmkOther,
                                                            IMoniker **ppmkPrefix)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->CommonPrefixWith(pmkOther, ppmkPrefix);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker:CommonPrefixWith: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::CommonPrefixWith: "
                      << *ppmkPrefix << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::RelativePathTo(IMoniker *pmkOther,
                                                          IMoniker **ppmkRelPath)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->RelativePathTo(pmkOther, ppmkRelPath);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker:RelativePathTo: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::RelativePathTo: "
                      << *ppmkRelPath << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::GetDisplayName(IBindCtx *pbc,
                                                          IMoniker *pmkToLeft,
                                                          LPOLESTR *ppszDisplayName)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->GetDisplayName(pbc, pmkToLeft, ppszDisplayName);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker:GetDisplayName: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::GetDisplayName: "
                      << convertUTF16ToUTF8(*ppszDisplayName) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::ParseDisplayName(IBindCtx *pbc,
                                                            IMoniker *pmkToLeft,
                                                            LPOLESTR pszDisplayName,
                                                            ULONG *pchEaten,
                                                            IMoniker **ppmkOut)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->ParseDisplayName(pbc, pmkToLeft, pszDisplayName, pchEaten, ppmkOut);

    if (getParam()->mbVerbose)
        if (nResult != S_OK)
            std::cout << this << "@CProxiedMoniker:ParseDisplayName: "
                      << HRESULT_to_string(nResult) << std::endl;
        else
            std::cout << this << "@CProxiedMoniker::ParseDisplayName: "
                      << *ppmkOut << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::IsSystemMoniker(DWORD *pdwMksys)
{
    HRESULT nResult;

    nResult = mpMonikerToProxy->IsSystemMoniker(pdwMksys);

    if (getParam()->mbVerbose)
    {
        std::cout << this << "@CProxiedMoniker:GetDisplayName: "
                      << HRESULT_to_string(nResult);
        if (nResult == S_OK)
            std::cout << *pdwMksys;
        std::cout << std::endl;
    }

    return nResult;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

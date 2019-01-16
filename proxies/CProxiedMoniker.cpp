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
}

HRESULT STDMETHODCALLTYPE CProxiedPersist::GetClassID(CLSID *pClassID)
{
    return mpPersistToProxy->GetClassID(pClassID);
}

CProxiedPersistStream::CProxiedPersistStream(IUnknown* pBaseClassUnknown, IPersistStream* pPersistStreamToProxy, const char* sLibName)
    : CProxiedPersist(pBaseClassUnknown, pPersistStreamToProxy, sLibName)
    , mpPersistStreamToProxy(pPersistStreamToProxy)
{
}

HRESULT STDMETHODCALLTYPE CProxiedPersistStream::IsDirty(void)
{
    return mpPersistStreamToProxy->IsDirty();
}

HRESULT STDMETHODCALLTYPE CProxiedPersistStream::Load(IStream *pStm)
{
    return mpPersistStreamToProxy->Load(pStm);
}

HRESULT STDMETHODCALLTYPE CProxiedPersistStream::Save(IStream *pStm,
                                                      BOOL fClearDirty)
{
    return mpPersistStreamToProxy->Save(pStm, fClearDirty);
}

HRESULT STDMETHODCALLTYPE CProxiedPersistStream::GetSizeMax(ULARGE_INTEGER *pcbSize)
{
    return mpPersistStreamToProxy->GetSizeMax(pcbSize);
}

CProxiedMoniker::CProxiedMoniker(IUnknown* pBaseClassUnknown, IMoniker* pMonikerToProxy, const char* sLibName)
    : CProxiedPersistStream(pBaseClassUnknown, pMonikerToProxy, sLibName)
    , mpMonikerToProxy(pMonikerToProxy)
{
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::BindToObject(IBindCtx *pbc,
                                                        IMoniker *pmkToLeft,
                                                        REFIID riidResult,
                                                        void **ppvResult)
{
    return mpMonikerToProxy->BindToObject(pbc, pmkToLeft, riidResult, ppvResult);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::BindToStorage(IBindCtx *pbc,
                                                         IMoniker *pmkToLeft,
                                                         REFIID riid,
                                                         void **ppvObj)
{
    return mpMonikerToProxy->BindToStorage(pbc, pmkToLeft, riid, ppvObj);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::Reduce(IBindCtx *pbc,
                                                  DWORD dwReduceHowFar,
                                                  IMoniker **ppmkToLeft,
                                                  IMoniker **ppmkReduced)
{
    return mpMonikerToProxy->Reduce(pbc, dwReduceHowFar, ppmkToLeft, ppmkReduced);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::ComposeWith(IMoniker *pmkRight,
                                                       BOOL fOnlyIfNotGeneric,
                                                       IMoniker **ppmkComposite)
{
    return mpMonikerToProxy->ComposeWith(pmkRight, fOnlyIfNotGeneric, ppmkComposite);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::Enum(BOOL fForward,
                                                IEnumMoniker **ppenumMoniker)
{
    return mpMonikerToProxy->Enum(fForward, ppenumMoniker);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::IsEqual(IMoniker *pmkOtherMoniker)
{
    return mpMonikerToProxy->IsEqual(pmkOtherMoniker);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::Hash(DWORD *pdwHash)
{
    return mpMonikerToProxy->Hash(pdwHash);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::IsRunning(IBindCtx *pbc,
                                                     IMoniker *pmkToLeft,
                                                     IMoniker *pmkNewlyRunning)
{
    return mpMonikerToProxy->IsRunning(pbc, pmkToLeft, pmkNewlyRunning);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::GetTimeOfLastChange(IBindCtx *pbc,
                                                               IMoniker *pmkToLeft,
                                                               FILETIME *pFileTime)
{
    return mpMonikerToProxy->GetTimeOfLastChange(pbc, pmkToLeft, pFileTime);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::Inverse(IMoniker **ppmk)
{
    return mpMonikerToProxy->Inverse(ppmk);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::CommonPrefixWith(IMoniker *pmkOther,
                                                            IMoniker **ppmkPrefix)
{
    return mpMonikerToProxy->CommonPrefixWith(pmkOther, ppmkPrefix);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::RelativePathTo(IMoniker *pmkOther,
                                                          IMoniker **ppmkRelPath)
{
    return mpMonikerToProxy->RelativePathTo(pmkOther, ppmkRelPath);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::GetDisplayName(IBindCtx *pbc,
                                                          IMoniker *pmkToLeft,
                                                          LPOLESTR *ppszDisplayName)
{
    return mpMonikerToProxy->GetDisplayName(pbc, pmkToLeft, ppszDisplayName);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::ParseDisplayName(IBindCtx *pbc,
                                                            IMoniker *pmkToLeft,
                                                            LPOLESTR pszDisplayName,
                                                            ULONG *pchEaten,
                                                            IMoniker **ppmkOut)
{
    return mpMonikerToProxy->ParseDisplayName(pbc, pmkToLeft, pszDisplayName, pchEaten, ppmkOut);
}

HRESULT STDMETHODCALLTYPE CProxiedMoniker::IsSystemMoniker(DWORD *pdwMksys)
{
    return mpMonikerToProxy->IsSystemMoniker(pdwMksys);
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

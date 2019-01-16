/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_CProxiedMoniker_hpp
#define INCLUDED_CProxiedMoniker_hpp

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <string>
#include <vector>

#include <Windows.h>
#include <ObjIdl.h>

#pragma warning(pop)

#include "CProxiedUnknown.hpp"

class CProxiedPersist : public CProxiedUnknown
{
    IPersist* mpPersistToProxy;

public:
    CProxiedPersist(IUnknown* pBaseClassUnknown, IPersist* pPersistToProxy, const char* sLibName);

    // IPersist
    virtual HRESULT STDMETHODCALLTYPE GetClassID(CLSID *pClassID);
};

class CProxiedPersistStream : public CProxiedPersist
{
    IPersistStream* mpPersistStreamToProxy;

public:
    CProxiedPersistStream(IUnknown* pBaseClassUnknown, IPersistStream* pPersistStreamToProxy, const char* sLibName);

    // IPersistStream
    virtual HRESULT STDMETHODCALLTYPE IsDirty(void);

    virtual HRESULT STDMETHODCALLTYPE Load(IStream *pStm);

    virtual HRESULT STDMETHODCALLTYPE Save(IStream *pStm,
                                           BOOL fClearDirty);

    virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER *pcbSize);
};

class CProxiedMoniker : public CProxiedPersistStream
{
private:
    IMoniker* const mpMonikerToProxy;

public:
    CProxiedMoniker(IUnknown* pBaseClassUnknown, IMoniker* pMonikerToProxy, const char* sLibName);

    // IMoniker
    virtual HRESULT STDMETHODCALLTYPE BindToObject(IBindCtx *pbc,
                                                   IMoniker *pmkToLeft,
                                                   REFIID riidResult,
                                                   void **ppvResult);

    virtual HRESULT STDMETHODCALLTYPE BindToStorage(IBindCtx *pbc,
                                                    IMoniker *pmkToLeft,
                                                    REFIID riid,
                                                    void **ppvObj);

    virtual HRESULT STDMETHODCALLTYPE Reduce(IBindCtx *pbc,
                                             DWORD dwReduceHowFar,
                                             IMoniker **ppmkToLeft,
                                             IMoniker **ppmkReduced);

    virtual HRESULT STDMETHODCALLTYPE ComposeWith(IMoniker *pmkRight,
                                                  BOOL fOnlyIfNotGeneric,
                                                  IMoniker **ppmkComposite);

    virtual HRESULT STDMETHODCALLTYPE Enum(BOOL fForward,
                                           IEnumMoniker **ppenumMoniker);

    virtual HRESULT STDMETHODCALLTYPE IsEqual(IMoniker *pmkOtherMoniker);
        
    virtual HRESULT STDMETHODCALLTYPE Hash(DWORD *pdwHash);

    virtual HRESULT STDMETHODCALLTYPE IsRunning(IBindCtx *pbc,
                                                IMoniker *pmkToLeft,
                                                IMoniker *pmkNewlyRunning);

    virtual HRESULT STDMETHODCALLTYPE GetTimeOfLastChange(IBindCtx *pbc,
                                                          IMoniker *pmkToLeft,
                                                          FILETIME *pFileTime);

    virtual HRESULT STDMETHODCALLTYPE Inverse(IMoniker **ppmk);

    virtual HRESULT STDMETHODCALLTYPE CommonPrefixWith(IMoniker *pmkOther,
                                                       IMoniker **ppmkPrefix);

    virtual HRESULT STDMETHODCALLTYPE RelativePathTo(IMoniker *pmkOther,
                                                     IMoniker **ppmkRelPath);

    virtual HRESULT STDMETHODCALLTYPE GetDisplayName(IBindCtx *pbc,
                                                     IMoniker *pmkToLeft,
                                                     LPOLESTR *ppszDisplayName);

    virtual HRESULT STDMETHODCALLTYPE ParseDisplayName(IBindCtx *pbc,
                                                       IMoniker *pmkToLeft,
                                                       LPOLESTR pszDisplayName,
                                                       ULONG *pchEaten,
                                                       IMoniker **ppmkOut);

    virtual HRESULT STDMETHODCALLTYPE IsSystemMoniker(DWORD *pdwMksys);
};

#endif // INCLUDED_CProxiedMoniker_hpp

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_CPROXIEDCOCLASS_HPP
#define INCLUDED_CPROXIEDCOCLASS_HPP

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <Windows.h>

#include "CProxiedDispatch.hpp"

#pragma warning(pop)

#include "interfacemap.hpp"

class CProxiedCoclass : public CProxiedDispatch
{
private:
    const IID maProxiedAppCoclassIID;
    const IID maProxiedAppDefaultIID;
    const IID maReplacementAppCoclassIID;
    IDispatch* mpReplacementAppDispatch;

    IDispatch* createDispatchToProxy(const InterfaceMapping& rMapping);

public:
    CProxiedCoclass(const InterfaceMapping& rMapping);

    static bool IsActive();

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;
};

#endif // INCLUDED_CPROXIEDCOCLASS_HPP

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

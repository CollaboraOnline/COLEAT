/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_CProxiedUnknown_hpp
#define INCLUDED_CProxiedUnknown_hpp

#pragma warning(push)
#pragma warning(disable : 4625 4668 4774 4820 4917)

#include <map>

#include <Windows.h>
#include <OCIdl.h>

#pragma warning(pop)

#include "exewrapper.hpp"
#include "utils.hpp"

class CProxiedUnknown : public IUnknown
{
private:
    IUnknown* const mpUnknownToProxy;

    // We want to have at most one unique CProxiedUnknown object for each COM object.

    struct UnknownMapHolder
    {
        std::map<IUnknown*, void*> maMap;
    };

    static UnknownMapHolder* const mpLookupMap;

    // Also, the proxied object might implement more interfaces than the one we know about (from the
    // type information at proxy generation time). Ones for which the proxied object has responded
    // positively to a QueryInterface(). We need a map from such discovered "extra" interfaces to
    // our proxy objects for them.
    //
    // Or wait. As we don't know anything about such interfaces, how can we proxy them, as we can't
    // construct an object with the necessary vtable entries? So this is a map to what the actual
    // object's QueryInterface returned then. But then QueryInterface calls on it for IID_IUnknown
    // will return the proxied object's IUnknown address, not that of our proxy IUnknow. Which
    // breaks the Object Identity rule on https://msdn.microsoft.com/en-us/library/ms810016.aspx .
    // Sigh...
    //
    // This is not just theoretical pondering. A real-world example that one comes across pretty
    // immediately when trying this stuff on real-world use cases is IConnectionPointContainer. The
    // type library for typical COM and Automation compatible apps don't mention anything about
    // that. (IConnectionPointContainer isn't visible in the (pseudo) IDL generated by the OLE/COM
    // Object Viewer, for instance.) So we must special-case at least IConnectionPointContainer.

    // Must have this in a separate object pointed to from CProxiedUnknown to avoid "C4265
    // 'CProxiedUnknown': class has virtual functions, but destructor is not virtual".

    struct IIDMapHolder
    {
        std::map<IID, void*> maMap;
    };

    IIDMapHolder* const mpExtraInterfaces;

    // For indenting trace output nicely
    static unsigned mnIndent;

protected:
    CProxiedUnknown(IUnknown* pBaseClassUnknown, IUnknown* pUnknownToProxy, const IID& rIID,
                    const char* sLibName);
    CProxiedUnknown(IUnknown* pBaseClassUnknown, IUnknown* pUnknownToProxy, const IID& rIID1,
                    const IID& rIID2, const char* sLibName);

    static CProxiedUnknown* find(IUnknown* pUnknownToProxy);

    IUnknown* const mpBaseClassUnknown;
    const IID maIID1;
    const IID maIID2;
    const char* const msLibName;

public:
    static CProxiedUnknown* get(IUnknown* pBaseClassUnknown, IUnknown* pUnknownToProxy,
                                const IID& rIID, const char* sLibName);
    static CProxiedUnknown* get(IUnknown* pBaseClassUnknown, IUnknown* pUnknownToProxy,
                                const IID& rIID1, const IID& rIID2, const char* sLibName);

    static void setParam(ThreadProcParam* pParam);
    static ThreadProcParam* getParam();

    // Used to keep track of indentation and line breaks in the output in tracing-only (and not
    // verbose) mode. Not sure whether I should just put these in utils.hpp instead as they don't
    // really have anything to do with CProxiedUnknown except that it is a handy location as most of
    // the code is derived from this class.
    static void increaseIndent();
    static void decreaseIndent();
    static std::string indent();

    static bool mbIsAtBeginningOfLine;

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;

    ULONG STDMETHODCALLTYPE AddRef() override;

    ULONG STDMETHODCALLTYPE Release() override;
};

#endif // INCLUDED_CProxiedUnknown_hpp

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

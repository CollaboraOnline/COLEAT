/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_OUTGOINGMAP_HPP
#define INCLUDED_OUTGOINGMAP_HPP

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <Windows.h>

#pragma warning(pop)

struct NameToMemberIdMapping
{
    NameToMemberIdMapping() = delete;

    const char* const mpName;
    MEMBERID mnMemberId;
};

struct OutgoingInterfaceMapping
{
    OutgoingInterfaceMapping() = delete;

    // An outgoing or "source" interface in the original application we generate proxies for
    IID maSourceInterfaceInProxiedApp;

    // The corresponding outgoing interface in the replacement application. The very point is that
    // it needs not be 1:1 to the interface in the original application, but the replacement needs
    // to implement only such callbacks that relevant client applications actually use.
    IID maOutgoingInterfaceInReplacement;

    // Array terminated by entry with null mpName
    const NameToMemberIdMapping* maNameToId;
};
    
#endif // INCLUDED_OUTGOINGMAP_HPP

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

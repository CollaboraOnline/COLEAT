/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_INTERFACEMAP_HPP
#define INCLUDED_INTERFACEMAP_HPP

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <Windows.h>

#pragma warning(pop)

struct InterfaceMapping
{
    IID maFromCoclass, maFromDefault, maReplacementCoclass;
    const char* msFromLibName;
    const char* msFromCoclassName;
};

#endif // INCLUDED_INTERFACEMAP_HPP

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

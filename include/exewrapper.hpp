/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_EXEWRAPPER_HPP
#define INCLUDED_EXEWRAPPER_HPP

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <Windows.h>

#pragma warning(pop)

union FunPtr
{
    PVOID pVoid;
    PROC pProc;
    PTHREAD_START_ROUTINE pStartRoutine;
    HMODULE (WINAPI *pLoadLibraryW)(wchar_t*);
    DWORD (WINAPI *pGetLastError)();
    PVOID (WINAPI *pGetProcAddress)(HMODULE, char*);
};

struct InterfaceMapping
{
    IID maFromCoclass, maFromDefault, maTo;
};
    
struct ThreadProcParam
{
    size_t mnSize;

    FunPtr mpLoadLibraryW;
    FunPtr mpGetLastError;
    FunPtr mpGetProcAddress;

    bool mbPassedSizeCheck;

    char msInjectedDllMainFunction[100];
    wchar_t msFileName[1000];
    wchar_t msErrorExplanation[1000];
    DWORD mnLastError;
    static const int NIIDMAP = 10;
    InterfaceMapping mvInterfaceMap[NIIDMAP];
    int mnMappings;
};

#endif // INCLUDED_EXEWRAPPER_HPP

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

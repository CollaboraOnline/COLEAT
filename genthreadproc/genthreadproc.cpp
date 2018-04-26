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

#define _CRT_SECURE_NO_WARNINGS
#include <cstdio>
#undef _CRT_SECURE_NO_WARNINGS
#include <cstdlib>
#include <cwchar>
#include <iomanip>
#include <iostream>

#include <Windows.h>

#pragma warning(pop)

#include "exewrapper.hpp"

extern "C" {

// This small function is injected into the executable. It is passed a parameter block. It must be
// carefully written to not refer directly to any external data or functions at all. Not even
// integer literals can be used. The mechanism to access such, import tables etc, obviously would
// not work when the just function code is bluntly written into a completely different process. But
// that's OK, we can pass pointers to functions from kernel32.dll at least in the parameter block.

#pragma runtime_checks("", off)
#pragma optimize("", off)

static DWORD WINAPI threadProc(ThreadProcParam* pParam)
{
    HMODULE hDll = pParam->mpLoadLibraryW.pLoadLibraryW(pParam->msFileName);
    if (hDll == NULL)
    {
        pParam->mnLastError = pParam->mpGetLastError.pGetLastError();
        return FALSE;
    }

    union {
        PVOID pVoid;
        PTHREAD_START_ROUTINE pStartRoutine;
    } pMainFunction;

    pMainFunction.pVoid
        = pParam->mpGetProcAddress.pGetProcAddress(hDll, pParam->msInjectedDllMainFunction);
    if (pMainFunction.pVoid == NULL)
    {
        pParam->mnLastError = pParam->mpGetLastError.pGetLastError();
        return FALSE;
    }

#ifndef _WIN64

#define p1(x) x x
#define p2(x) p1(p1(x))
#define p3(x) p2(p2(x))

    // I tried to use inline __asm to place a public label endOfFun here, but
    // doesn't seem possible.

    // Instead place a large bunch of nops here, and search for them, and then search for the
    // "ret" instruction. As we turn off optimization for this function, no funky reordering or
    // other crack should be going on.

    p3(__asm { nop });

#else

    // For 64-bit, we just need to look for the ret instruction and hope there is just one of them?
    ;

#endif

    return (*pMainFunction.pStartRoutine)(pParam);
}

#pragma optimize("", on)
#pragma runtime_checks("", restore)
}

int main(int /* argc */, char** /* argv */)
{
    const unsigned char* const pThreadProc = (unsigned char*)&threadProc;
    const unsigned char* pRover = pThreadProc;

#ifndef _WIN64
    while (std::memcmp(pRover, p3("\x90"), 16) != 0)
        pRover++;
    pRover += 16;
    const unsigned char* const pEndOfNops = pRover;
    // Search for "ret 4" == 0xc2 0x04 0x00
    while (std::memcmp(pRover, "\xc2\x04\x00", 3) != 0 && (pRover - pEndOfNops < 100))
        pRover++;
    if (pRover - pEndOfNops >= 100)
    {
        std::cerr << "Can't find end of threadProc\n";
        std::exit(1);
    }
    pRover += 4;

#else
    // Search for "pop rbp; ret" == 0x5d 0xc3
    while (pRover[0] != 0x5d && pRover[1] != 0xc3 && (pRover < pThreadProc + 1000))
        pRover++;
    if (pRover - pThreadProc >= 1000)
    {
        std::cerr << "Can't find end of threadProc\n";
        std::exit(1);
    }
    pRover += 2;
#endif

    const SIZE_T nSizeOfThreadProc = (SIZE_T)(pRover - pThreadProc);

    std::cout << "static const unsigned char threadProc_";
#ifndef _WIN64
    std::cout << "x86";
#else
    std::cout << "x64";
#endif
    std::cout << "[] = {\n";

    for (SIZE_T i = 0; i < nSizeOfThreadProc; i++)
    {
        if ((i % 16) == 0)
            std::cout << "    ";
        std::cout << "0x" << std::setfill('0') << std::setw(2) << std::hex
                  << (unsigned)pThreadProc[i];
        if (i < nSizeOfThreadProc - 1)
            std::cout << ",";
        if (((i + 1) % 16) == 0)
            std::cout << "\n";
    }
    if ((nSizeOfThreadProc % 16) != 0)
        std::cout << "\n";
    std::cout << "};\n";

    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

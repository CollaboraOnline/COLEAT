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
#include <cwchar>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <map>
#include <sstream>
#include <string>

#define PSAPI_VERSION 1
#include <Windows.h>
#include <Psapi.h>

#pragma warning(pop)

#include "exewrapper.hpp"
#include "utils.hpp"

static void Usage(wchar_t** argv)
{
    std::cerr << convertUTF16ToUTF8(programName(argv[0]))
              << ": Usage error. This is not a program that users should run directly.\n";
    std::exit(1);
}

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

    // For 64-bit, we can only look for the ret instruction and hope there is just one of them?
    ;

#endif

    return (*pMainFunction.pStartRoutine)(pParam);
}

#pragma optimize("", on)
#pragma runtime_checks("", restore)
}

int wmain(int argc, wchar_t** argv)
{
    tryToEnsureStdHandlesOpen();

    if (argc < 3)
        Usage(argv);

    HANDLE hWrappedProcess = (HANDLE)std::stoull(argv[1]);
    HANDLE hWrappedThread = (HANDLE)std::stoull(argv[2]);
    DWORD nHandleFlags;
    if (!GetHandleInformation(hWrappedProcess, &nHandleFlags)
        || !GetHandleInformation(hWrappedThread, &nHandleFlags))
    {
        std::cerr << "Inherited handle to wrapped process or its start thread is invalid?\n";
        std::exit(1);
    }

    int argi = 3;

    bool bDebug = false;
    bool bTraceOnly = false;
    bool bVerbose = false;

    while (argi < argc && argv[argi][0] == L'-')
    {
        switch (argv[argi][1])
        {
            case L'd':
            {
                bDebug = true;
                break;
            }
            case L't':
                bTraceOnly = true;
                break;
            case L'v':
                bVerbose = true;
                break;
            default:
                Usage(argv);
        }
        argi++;
    }

    if (argi != argc)
        Usage(argv);

    // Get our exe pathname and the wrapped process exe pathname.

    const DWORD NFILENAME = 1000;
    wchar_t sMyFileName[NFILENAME];

    // NFILENAME-20 let the longer name of our DLL fit.
    DWORD nSize = GetModuleFileNameW(NULL, sMyFileName, NFILENAME - 20);
    if (nSize == NFILENAME - 20)
    {
        std::cerr << "Pathname of this exe ridiculously long\n";
        std::exit(1);
    }

    wchar_t sWrappedFileName[NFILENAME];
    if (!GetModuleFileNameExW(hWrappedProcess, NULL, sWrappedFileName, NFILENAME))
    {
        std::cerr << "GetModuleFileNameExW failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    if (bDebug)
    {
        // Give the developer a chance to attach us in a debugger
        std::cout << "Waiting for you to attach a debugger to the '"
                  << convertUTF16ToUTF8(baseName(sMyFileName)) << "' process "
                  << std::to_string(GetProcessId(GetCurrentProcess())) << "\n";
        std::cout << "When done with that, set the bWait in the '"
                  << convertUTF16ToUTF8(baseName(sMyFileName)) << "' process to false" << std::endl;
        volatile bool bWait = true;
        while (bWait)
            Sleep(100);
    }

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
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
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
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }
    pRover += 2;
#endif

    const SIZE_T nSizeOfThreadProc = (SIZE_T)(pRover - pThreadProc);

    // Construct name of DLL to inject.
    wchar_t sDllFileName[NFILENAME];
    wcscpy_s(sDllFileName, NFILENAME, sMyFileName);
    wchar_t* pLastDot = wcsrchr(sDllFileName, L'.');
    if (pLastDot == NULL)
    {
        std::cerr << "No period in '" << convertUTF16ToUTF8(sDllFileName) << "'?\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }
    wcscpy_s(pLastDot, NFILENAME - wcslen(sDllFileName), L"-injected.dll");

    // Inject our magic DLL, and start its main function.

    // Can we really trust that kernel32.dll is at the same address in all processes (exempt from
    // ASLR) also in the future, even if it apparently now is?
    HMODULE hKernel32 = GetModuleHandleW(L"kernel32.dll");

    ThreadProcParam aParam;

    std::memset(&aParam, 0, sizeof(ThreadProcParam));
    aParam.mnSize = sizeof(ThreadProcParam);
    aParam.mpLoadLibraryW.pVoid = GetProcAddress(hKernel32, "LoadLibraryW");
    aParam.mpGetLastError.pVoid = GetProcAddress(hKernel32, "GetLastError");
    aParam.mpGetProcAddress.pVoid = GetProcAddress(hKernel32, "GetProcAddress");
    aParam.mbTraceOnly = bTraceOnly;
    aParam.mbVerbose = bVerbose;
    strcpy_s(aParam.msInjectedDllMainFunction, ThreadProcParam::NFUNCTION,
             "InjectedDllMainFunction");
    wcscpy_s(aParam.msFileName, ThreadProcParam::NFILENAME, sDllFileName);

    void* pParamRemote
        = VirtualAllocEx(hWrappedProcess, NULL, sizeof(aParam), MEM_COMMIT, PAGE_READWRITE);
    if (pParamRemote == NULL)
    {
        std::cerr << "VirtualAllocEx failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    SIZE_T nBytesWritten;
    if (!WriteProcessMemory(hWrappedProcess, pParamRemote, &aParam, sizeof(aParam), &nBytesWritten))
    {
        std::cerr << "WriteProcessMemory failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    void* pThreadProcRemote
        = VirtualAllocEx(hWrappedProcess, NULL, nSizeOfThreadProc, MEM_COMMIT, PAGE_READWRITE);
    if (pThreadProcRemote == NULL)
    {
        std::cerr << "VirtualAllocEx failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    if (!WriteProcessMemory(hWrappedProcess, pThreadProcRemote, pThreadProc, nSizeOfThreadProc,
                            &nBytesWritten))
    {
        std::cerr << "WriteProcessMemory failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    DWORD nOldProtection;
    if (!VirtualProtectEx(hWrappedProcess, pThreadProcRemote, nSizeOfThreadProc, PAGE_EXECUTE,
                          &nOldProtection))
    {
        std::cerr << "VirtualProtectEx failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    if (bDebug)
    {
        // Give the developer a chance to debug the wrapped program, too.
        std::cout << "Waiting for you to attach a debugger to the wrapped '"
                  << convertUTF16ToUTF8(baseName(sWrappedFileName)) << "' process "
                  << std::to_string(GetProcessId(hWrappedProcess)) << "\n";
        std::cout << "When done with that, set the bWait in the '"
                  << convertUTF16ToUTF8(baseName(sMyFileName)) << "' process to false" << std::endl;
        volatile bool bWait = true;
        while (bWait)
            Sleep(100);
    }

    FunPtr pProc;
    pProc.pVoid = pThreadProcRemote;
    HANDLE hThread
        = CreateRemoteThread(hWrappedProcess, NULL, 0, pProc.pStartRoutine, pParamRemote, 0, NULL);
    if (!hThread)
    {
        std::cerr << "CreateRemoteThread failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    WaitForSingleObject(hThread, INFINITE);

    SIZE_T nBytesRead;
    if (!ReadProcessMemory(hWrappedProcess, pParamRemote, &aParam, sizeof(aParam), &nBytesRead))
    {
        std::cerr << "ReadProcessMemory failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    DWORD nExitCode;
    if (!GetExitCodeThread(hThread, &nExitCode))
    {
        std::cerr << "GetExitCodeThread failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }
    if (!nExitCode)
    {
        std::cerr << "Injected thread failed: ";

        if (!aParam.mbPassedSizeCheck)
        {
            if (aParam.mnLastError != 0)
                std::cerr << WindowsErrorString(aParam.mnLastError) << "\n";
            else
                std::cerr << "Mismatched parameter structure sizes, COLEAT build or installation "
                             "problem.\n";
        }
        else
        {
            if (aParam.msErrorExplanation[0])
                std::cerr << convertUTF16ToUTF8(aParam.msErrorExplanation) << "\n";
            else
                std::cerr << WindowsErrorString(aParam.mnLastError) << "\n";
        }

        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    CloseHandle(hThread);

    DWORD nPreviousSuspendCount = ResumeThread(hWrappedThread);
    if (nPreviousSuspendCount == (DWORD)-1)
    {
        std::cerr << "ResumeThread failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }
    else if (nPreviousSuspendCount == 0)
    {
        std::cerr << "Huh, thread was not suspended?\n";
    }
    else if (nPreviousSuspendCount > 1)
    {
        std::cerr << "Thread still suspended after ResumeThread\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    // Wait for the process to finish.
    WaitForSingleObject(hWrappedProcess, INFINITE);

    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

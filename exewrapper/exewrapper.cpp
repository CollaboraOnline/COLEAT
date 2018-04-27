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
#include <cassert>
#include <codecvt>
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

static bool bDebug = false;

inline bool operator<(const IID& a, const IID& b) { return std::memcmp(&a, &b, sizeof(a)) < 0; }

static void Usage(wchar_t** argv)
{
    std::wcerr << programName(argv[0])
               << L": Usage error. This is not a program that users should run directly.\n";
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

static bool parseMapping(wchar_t* pLine, InterfaceMapping& rMapping)
{
    wchar_t* pColon1 = std::wcschr(pLine, L':');
    if (!pColon1)
        return false;
    *pColon1 = L'\0';

    wchar_t* pColon2 = std::wcschr(pColon1 + 1, L':');
    if (!pColon2)
        return false;
    *pColon2 = L'\0';

    if (FAILED(IIDFromString(pLine, &rMapping.maFromCoclass))
        || FAILED(IIDFromString(pColon1 + 1, &rMapping.maFromDefault))
        || FAILED(IIDFromString(pColon2 + 1, &rMapping.maTo)))
        return false;

    return true;
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
        std::wcerr << L"Inherited handle to wrapped process or its start thread is invalid?\n";
        std::exit(1);
    }

    int argi = 3;
    std::map<IID, InterfaceMapping> aInterfaceMap;

    while (argi < argc && argv[argi][0] == L'-')
    {
        switch (argv[argi][1])
        {
            case L'd':
            {
                bDebug = true;
                break;
            }
            case L'm':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                InterfaceMapping aMapping;
                if (!parseMapping(argv[argi + 1], aMapping))
                {
                    std::wcerr << L"Invalid IIDs\n";
                    Usage(argv);
                }
                aInterfaceMap[aMapping.maFromCoclass] = aMapping;
                argi++;
                break;
            }
            case L'M':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                std::wifstream aMFile(argv[argi + 1]);
                if (!aMFile.good())
                {
                    std::wcerr << L"Could not open " << argv[argi + 1] << L" for reading\n";
                    std::exit(1);
                }
                while (!aMFile.eof())
                {
                    std::wstring sLine;
                    std::getline(aMFile, sLine);
                    if (sLine.length() == 0 || sLine[0] == L'#')
                        continue;
                    wchar_t* pLine = _wcsdup(sLine.data());
                    InterfaceMapping aMapping;
                    if (!parseMapping(pLine, aMapping))
                    {
                        std::wcerr << L"Invalid IIDs\n";
                        Usage(argv);
                    }
                    aInterfaceMap[aMapping.maFromCoclass] = aMapping;
                    std::free(pLine);
                }
                aMFile.close();
                argi++;
                break;
            }
            default:
                Usage(argv);
        }
        argi++;
    }

    if (argi != argc)
        Usage(argv);

    if (aInterfaceMap.size() > ThreadProcParam::NIIDMAP)
    {
        std::wcerr << L"At most " << ThreadProcParam::NIIDMAP << " IID mappings possible\n";
        std::exit(1);
    }

    // Get our exe pathname and the wrapped process exe pathname.

    const DWORD NFILENAME = 1000;
    wchar_t sMyFileName[NFILENAME];

    // NFILENAME-20 let the longer name of our DLL fit.
    DWORD nSize = GetModuleFileNameW(NULL, sMyFileName, NFILENAME - 20);
    if (nSize == NFILENAME - 20)
    {
        std::wcerr << L"Pathname of this exe ridiculously long\n";
        std::exit(1);
    }

    wchar_t sWrappedFileName[NFILENAME];
    if (!GetModuleFileNameExW(hWrappedProcess, NULL, sWrappedFileName, NFILENAME))
    {
        std::wcerr << L"GetModuleFileNameExW failed: " << WindowsErrorString(GetLastError())
                   << L"\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    if (bDebug)
    {
        // Give the developer a chance to attach us in a debugger
        std::wcout << L"Waiting for you to attach a debugger to the " << baseName(sMyFileName)
                   << L" process " << std::to_wstring(GetProcessId(GetCurrentProcess())) << L"\n";
        std::wcout << L"When done with that, set the bWait in the " << baseName(sMyFileName)
                   << L" process to false" << std::endl;
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
        std::wcerr << L"Can't find end of threadProc\n";
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
        std::wcerr << L"Can't find end of threadProc\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }
    pRover += 2;
#endif

    const SIZE_T nSizeOfThreadProc = (SIZE_T)(pRover - pThreadProc);

    // Construct name of DLL to inject.
    wchar_t sDllFileName[NFILENAME];
    wcscpy(sDllFileName, sMyFileName);
    wchar_t* pLastDot = wcsrchr(sDllFileName, L'.');
    if (pLastDot == NULL)
    {
        std::wcerr << L"No period in " << sDllFileName << L"?\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }
    wcscpy(pLastDot, L"-injected.dll");

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
    std::strcpy(aParam.msInjectedDllMainFunction, "InjectedDllMainFunction");
    wcscpy(aParam.msFileName, sDllFileName);

    for (auto i : aInterfaceMap)
    {
        aParam.mvInterfaceMap[aParam.mnMappings] = i.second;
        ++aParam.mnMappings;
    }

    void* pParamRemote
        = VirtualAllocEx(hWrappedProcess, NULL, sizeof(aParam), MEM_COMMIT, PAGE_READWRITE);
    if (pParamRemote == NULL)
    {
        std::wcerr << L"VirtualAllocEx failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    SIZE_T nBytesWritten;
    if (!WriteProcessMemory(hWrappedProcess, pParamRemote, &aParam, sizeof(aParam), &nBytesWritten))
    {
        std::wcerr << L"WriteProcessMemory failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    void* pThreadProcRemote
        = VirtualAllocEx(hWrappedProcess, NULL, nSizeOfThreadProc, MEM_COMMIT, PAGE_READWRITE);
    if (pThreadProcRemote == NULL)
    {
        std::wcerr << L"VirtualAllocEx failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    if (!WriteProcessMemory(hWrappedProcess, pThreadProcRemote, pThreadProc, nSizeOfThreadProc,
                            &nBytesWritten))
    {
        std::wcerr << L"WriteProcessMemory failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    DWORD nOldProtection;
    if (!VirtualProtectEx(hWrappedProcess, pThreadProcRemote, nSizeOfThreadProc, PAGE_EXECUTE,
                          &nOldProtection))
    {
        std::wcerr << L"VirtualProtectEx failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    if (bDebug)
    {
        // Give the developer a chance to debug the wrapped program, too.
        std::wcout << L"Waiting for you to attach a debugger to the wrapped "
                   << baseName(sWrappedFileName) << " process "
                   << std::to_wstring(GetProcessId(hWrappedProcess)) << L"\n";
        std::wcout << L"When done with that, set the bWait in the " << baseName(sMyFileName)
                   << L" process to false" << std::endl;
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
        std::wcerr << L"CreateRemoteThread failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    WaitForSingleObject(hThread, INFINITE);

    SIZE_T nBytesRead;
    if (!ReadProcessMemory(hWrappedProcess, pParamRemote, &aParam, sizeof(aParam), &nBytesRead))
    {
        std::wcerr << L"ReadProcessMemory failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    DWORD nExitCode;
    if (!GetExitCodeThread(hThread, &nExitCode))
    {
        std::wcerr << L"GetExitCodeThread failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }
    if (!nExitCode)
    {
        std::wcerr << L"Injected thread failed: ";

        if (!aParam.mbPassedSizeCheck)
        {
            if (aParam.mnLastError != 0)
                std::wcerr << WindowsErrorString(aParam.mnLastError) << L"\n";
            else
                std::wcerr << L"Mismatched parameter structure sizes, COLEAT build or installation "
                              L"problem.\n";
        }
        else
        {
            if (aParam.msErrorExplanation[0])
                std::wcerr << aParam.msErrorExplanation << L"\n";
            else
                std::wcerr << WindowsErrorString(aParam.mnLastError);
        }

        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    CloseHandle(hThread);

    DWORD nPreviousSuspendCount = ResumeThread(hWrappedThread);
    if (nPreviousSuspendCount == (DWORD)-1)
    {
        std::wcerr << L"ResumeThread failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }
    else if (nPreviousSuspendCount == 0)
    {
        std::wcerr << L"Huh, thread was not suspended?\n";
    }
    else if (nPreviousSuspendCount > 1)
    {
        std::wcerr << L"Thread still suspended after ResumeThread\n";
        TerminateProcess(hWrappedProcess, 1);
        WaitForSingleObject(hWrappedProcess, INFINITE);
        std::exit(1);
    }

    // Wait for the process to finish.
    WaitForSingleObject(hWrappedProcess, INFINITE);

    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

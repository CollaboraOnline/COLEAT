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

#include <Windows.h>

#pragma warning(pop)

#include "exewrapper.hpp"
#include "utils.hpp"

inline bool operator<(const IID& a, const IID& b) { return std::memcmp(&a, &b, sizeof(a)) < 0; }

static void Usage(wchar_t** argv)
{
    const wchar_t* const pBackslash = wcsrchr(argv[0], L'\\');
    const wchar_t* const pSlash = wcsrchr(argv[0], L'/');
    const wchar_t* const pProgram
        = (pBackslash && pSlash)
              ? ((pBackslash > pSlash) ? (pBackslash + 1) : (pSlash + 1))
              : (pBackslash ? (pBackslash + 1) : (pSlash ? (pSlash + 1) : argv[0]));

    std::wcerr
        << L"Usage: " << pProgram
        << L" [options] exeFile [arguments...]\n"
           "  Options:\n"
           "    -m ifaceA:ifaceB:ifaceC  Map coclass A with dedault interface B to\n"
           "                             replacement interface C. Option can be repeated.\n"
           "                             Each IID is in the normal brace-enclosed format.\n"
           "    -M file                  file contains interface mappings, like the -m option.\n"
           "\n";
    std::exit(1);
}

inline std::int16_t operator"" _short(unsigned long long value)
{
    return static_cast<std::int16_t>(value);
}

extern "C" {

// This small function is injected into the executable. It is passed a parameter block. It must be
// carefully written to be position-independent and is not able to call any imported functions, as
// the mechanism to access such, import tables etc, obviously would not work when the just function
// code is bluntly written into a completely different process. But that's OK, we can pass pointers
// to functions from kernel32.dll at least in the parameter block.

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

#define p1(x) x x
#define p2(x) p1(p1(x))
#define p3(x) p2(p2(x))

    // I tried to use inline __asm to place a public label endOfFun here, but
    // doesn't seem possible.

    // Instead place a large bunch of nops here, and search for them, and then search for the
    // "ret" instruction. As we turn off optimization for this function, no funky reordering or
    // other crack should be going on.

    p3(__asm { nop });

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

    int argi = 1;

    std::map<IID, InterfaceMapping> aInterfaceMap;

    while (argi < argc && argv[argi][0] == L'-')
    {
        switch (argv[argi][1])
        {
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

    if (argc - argi < 1)
        Usage(argv);

    if (aInterfaceMap.size() > ThreadProcParam::NIIDMAP)
    {
        std::wcerr << L"At most " << ThreadProcParam::NIIDMAP << " IID mappings possible\n";
        std::exit(1);
    }

    const unsigned char* const pThreadProc = (unsigned char*)&threadProc;
    const unsigned char* pRover = pThreadProc;

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
        std::exit(1);
    }
    pRover += 4;

    const SIZE_T nSizeOfThreadProc = (SIZE_T)(pRover - pThreadProc);

    // Get our exe pathname.

    const DWORD NFILENAME = 1000;
    wchar_t sFileName[NFILENAME];

    // NFILENAME-20 let the longer name of our DLL fit.
    DWORD nSize = GetModuleFileNameW(NULL, sFileName, NFILENAME - 20);
    if (nSize == NFILENAME - 20)
    {
        std::wcerr << L"Pathname of this exe ridiculously long\n";
        std::exit(1);
    }

    // Path to the DLL to inject.
    wchar_t* pLastDot = wcsrchr(sFileName, L'.');
    if (pLastDot == NULL)
    {
        std::wcerr << L"No period in " << sFileName << L"?\n";
        std::exit(1);
    }
    wcscpy(pLastDot, L"-injected.dll");

    // Start the process to wrap. Start it as suspended.

    std::wstring sCommandLine;
    for (int i = argi; i < argc; ++i)
    {
        if (i > argi)
            sCommandLine += L" ";
        sCommandLine += argv[i];
    }

    wchar_t* pCommandLine = _wcsdup(sCommandLine.data());

    STARTUPINFOW aStartupInfo;
    PROCESS_INFORMATION aProcessInfo;

    std::memset(&aStartupInfo, 0, sizeof(aStartupInfo));
    aStartupInfo.cb = sizeof(aStartupInfo);
    aStartupInfo.dwFlags = STARTF_USESTDHANDLES;

    if (!DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_INPUT_HANDLE), GetCurrentProcess(),
                         &aStartupInfo.hStdInput, 0, TRUE, DUPLICATE_SAME_ACCESS)
        || !DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_OUTPUT_HANDLE),
                            GetCurrentProcess(), &aStartupInfo.hStdOutput, 0, TRUE,
                            DUPLICATE_SAME_ACCESS)
        || !DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_ERROR_HANDLE),
                            GetCurrentProcess(), &aStartupInfo.hStdError, 0, TRUE,
                            DUPLICATE_SAME_ACCESS))
    {
        std::wcerr << L"DuplicateHandle failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    if (!CreateProcessW(NULL, pCommandLine, NULL, NULL, TRUE, CREATE_SUSPENDED, NULL, NULL,
                        &aStartupInfo, &aProcessInfo))
    {
        std::wcerr << L"CreateProcess failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

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
    wcscpy(aParam.msFileName, sFileName);

    for (auto i : aInterfaceMap)
    {
        aParam.mvInterfaceMap[aParam.mnMappings] = i.second;
        ++aParam.mnMappings;
    }

    void* pParamRemote
        = VirtualAllocEx(aProcessInfo.hProcess, NULL, sizeof(aParam), MEM_COMMIT, PAGE_READWRITE);
    if (pParamRemote == NULL)
    {
        std::wcerr << L"VirtualAllocEx failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    SIZE_T nBytesWritten;
    if (!WriteProcessMemory(aProcessInfo.hProcess, pParamRemote, &aParam, sizeof(aParam),
                            &nBytesWritten))
    {
        std::wcerr << L"WriteProcessMemory failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    void* pThreadProcRemote = VirtualAllocEx(aProcessInfo.hProcess, NULL, nSizeOfThreadProc,
                                             MEM_COMMIT, PAGE_READWRITE);
    if (pThreadProcRemote == NULL)
    {
        std::wcerr << L"VirtualAllocEx failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    if (!WriteProcessMemory(aProcessInfo.hProcess, pThreadProcRemote, pThreadProc,
                            nSizeOfThreadProc, &nBytesWritten))
    {
        std::wcerr << L"WriteProcessMemory failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    DWORD nOldProtection;
    if (!VirtualProtectEx(aProcessInfo.hProcess, pThreadProcRemote, nSizeOfThreadProc, PAGE_EXECUTE,
                          &nOldProtection))
    {
        std::wcerr << L"VirtualProtectEx failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    FunPtr pProc;
    pProc.pVoid = pThreadProcRemote;
    HANDLE hThread = CreateRemoteThread(aProcessInfo.hProcess, NULL, 0, pProc.pStartRoutine,
                                        pParamRemote, 0, NULL);
    if (!hThread)
    {
        std::wcerr << L"CreateRemoteThread failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    WaitForSingleObject(hThread, INFINITE);

    SIZE_T nBytesRead;
    if (!ReadProcessMemory(aProcessInfo.hProcess, pParamRemote, &aParam, sizeof(aParam),
                           &nBytesRead))
    {
        std::wcerr << L"ReadProcessMemory failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    DWORD nExitCode;
    if (!GetExitCodeThread(hThread, &nExitCode))
    {
        std::wcerr << L"GetExitCodeThread failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }
    if (!nExitCode)
    {
        std::wcerr << L"Injected thread failed";
        // The only way it can fail without filling in the msErrorExplanation is if the initial
        // sanity check of ThreadProcParam::mnSize fails.
        if (aParam.msErrorExplanation[0] == L'\0')
            std::wcerr << L": Mismatched thread procedure parameter structure?\n";
        else
            std::wcerr << L": " << aParam.msErrorExplanation << L"\n";
        std::exit(1);
    }

    CloseHandle(hThread);

    DWORD nPreviousSuspendCount = ResumeThread(aProcessInfo.hThread);
    if (nPreviousSuspendCount == (DWORD)-1)
    {
        std::wcerr << L"ResumeThread failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }
    else if (nPreviousSuspendCount == 0)
    {
        std::wcerr << L"Huh, thread was not suspended?\n";
    }
    else if (nPreviousSuspendCount > 1)
    {
        std::wcerr << L"Thread still suspended after ResumeThread\n";
        std::exit(1);
    }

    // Wait for the process to finish.
    WaitForSingleObject(aProcessInfo.hThread, INFINITE);

    CloseHandle(aProcessInfo.hThread);
    CloseHandle(aProcessInfo.hProcess);

    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

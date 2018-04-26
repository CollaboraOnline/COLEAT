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

static bool bDebug = false;

static void Usage(wchar_t** argv)
{
    std::wcerr
        << L"Usage: " << programName(argv[0])
        << L" [options] exeFile [arguments...]\n"
           "  Options:\n"
           "    -m ifaceA:ifaceB:ifaceC  Map coclass A with dedault interface B to\n"
           "                             replacement interface C. Option can be repeated.\n"
           "                             Each IID is in the normal brace-enclosed format.\n"
           "    -M file                  file contains interface mappings, like the -m option.\n"
           "\n";
    std::exit(1);
}

int wmain(int argc, wchar_t** argv)
{
    tryToEnsureStdHandlesOpen();

    int argi = 1;

    // Just ensure syntax is right, we don't need to actualy handle the arguments in this program,
    // they are passed on to exewrapper.
    while (argi < argc && argv[argi][0] == L'-')
    {
        switch (argv[argi][1])
        {
            case L'd':
            {
                // secret debug switch
                bDebug = true;
                break;
            }
            case L'm':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                argi++;
                break;
            }
            case L'M':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
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

    // Get our exe pathname.

    const DWORD NFILENAME = 1000;
    wchar_t sFileName[NFILENAME];

    // NFILENAME-20 let the longer name of exewrapper fit.
    DWORD nSize = GetModuleFileNameW(NULL, sFileName, NFILENAME - 20);
    if (nSize == NFILENAME - 20)
    {
        std::wcerr << L"Pathname of this exe ridiculously long\n";
        std::exit(1);
    }

    // Construct name of the right exewrapper.
    wchar_t* pLastBackslash = wcsrchr(sFileName, L'\\');
    wchar_t* pLastSlash = wcsrchr(sFileName, L'/');
    if (pLastBackslash == NULL && pLastSlash == NULL)
    {
        std::wcerr << L"Not a full path: '" << sFileName << L"'?\n";
        std::exit(1);
    }
    wchar_t* pAfterSlash
        = (pLastSlash == NULL ? pLastBackslash
                              : (pLastBackslash == NULL
                                     ? pLastSlash
                                     : (pLastSlash > pLastBackslash ? pLastSlash : pLastBackslash)))
          + 1;

    // Start the process to wrap. Start it as suspended.

    std::wstring sCommandLine;
    for (int i = argi; i < argc; ++i)
    {
        if (i > argi)
            sCommandLine += L" ";
        sCommandLine += argv[i];
    }

    STARTUPINFOW aWrappedProcessStartupInfo;
    PROCESS_INFORMATION aWrappedProcessInfo;

    std::memset(&aWrappedProcessStartupInfo, 0, sizeof(aWrappedProcessStartupInfo));
    aWrappedProcessStartupInfo.cb = sizeof(aWrappedProcessStartupInfo);
    aWrappedProcessStartupInfo.dwFlags = STARTF_USESTDHANDLES;

    if (!DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_INPUT_HANDLE), GetCurrentProcess(),
                         &aWrappedProcessStartupInfo.hStdInput, 0, TRUE, DUPLICATE_SAME_ACCESS)
        || !DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_OUTPUT_HANDLE),
                            GetCurrentProcess(), &aWrappedProcessStartupInfo.hStdOutput, 0, TRUE,
                            DUPLICATE_SAME_ACCESS)
        || !DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_ERROR_HANDLE),
                            GetCurrentProcess(), &aWrappedProcessStartupInfo.hStdError, 0, TRUE,
                            DUPLICATE_SAME_ACCESS))
    {
        std::wcerr << L"DuplicateHandle failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    wchar_t* pCommandLine = _wcsdup(sCommandLine.data());

    if (!CreateProcessW(NULL, pCommandLine, NULL, NULL, TRUE, CREATE_SUSPENDED, NULL, NULL,
                        &aWrappedProcessStartupInfo, &aWrappedProcessInfo))
    {
        std::wcerr << L"CreateProcess failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    free(pCommandLine);

    BOOL bIsWow64;
    if (!IsWow64Process(aWrappedProcessInfo.hProcess, &bIsWow64))
    {
        std::wcerr << L"IsWow64Process failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(aWrappedProcessInfo.hProcess, 1);
        WaitForSingleObject(aWrappedProcessInfo.hProcess, INFINITE);
        std::exit(1);
    }

    // bIsWow64 is TRUE iff the wrapped process is a 32-bit process running under WOW64, i.e. on a
    // 64-bit OS.
    //
    // If bIsWow64 is FALSE we don't know based on that whether it is a 32-bit process on a 32-bit
    // OS, or a 64-bit process on a 64-bit OS.

    SYSTEM_INFO aSystemInfo;
    GetNativeSystemInfo(&aSystemInfo);

    // If we are on a 32-bit OS, then obviously the wrapped process must be 32-bit, too. If we are
    // on a 64-bit OS, then if bIsWow64 is FALSE, the wrapped process must be 64-bit.

    const wchar_t* pArchOfWrappedProcess
        = (bIsWow64 || aSystemInfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_INTEL) ? L"x86"
                                                                                           : L"x64";

    // Duplicate the handles to the wrapped process and its start thread as inheritable
    HANDLE hInheritableWrappedProcess;
    HANDLE hInheritableWrappedThread;
    if (!DuplicateHandle(GetCurrentProcess(), aWrappedProcessInfo.hProcess, GetCurrentProcess(),
                         &hInheritableWrappedProcess, 0, TRUE, DUPLICATE_SAME_ACCESS)
        || !DuplicateHandle(GetCurrentProcess(), aWrappedProcessInfo.hThread, GetCurrentProcess(),
                            &hInheritableWrappedThread, 0, TRUE, DUPLICATE_SAME_ACCESS))
    {
        std::wcerr << L"DuplicateHandle failed: " << WindowsErrorString(GetLastError()) << L"\n";
        TerminateProcess(aWrappedProcessInfo.hProcess, 1);
        WaitForSingleObject(aWrappedProcessInfo.hProcess, INFINITE);
        std::exit(1);
    }

    // Start exewrapper, pass it the inheritable handle to the wrapped process and its startup
    // thread, and after those the same command-line arguments we got.

    wcscpy(pAfterSlash, L"exewrapper-");
    wcscat(pAfterSlash, pArchOfWrappedProcess);
    wcscat(pAfterSlash, L".exe");

    // exewrapper's command line is: inherited process handle, inherited main thread handle followed by our options. The wrapped program and
    // its parameters don't need to be passed to exewrapper.

    sCommandLine = std::wstring(sFileName) + L" "
                   + std::to_wstring((intptr_t)hInheritableWrappedProcess) + L" "
                   + std::to_wstring((intptr_t)hInheritableWrappedThread);

    for (int i = 1; i < argi; ++i)
    {
        sCommandLine += L" ";
        sCommandLine += argv[i];
    }

    pCommandLine = _wcsdup(sCommandLine.data());

    STARTUPINFO aExeWrapperStartupInfo;
    PROCESS_INFORMATION aExeWrapperProcessInfo;

    std::memset(&aExeWrapperStartupInfo, 0, sizeof(aExeWrapperStartupInfo));
    aExeWrapperStartupInfo.cb = sizeof(aExeWrapperStartupInfo);
    aExeWrapperStartupInfo.dwFlags = STARTF_USESTDHANDLES;

    if (!DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_INPUT_HANDLE), GetCurrentProcess(),
                         &aExeWrapperStartupInfo.hStdInput, 0, TRUE, DUPLICATE_SAME_ACCESS)
        || !DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_OUTPUT_HANDLE),
                            GetCurrentProcess(), &aExeWrapperStartupInfo.hStdOutput, 0, TRUE,
                            DUPLICATE_SAME_ACCESS)
        || !DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_ERROR_HANDLE),
                            GetCurrentProcess(), &aExeWrapperStartupInfo.hStdError, 0, TRUE,
                            DUPLICATE_SAME_ACCESS))
    {
        std::wcerr << L"DuplicateHandle failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    if (!CreateProcessW(NULL, pCommandLine, NULL, NULL, TRUE, 0, NULL, NULL,
                        &aExeWrapperStartupInfo, &aExeWrapperProcessInfo))
    {
        std::wcerr << L"CreateProcess failed: " << WindowsErrorString(GetLastError()) << L"\n";
        std::exit(1);
    }

    free(pCommandLine);

    // Wait for the exewrapper process to finish. (It will wait for the wrapped process to finish.)
    WaitForSingleObject(aExeWrapperProcessInfo.hProcess, INFINITE);

    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

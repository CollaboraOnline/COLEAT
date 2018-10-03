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
#include <cstdio>
#include <cstdlib>
#include <cwchar>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <map>
#include <sstream>
#include <string>

#include <io.h>
#include <time.h>

#include <Windows.h>

#pragma warning(pop)

#include "coleat-version.h"
#include "coleat-git-version.h"
#include "exewrapper.hpp"
#include "utils.hpp"

#define STRINGIFY2(x) #x
#define STRINGIFY(x) STRINGIFY2(x)

static bool bDebug = false;
static bool bDidAllocConsole;

static void Usage(wchar_t** argv)
{
    tryToEnsureStdHandlesOpen(bDidAllocConsole);

    std::cout << "Usage: " << convertUTF16ToUTF8(programName(argv[0]))
              << " [options] program [arguments...]\n"
                 "\n"
                 "  Options:\n"
                 "    -n                           no redirection to replacement app\n"
                 "    -o file                      output file (default: stdout, in new console if "
                 "necessary)\n"
                 "    -t                           terse trace output\n"
                 "    -v                           verbose logging of internal operation\n"
                 "    -V                           print COLEAT version information\n";
    if (bDidAllocConsole)
        waitForAnyKey();
    std::exit(1);
}

int wmain(int argc, wchar_t** argv)
{
    wchar_t* pOutputFile = nullptr;

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
            case L'n':
                break;
            case L'o':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                pOutputFile = argv[argi + 1];
                argi++;
                break;
            }
            case L't':
                break;
            case L'v':
                break;
            case L'V':
                std::cout << "COLEAT " << COLEAT_VERSION_MAJOR << "." << COLEAT_VERSION_MINOR
                          << " (" STRINGIFY(COLEAT_GIT_HEAD) ")" << std::endl;
                std::exit(0);
                break;
            default:
                Usage(argv);
        }
        argi++;
    }

    if (argc - argi < 1)
        Usage(argv);

    HANDLE hOutputHandle;

    if (pOutputFile != nullptr)
    {
        const int NTIMEBUF = 100;
        wchar_t sTimeBuf[NTIMEBUF];
        time_t nNow = time(nullptr);
        struct tm aNow;
        localtime_s(&aNow, &nNow);

        // Expand "variables" in the output file name.

        std::wstring sOutputFile(pOutputFile);
        while (true)
        {
            std::size_t i;
            i = sOutputFile.find(L"%t");
            if (i != std::wstring::npos)
            {
                wcsftime(sTimeBuf, NTIMEBUF, L"%Y%m%d.%H%M%S", &aNow);
                sOutputFile = sOutputFile.substr(0, i) + sTimeBuf + sOutputFile.substr(i + 2);
                continue;
            }
            break;
        }

        FILE* aStream;
        if (_wfreopen_s(&aStream, sOutputFile.c_str(), L"a", stdout) != 0)
        {
            tryToEnsureStdHandlesOpen(bDidAllocConsole);

            std::cout << "Could not open given output file "
                      << convertUTF16ToUTF8(sOutputFile.c_str()) << " for writing\n";
            if (bDidAllocConsole)
                waitForAnyKey();
            std::exit(1);
        }

        // If appending to a file, write a header.

        if (wcsftime(sTimeBuf, NTIMEBUF, L"%F %H:%M:%S", &aNow) != 0)
            std::cout << "=== COLEAT " << COLEAT_VERSION_MAJOR << "." << COLEAT_VERSION_MINOR
                      << " (" STRINGIFY(COLEAT_GIT_HEAD) ") output at "
                      << convertUTF16ToUTF8(sTimeBuf) << " ===" << std::endl;

        hOutputHandle = (HANDLE)_get_osfhandle(_fileno(stdout));
    }
    else
    {
        tryToEnsureStdHandlesOpen(bDidAllocConsole);
        hOutputHandle = GetStdHandle(STD_OUTPUT_HANDLE);
    }

    // Get our exe pathname.

    const DWORD NFILENAME = 1000;
    wchar_t sFileName[NFILENAME];

    // NFILENAME-20 let the longer name of exewrapper fit.
    DWORD nSize = GetModuleFileNameW(NULL, sFileName, NFILENAME - 20);
    if (nSize == NFILENAME - 20)
    {
        std::cout << "Pathname of this exe ridiculously long\n";
        if (bDidAllocConsole)
            waitForAnyKey();
        std::exit(1);
    }

    // Construct name of the right exewrapper.
    wchar_t* pLastBackslash = wcsrchr(sFileName, L'\\');
    wchar_t* pLastSlash = wcsrchr(sFileName, L'/');
    if (pLastBackslash == NULL && pLastSlash == NULL)
    {
        std::cout << "Not a full path: '" << convertUTF16ToUTF8(sFileName) << "'?\n";
        if (bDidAllocConsole)
            waitForAnyKey();
        std::exit(1);
    }
    wchar_t* pAfterSlash
        = (pLastSlash == NULL ? pLastBackslash
                              : (pLastBackslash == NULL
                                     ? pLastSlash
                                     : (pLastSlash > pLastBackslash ? pLastSlash : pLastBackslash)))
          + 1;

    std::cout << "Command line: " << convertUTF16ToUTF8(GetCommandLineW()) << "\n" << std::endl;

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

    // We only *need* STD_OUTPUT_HANDLE
    if (!DuplicateHandle(GetCurrentProcess(), hOutputHandle, GetCurrentProcess(),
                         &aWrappedProcessStartupInfo.hStdOutput, 0, TRUE, DUPLICATE_SAME_ACCESS))
    {
        std::cout << "DuplicateHandle of output handle " << hOutputHandle
                  << " for wrapped process failed: " << WindowsErrorString(GetLastError()) << "\n";
        if (bDidAllocConsole)
            waitForAnyKey();
        std::exit(1);
    }

    // Ignore errors duplicatig the STD_INPUT_HANDLE or STD_ERROR_HANDLE
    DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_INPUT_HANDLE), GetCurrentProcess(),
                    &aWrappedProcessStartupInfo.hStdInput, 0, TRUE, DUPLICATE_SAME_ACCESS);
    DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_ERROR_HANDLE), GetCurrentProcess(),
                    &aWrappedProcessStartupInfo.hStdError, 0, TRUE, DUPLICATE_SAME_ACCESS);

    wchar_t* pCommandLine = _wcsdup(sCommandLine.data());

    if (!CreateProcessW(NULL, pCommandLine, NULL, NULL, TRUE, CREATE_SUSPENDED, NULL, NULL,
                        &aWrappedProcessStartupInfo, &aWrappedProcessInfo))
    {
        std::cout << "CreateProcess(" << convertUTF16ToUTF8(pCommandLine)
                  << ") failed: " << WindowsErrorString(GetLastError()) << "\n";
        if (bDidAllocConsole)
            waitForAnyKey();
        std::exit(1);
    }

    free(pCommandLine);

    BOOL bIsWow64;
    if (!IsWow64Process(aWrappedProcessInfo.hProcess, &bIsWow64))
    {
        std::cout << "IsWow64Process failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(aWrappedProcessInfo.hProcess, 1);
        WaitForSingleObject(aWrappedProcessInfo.hProcess, INFINITE);
        if (bDidAllocConsole)
            waitForAnyKey();
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
        std::cout << "DuplicateHandle of wrapped process or thread handle failed: "
                  << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(aWrappedProcessInfo.hProcess, 1);
        WaitForSingleObject(aWrappedProcessInfo.hProcess, INFINITE);
        if (bDidAllocConsole)
            waitForAnyKey();
        std::exit(1);
    }

    // Start exewrapper, pass it the inheritable handle to the wrapped process and its startup
    // thread, and after those the same command-line arguments we got.

    wcscpy_s(pAfterSlash, NFILENAME - (std::size_t)(pAfterSlash - sFileName), L"exewrapper-");
    wcscat_s(pAfterSlash, NFILENAME - wcslen(sFileName), pArchOfWrappedProcess);
    wcscat_s(pAfterSlash, NFILENAME - wcslen(sFileName), L".exe");

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

    if (!DuplicateHandle(GetCurrentProcess(), hOutputHandle, GetCurrentProcess(),
                         &aExeWrapperStartupInfo.hStdOutput, 0, TRUE, DUPLICATE_SAME_ACCESS))
    {
        std::cout << "DuplicateHandle of output handle " << hOutputHandle
                  << " for exewrapper failed: " << WindowsErrorString(GetLastError()) << "\n";
        if (bDidAllocConsole)
            waitForAnyKey();
        std::exit(1);
    }

    DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_INPUT_HANDLE), GetCurrentProcess(),
                    &aExeWrapperStartupInfo.hStdInput, 0, TRUE, DUPLICATE_SAME_ACCESS);

    DuplicateHandle(GetCurrentProcess(), GetStdHandle(STD_ERROR_HANDLE), GetCurrentProcess(),
                    &aExeWrapperStartupInfo.hStdError, 0, TRUE, DUPLICATE_SAME_ACCESS);

    if (!CreateProcessW(NULL, pCommandLine, NULL, NULL, TRUE, 0, NULL, NULL,
                        &aExeWrapperStartupInfo, &aExeWrapperProcessInfo))
    {
        std::cout << "CreateProcess(" << convertUTF16ToUTF8(sFileName)
                  << ") failed: " << WindowsErrorString(GetLastError()) << "\n";
        TerminateProcess(aWrappedProcessInfo.hProcess, 1);
        WaitForSingleObject(aWrappedProcessInfo.hProcess, INFINITE);
        if (bDidAllocConsole)
            waitForAnyKey();
        std::exit(1);
    }

    free(pCommandLine);

    // Wait for the exewrapper process to finish. (It will wait for the wrapped process to finish.)
    WaitForSingleObject(aExeWrapperProcessInfo.hProcess, INFINITE);

    if (bDidAllocConsole)
        waitForAnyKey();

    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

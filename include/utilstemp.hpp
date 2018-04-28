/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora Office OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

// Temporary. This file will go away once I get around to stop using std::wcout and std::wcerr in a
// few files, including utils.hpp, which is where this stuff will be.


#ifndef INCLUDED_UTILSTEMP_HPP
#define INCLUDED_UTILSTEMP_HPP

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>
#include <vector>

#include <Windows.h>
#include <initguid.h>

#pragma warning(pop)

inline std::string to_hex(uint64_t n, int w = 0)
{
    std::stringstream aStringStream;
    aStringStream << std::setfill('0') << std::setw(w) << std::hex << n;
    return aStringStream.str();
}

inline std::string WindowsErrorString(DWORD nErrorCode)
{
    LPWSTR pMsgBuf;

    if (FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM
                           | FORMAT_MESSAGE_IGNORE_INSERTS,
                       nullptr, nErrorCode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                       reinterpret_cast<LPWSTR>(&pMsgBuf), 0, nullptr)
        == 0)
    {
        return to_hex(nErrorCode);
    }

    std::string sResult
        = std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t>().to_bytes(pMsgBuf);
    if (sResult.length() > 2 && sResult.substr(sResult.length() - 2, 2) == "\r\n")
        sResult.resize(sResult.length() - 2);

    HeapFree(GetProcessHeap(), 0, pMsgBuf);

    return sResult;
}

inline std::string WindowsErrorStringFromHRESULT(HRESULT nResult)
{
    // Return common HRESULT codes symbolically. This is for developer use anyway, much easier to
    // read "E_NOTIMPL" than the English prose description.
    switch (nResult)
    {
        case S_OK:
            return "S_OK";
        case S_FALSE:
            return "S_FALSE";
        case E_UNEXPECTED:
            return "E_UNEXPECTED";
        case E_NOTIMPL:
            return "E_NOTIMPL";
        case E_OUTOFMEMORY:
            return "E_OUTOFMEMORY";
        case E_INVALIDARG:
            return "E_INVALIDARG";
        case E_NOINTERFACE:
            return "E_NOINTERFACE";
        case E_POINTER:
            return "E_POINTER";
        case E_HANDLE:
            return "E_HANDLE";
        case E_ABORT:
            return "E_ABORT";
        case E_FAIL:
            return "E_FAIL";
        case E_ACCESSDENIED:
            return "E_ACCESSDENIED";
    }

    // See https://blogs.msdn.microsoft.com/oldnewthing/20061103-07/?p=29133
    // Also https://social.msdn.microsoft.com/Forums/vstudio/en-US/c33d9a4a-1077-4efd-99e8-0c222743d2f8
    // (which refers to https://msdn.microsoft.com/en-us/library/aa382475)
    // explains why can't we just reinterpret_cast HRESULT to DWORD Win32 error:
    // we might actually have a Win32 error code converted using HRESULT_FROM_WIN32 macro

    DWORD nErrorCode = DWORD(nResult);
    if (HRESULT(nResult & 0xFFFF0000) == MAKE_HRESULT(SEVERITY_ERROR, FACILITY_WIN32, 0)
        || nResult == S_OK)
    {
        nErrorCode = (DWORD)HRESULT_CODE(nResult);
        // https://msdn.microsoft.com/en-us/library/ms679360 mentions that the codes might have
        // high word bits set (e.g., bit 29 could be set if error comes from a 3rd-party library).
        // So try to restore the original error code to avoid wrong error messages
        DWORD nLastError = GetLastError();
        if ((nLastError & 0xFFFF) == nErrorCode)
            nErrorCode = nLastError;
    }

    return WindowsErrorString(nErrorCode);
}

#endif // INCLUDED_UTILSTEMP_HPP

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

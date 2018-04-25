/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */
/*
 * This file is part of Collabora Office OLE Automation Translator.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#ifndef INCLUDED_UTILS_HPP
#define INCLUDED_UTILS_HPP

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

inline std::wstring to_hex(uint64_t n, int w = 0)
{
    std::wstringstream aStringStream;
    aStringStream << std::setfill(L'0') << std::setw(w) << std::hex << n;
    return aStringStream.str();
}

inline bool GetWindowsErrorString(DWORD nErrorCode, LPWSTR* pPMsgBuf)
{
    if (FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM
                           | FORMAT_MESSAGE_IGNORE_INSERTS,
                       nullptr, nErrorCode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                       reinterpret_cast<LPWSTR>(pPMsgBuf), 0, nullptr)
        == 0)
    {
        return false;
    }

    std::size_t nLen = std::wcslen(*pPMsgBuf);
    if (nLen > 2 && (*pPMsgBuf)[nLen-2] == L'\r' && (*pPMsgBuf)[nLen-1] == L'\n')
        (*pPMsgBuf)[nLen-2] = L'\0';
    return true;
}

inline std::wstring WindowsErrorString(DWORD nErrorCode)
{
    LPWSTR pMsgBuf;

    if (!GetWindowsErrorString(nErrorCode, &pMsgBuf))
    {
        return to_hex(nErrorCode);
    }

    std::wstring sResult(pMsgBuf);

    HeapFree(GetProcessHeap(), 0, pMsgBuf);

    return sResult;
}

inline std::wstring WindowsErrorStringFromHRESULT(HRESULT nResult)
{
    // Return common HRESULT codes symbolically. This is for developer use anyway, much easier to
    // read "E_NOTIMPL" than the English prose description.
    switch (nResult)
    {
        case S_OK:
            return L"S_OK";
        case S_FALSE:
            return L"S_FALSE";
        case E_UNEXPECTED:
            return L"E_UNEXPECTED";
        case E_NOTIMPL:
            return L"E_NOTIMPL";
        case E_OUTOFMEMORY:
            return L"E_OUTOFMEMORY";
        case E_INVALIDARG:
            return L"E_INVALIDARG";
        case E_NOINTERFACE:
            return L"E_NOINTERFACE";
        case E_POINTER:
            return L"E_POINTER";
        case E_HANDLE:
            return L"E_HANDLE";
        case E_ABORT:
            return L"E_ABORT";
        case E_FAIL:
            return L"E_FAIL";
        case E_ACCESSDENIED:
            return L"E_ACCESSDENIED";
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

inline void tryToEnsureStdHandlesOpen()
{
    // Make sure we have a stdout for debugging output, for now
    if (GetStdHandle(STD_OUTPUT_HANDLE) == NULL)
    {
        STARTUPINFOW aStartupInfo;
        aStartupInfo.cb = sizeof(aStartupInfo);
        GetStartupInfoW(&aStartupInfo);
        if ((aStartupInfo.dwFlags & STARTF_USESTDHANDLES) == STARTF_USESTDHANDLES &&
            aStartupInfo.hStdInput != INVALID_HANDLE_VALUE &&
            aStartupInfo.hStdInput != NULL &&
            aStartupInfo.hStdOutput != INVALID_HANDLE_VALUE &&
            aStartupInfo.hStdOutput != NULL &&
            aStartupInfo.hStdError != INVALID_HANDLE_VALUE &&
            aStartupInfo.hStdError != NULL)
        {
            // If standard handles had been passed to this process, use them
            SetStdHandle(STD_INPUT_HANDLE, aStartupInfo.hStdInput);
            SetStdHandle(STD_OUTPUT_HANDLE, aStartupInfo.hStdOutput);
            SetStdHandle(STD_ERROR_HANDLE, aStartupInfo.hStdError);
        }
        else
        {
            // Try to attach parent console; on error try to create new.
            // If this process already has its console, these will simply fail.
            if (!AttachConsole(ATTACH_PARENT_PROCESS))
                AllocConsole();
        }
        std::freopen("CON", "r", stdin);
        std::freopen("CON", "w", stdout);
        std::freopen("CON", "w", stderr);
        std::ios::sync_with_stdio(true);
    }
}

namespace
{
DEFINE_GUID(IID_IdentityUnmarshal, 0x0000001B, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x46);
}

template <typename traits>
inline std::basic_ostream<wchar_t, traits>& operator<<(std::basic_ostream<wchar_t, traits>& stream,
                                                     const IID& rIid)
{
    LPOLESTR pRiid;
    if (StringFromIID(rIid, &pRiid) != S_OK)
        return stream << L"?";

    stream << pRiid;

    DWORD nSize;
    if (RegGetValueW(HKEY_CLASSES_ROOT, std::wstring(L"CLSID\\").append(pRiid).data(), NULL,
                     RRF_RT_REG_SZ, NULL, NULL, &nSize)
        == ERROR_SUCCESS)
    {
        std::vector<wchar_t> sValue(nSize / 2);
        if (RegGetValueW(HKEY_CLASSES_ROOT, std::wstring(L"CLSID\\").append(pRiid).data(), NULL,
                         RRF_RT_REG_SZ, NULL, sValue.data(), &nSize)
            == ERROR_SUCCESS)
        {
            stream << L"=\"" << sValue.data() << L"\"";
        }
    }
    else if (RegGetValueW(HKEY_CLASSES_ROOT, std::wstring(L"Interface\\").append(pRiid).data(),
                          NULL, RRF_RT_REG_SZ, NULL, NULL, &nSize)
             == ERROR_SUCCESS)
    {
        std::vector<wchar_t> sValue(nSize / 2);
        if (RegGetValueW(HKEY_CLASSES_ROOT, std::wstring(L"Interface\\").append(pRiid).data(), NULL,
                         RRF_RT_REG_SZ, NULL, sValue.data(), &nSize)
            == ERROR_SUCCESS)
        {
            stream << L"=\"" << sValue.data() << L"\"";
        }
    }
    else
    {
        // Special case well-known interfaces that pop up a lot, but which don't have their name in
        // the Registry.

        if (IsEqualIID(rIid, IID_IMarshal))
            stream << L"=\"IMarshal\"";
        else if (IsEqualIID(rIid, IID_IMarshal2))
            stream << L"=\"IMarshal2\"";
        else if (IsEqualIID(rIid, IID_INoMarshal))
            stream << L"=\"INoMarshal\"";
        else if (IsEqualIID(rIid, IID_IdentityUnmarshal))
            stream << L"=\"IdentityUnmarshal\"";
        else if (IsEqualIID(rIid, IID_IFastRundown))
            stream << L"=\"IFastRundown\"";
        else if (IsEqualIID(rIid, IID_IStdMarshalInfo))
            stream << L"=\"IStdMarshalInfo\"";
        else if (IsEqualIID(rIid, IID_IAgileObject))
            stream << L"=\"IAgileObject\"";
        else if (IsEqualIID(rIid, IID_IExternalConnection))
            stream << L"=\"IExternalConnection\"";
        else if (IsEqualIID(rIid, IID_ICallFactory))
            stream << L"=\"ICallFactory\"";
    }

    CoTaskMemFree(pRiid);
    return stream;
}

#endif // INCLUDED_UTILS_HPP

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

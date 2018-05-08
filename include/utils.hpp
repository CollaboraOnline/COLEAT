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

#include <cctype>
#include <codecvt>
#include <cstdio>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>
#include <vector>

#include <Windows.h>
#include <initguid.h>

#pragma warning(pop)

inline bool operator<(const IID& a, const IID& b) { return std::memcmp(&a, &b, sizeof(a)) < 0; }

inline std::string convertUTF16ToUTF8(const wchar_t* pWchar)
{
    static std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t> aUTF16ToUTF8;

    return std::string(aUTF16ToUTF8.to_bytes(pWchar));
}

inline std::wstring convertUTF8ToUTF16(const char* pChar)
{
    static std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>, wchar_t> aUTF8ToUTF16;

    return std::wstring(aUTF8ToUTF16.from_bytes(pChar));
}

inline const wchar_t* baseName(const wchar_t* sPathname)
{
    const wchar_t* const pBackslash = wcsrchr(sPathname, L'\\');
    const wchar_t* const pSlash = wcsrchr(sPathname, L'/');
    const wchar_t* const pRetval
        = (pBackslash && pSlash)
              ? ((pBackslash > pSlash) ? (pBackslash + 1) : (pSlash + 1))
              : (pBackslash ? (pBackslash + 1) : (pSlash ? (pSlash + 1) : sPathname));
    return pRetval;
}

inline wchar_t* programName(const wchar_t* sPathname)
{
    wchar_t* pRetval = _wcsdup(baseName(sPathname));
    wchar_t* const pDot = std::wcsrchr(pRetval, L'.');
    if (pDot && _wcsicmp(pDot, L".exe") == 0)
        *pDot = L'\0';
    return pRetval;
}

inline std::string to_ullhex(uint64_t n, int w = 0)
{
    std::stringstream aStringStream;
    aStringStream << std::setfill('0') << std::setw(w) << std::uppercase << std::hex << n;
    return aStringStream.str();
}

inline std::string to_uhex(uint32_t n, int w = 0)
{
    std::stringstream aStringStream;
    aStringStream << std::setfill('0') << std::setw(w) << std::uppercase << std::hex << n;
    return aStringStream.str();
}

inline std::string to_hex(int32_t n, int w = 0) { return to_uhex((uint32_t)n, w); }

inline std::string IID_initializer(const IID& aIID)
{
    std::string sResult;
    sResult = "{0x" + to_uhex(aIID.Data1, 8) + ",0x" + to_uhex(aIID.Data2, 4) + ",0x"
              + to_uhex(aIID.Data3, 4);
    for (int i = 0; i < 8; ++i)
        sResult += ",0x" + to_hex(aIID.Data4[i], 2);
    sResult += "}";

    return sResult;
}

inline std::string IID_to_string(const IID& aIID)
{
    LPOLESTR pRiid;
    if (StringFromIID(aIID, &pRiid) != S_OK)
        return "?";
    return convertUTF16ToUTF8(pRiid);
}

inline std::string VARTYPE_to_string(VARTYPE nVt)
{
    if (nVt == VT_ILLEGAL)
        return "ILLEGAL";

    std::string sResult;

    if (nVt & VT_VECTOR)
        sResult += "VECTOR:";
    if (nVt & VT_ARRAY)
        sResult += "ARRAY:";
    if (nVt & VT_BYREF)
        sResult += "BYREF:";

    switch (nVt & VT_TYPEMASK)
    {
        case VT_EMPTY:
            sResult += "EMPTY";
            break;
        case VT_NULL:
            sResult += "NULL";
            break;
        case VT_I2:
            sResult += "I2";
            break;
        case VT_I4:
            sResult += "I4";
            break;
        case VT_R4:
            sResult += "R4";
            break;
        case VT_R8:
            sResult += "R8";
            break;
        case VT_CY:
            sResult += "CY";
            break;
        case VT_DATE:
            sResult += "DATE";
            break;
        case VT_BSTR:
            sResult += "BSTR";
            break;
        case VT_DISPATCH:
            sResult += "DISPATCH";
            break;
        case VT_ERROR:
            sResult += "ERROR";
            break;
        case VT_BOOL:
            sResult += "BOOL";
            break;
        case VT_VARIANT:
            sResult += "VARIANT";
            break;
        case VT_UNKNOWN:
            sResult += "UNKNOWN";
            break;
        case VT_DECIMAL:
            sResult += "DECIMAL";
            break;
        case VT_I1:
            sResult += "I1";
            break;
        case VT_UI1:
            sResult += "UI1";
            break;
        case VT_UI2:
            sResult += "UI2";
            break;
        case VT_UI4:
            sResult += "UI4";
            break;
        case VT_I8:
            sResult += "I8";
            break;
        case VT_UI8:
            sResult += "UI8";
            break;
        case VT_INT:
            sResult += "INT";
            break;
        case VT_UINT:
            sResult += "UINT";
            break;
        case VT_VOID:
            sResult += "VOID";
            break;
        case VT_HRESULT:
            sResult += "HRESULT";
            break;
        case VT_PTR:
            sResult += "PTR";
            break;
        case VT_SAFEARRAY:
            sResult += "SAFEARRAY";
            break;
        case VT_CARRAY:
            sResult += "CARRAY";
            break;
        case VT_USERDEFINED:
            sResult += "USERDEFINED";
            break;
        case VT_LPSTR:
            sResult += "LPSTR";
            break;
        case VT_LPWSTR:
            sResult += "LPWSTR";
            break;
        case VT_RECORD:
            sResult += "RECORD";
            break;
        case VT_INT_PTR:
            sResult += "INT_PTR";
            break;
        case VT_UINT_PTR:
            sResult += "UINT_PTR";
            break;
        case VT_FILETIME:
            sResult += "FILETIME";
            break;
        case VT_BLOB:
            sResult += "BLOB";
            break;
        case VT_STREAM:
            sResult += "STREAM";
            break;
        case VT_STORAGE:
            sResult += "STORAGE";
            break;
        case VT_STREAMED_OBJECT:
            sResult += "STREAMED_OBJECT";
            break;
        case VT_STORED_OBJECT:
            sResult += "STORED_OBJECT";
            break;
        case VT_BLOB_OBJECT:
            sResult += "BLOB_OBJECT";
            break;
        case VT_CF:
            sResult += "CF";
            break;
        case VT_CLSID:
            sResult += "CLSID";
            break;
        case VT_VERSIONED_STREAM:
            sResult += "VERSIONED_STREAM";
            break;
        case VT_BSTR_BLOB:
            sResult += "BSTR_BLOB";
            break;
        default:
            sResult += "?(" + std::to_string(nVt & VT_TYPEMASK) + ")";
            break;
    }
    return sResult;
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
    if (nLen > 2 && (*pPMsgBuf)[nLen - 2] == L'\r' && (*pPMsgBuf)[nLen - 1] == L'\n')
        (*pPMsgBuf)[nLen - 2] = L'\0';
    return true;
}

inline std::string WindowsErrorString(DWORD nErrorCode)
{
    LPWSTR pMsgBuf;

    if (!GetWindowsErrorString(nErrorCode, &pMsgBuf))
    {
        return to_uhex(nErrorCode, 8);
    }

    std::string sResult(convertUTF16ToUTF8(pMsgBuf));

    HeapFree(GetProcessHeap(), 0, pMsgBuf);

    return sResult;
}

inline std::string HRESULT_to_string(HRESULT nResult)
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
        default:
            return to_hex(nResult, 8);
    }
}

inline std::string WindowsErrorStringFromHRESULT(HRESULT nResult)
{
    std::string sSymbolic = HRESULT_to_string(nResult);

    if (!(sSymbolic.length() == 8 && std::isxdigit(sSymbolic[0]) && std::isxdigit(sSymbolic[1])))
        return sSymbolic;

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
        if ((aStartupInfo.dwFlags & STARTF_USESTDHANDLES) == STARTF_USESTDHANDLES
            && aStartupInfo.hStdInput != INVALID_HANDLE_VALUE && aStartupInfo.hStdInput != NULL
            && aStartupInfo.hStdOutput != INVALID_HANDLE_VALUE && aStartupInfo.hStdOutput != NULL
            && aStartupInfo.hStdError != INVALID_HANDLE_VALUE && aStartupInfo.hStdError != NULL)
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

        // Ignore errors from freopen_s()
        FILE* pStream;
        freopen_s(&pStream, "CON", "r", stdin);
        freopen_s(&pStream, "CON", "w", stdout);
        freopen_s(&pStream, "CON", "w", stderr);
        std::ios::sync_with_stdio(true);
    }
}

template <typename traits>
inline std::basic_ostream<char, traits>& operator<<(std::basic_ostream<char, traits>& stream,
                                                    const IID& rIid)
{
    LPOLESTR pRiid;
    if (StringFromIID(rIid, &pRiid) != S_OK)
        return stream << "?";

    // Special case well-known interfaces that pop up a lot, but which don't have their name in
    // the Registry.

    if (IsEqualIID(rIid, IID_IAgileObject))
        return stream << "IID_IAgileObject";
    if (IsEqualIID(rIid, IID_ICallFactory))
        return stream << "IID_ICallFactory";
    if (IsEqualIID(rIid, IID_IExternalConnection))
        return stream << "IID_IExternalConnection";
    if (IsEqualIID(rIid, IID_IFastRundown))
        return stream << "IID_IFastRundown";
    if (IsEqualIID(rIid, IID_IMarshal))
        return stream << "IID_IMarshal";
    if (IsEqualIID(rIid, IID_IMarshal2))
        return stream << "IID_IMarshal2";
    if (IsEqualIID(rIid, IID_INoMarshal))
        return stream << "IID_INoMarshal";
    if (IsEqualIID(rIid, IID_NULL))
        return stream << "IID_NULL";
    if (IsEqualIID(rIid, IID_IPersistPropertyBag))
        return stream << "IID_IPersistPropertyBag";
    if (IsEqualIID(rIid, IID_IPersistStreamInit))
        return stream << "IID_IPersistStreamInit";
    if (IsEqualIID(rIid, IID_IStdMarshalInfo))
        return stream << "IID_IStdMarshalInfo";

    DWORD nSize;
    if (RegGetValueW(HKEY_CLASSES_ROOT, std::wstring(L"Interface\\").append(pRiid).data(), NULL,
                     RRF_RT_REG_SZ, NULL, NULL, &nSize)
        == ERROR_SUCCESS)
    {
        std::vector<wchar_t> sValue(nSize / 2);
        if (RegGetValueW(HKEY_CLASSES_ROOT, std::wstring(L"Interface\\").append(pRiid).data(), NULL,
                         RRF_RT_REG_SZ, NULL, sValue.data(), &nSize)
            == ERROR_SUCCESS)
        {
            stream << "IID_" << convertUTF16ToUTF8(sValue.data());
        }
    }
    else if (RegGetValueW(HKEY_CLASSES_ROOT, std::wstring(L"CLSID\\").append(pRiid).data(), NULL,
                          RRF_RT_REG_SZ, NULL, NULL, &nSize)
             == ERROR_SUCCESS)
    {
        std::vector<wchar_t> sValue(nSize / 2);
        if (RegGetValueW(HKEY_CLASSES_ROOT, std::wstring(L"CLSID\\").append(pRiid).data(), NULL,
                         RRF_RT_REG_SZ, NULL, sValue.data(), &nSize)
            == ERROR_SUCCESS)
        {
            stream << convertUTF16ToUTF8(pRiid) << ":\"" << convertUTF16ToUTF8(sValue.data())
                   << "\"";
        }
    }
    else
    {
        stream << convertUTF16ToUTF8(pRiid);
    }

    CoTaskMemFree(pRiid);
    return stream;
}

inline bool isHighSurrogate(wchar_t c) { return (0xD800 <= c && c <= 0xDBFF); }

inline bool isLowSurrogate(wchar_t c) { return (0xDC00 <= c && c <= 0xDFFF); }

inline UINT surrogatePair(const wchar_t* pWchar)
{
    return (UINT)((((pWchar[0] - 0xD800) << 10) | (pWchar[1] - 0xDC00)) | 0x010000);
}

template <typename traits>
inline std::basic_ostream<char, traits>& outputCharString(const char* pChar, UINT nLength,
                                                          std::basic_ostream<char, traits>& rStream)
{
    if (pChar == nullptr)
    {
        rStream << "(null)";
        return rStream;
    }

    rStream << "\"";
    if (nLength > 100)
        nLength = 100;

    for (UINT i = 0; i < nLength; i++)
    {
        if (pChar[i] == '"' || pChar[i] == '\\')
            rStream << '\\' << pChar[i];
        else if (pChar[i] == '\n')
            rStream << "\\n";
        else if (pChar[i] == '\r')
            rStream << "\\r";
        else if (pChar[i] >= ' ' && pChar[i] <= '~')
            rStream << pChar[i];
        else
            rStream << "\\x" << to_uhex((UINT)pChar[i], 2);
    }

    rStream << "\"";

    return rStream;
}

template <typename traits>
inline std::basic_ostream<char, traits>&
outputWcharString(const wchar_t* pWchar, UINT nLength, std::basic_ostream<char, traits>& rStream)
{
    if (pWchar == nullptr)
    {
        rStream << "(null)";
        return rStream;
    }

    rStream << "\"";
    if (nLength > 100)
    {
        nLength = 100;
        if (isHighSurrogate(pWchar[nLength - 1]) && isLowSurrogate(pWchar[nLength]))
            nLength++;
    }

    for (UINT i = 0; i < nLength; i++)
    {
        if (isHighSurrogate(pWchar[i]) && i < nLength - 1 && isLowSurrogate(pWchar[i + 1]))
        {
            rStream << "\\u{" << to_uhex(surrogatePair(pWchar + i)) << "}";
            i++;
        }
        else if (pWchar[i] == '"' || pWchar[i] == '\\')
            rStream << '\\' << (char)pWchar[i];
        else if (pWchar[i] == '\n')
            rStream << "\\n";
        else if (pWchar[i] == '\r')
            rStream << "\\r";
        else if (pWchar[i] >= ' ' && pWchar[i] <= '~')
            rStream << (char)pWchar[i];
        else
            rStream << "\\u{" << to_uhex((UINT)pWchar[i]) << "}";
    }

    rStream << "\"";

    return rStream;
}

template <typename traits>
inline std::basic_ostream<char, traits>& operator<<(std::basic_ostream<char, traits>& stream,
                                                    const BSTR& rBstr)
{
    return outputWcharString(rBstr, SysStringLen(rBstr), stream);
}

template <typename traits>
inline std::basic_ostream<char, traits>& operator<<(std::basic_ostream<char, traits>& stream,
                                                    const VARIANT& rVariant)
{
    if (rVariant.vt & (VT_VECTOR | VT_ARRAY))
    {
        return stream << VARTYPE_to_string(rVariant.vt);
    }

    if (rVariant.vt & VT_BYREF)
        return stream << rVariant.byref;

    switch (rVariant.vt & VT_TYPEMASK)
    {
        case VT_I2:
            stream << rVariant.iVal;
            break;
        case VT_I4:
            stream << rVariant.lVal;
            break;
        case VT_R4:
            stream << rVariant.fltVal;
            break;
        case VT_R8:
            stream << rVariant.dblVal;
            break;
        case VT_CY:
            stream << rVariant.cyVal.int64;
            break;
        case VT_DATE:
            stream << (double)rVariant.date;
            break; // FIXME
        case VT_BSTR:
            stream << rVariant.bstrVal;
            break;
        case VT_DISPATCH:
            stream << rVariant.pdispVal;
            break;
        case VT_ERROR:
            stream << to_hex(rVariant.lVal, 8);
            break;
        case VT_BOOL:
            stream << (rVariant.boolVal ? "True" : "False");
            break;
        case VT_UNKNOWN:
            stream << rVariant.punkVal;
            break;
        case VT_DECIMAL:
            stream << to_uhex(rVariant.decVal.Hi32, 8) << to_ullhex(rVariant.decVal.Lo64, 16);
            break;
        case VT_I1:
            stream << (int)rVariant.bVal;
            break;
        case VT_UI1:
            stream << (unsigned int)rVariant.bVal;
            break;
        case VT_UI2:
            stream << (unsigned short)rVariant.iVal;
            break;
        case VT_UI4:
            stream << (unsigned int)rVariant.lVal;
            break;
        case VT_I8:
            stream << rVariant.llVal;
            break;
        case VT_UI8:
            stream << (unsigned long long)rVariant.llVal;
            break;
        case VT_INT:
            stream << rVariant.lVal;
            break;
        case VT_UINT:
            stream << (unsigned int)rVariant.lVal;
            break;
        case VT_HRESULT:
            stream << HRESULT_to_string(rVariant.lVal);
            break;
        case VT_PTR:
        case VT_CARRAY:
            stream << rVariant.byref;
            break;
        case VT_SAFEARRAY:
            stream << rVariant.parray;
            break;
        case VT_LPSTR:
            outputCharString(rVariant.pcVal, (UINT)std::strlen(rVariant.pcVal), stream);
            break;
        case VT_LPWSTR:
            outputWcharString((const wchar_t*)rVariant.byref,
                              (UINT)wcslen((const wchar_t*)rVariant.byref), stream);
            break;
        case VT_INT_PTR:
        case VT_UINT_PTR:
            stream << rVariant.plVal;
            break;
        default:
            stream << "?(" << (rVariant.vt & VT_TYPEMASK) << ")";
            break;
    }
    return stream;
}

inline bool isDirectlyPrintableType(VARTYPE nVt)
{
    switch (nVt)
    {
        case VT_I2:
        case VT_I4:
        case VT_R4:
        case VT_R8:
        case VT_BSTR:
        case VT_BOOL:
        case VT_I1:
        case VT_UI1:
        case VT_UI2:
        case VT_UI4:
        case VT_I8:
        case VT_UI8:
        case VT_INT:
        case VT_UINT:
        case VT_HRESULT:
        case VT_PTR:
        case VT_INT_PTR:
        case VT_UINT_PTR:
        case VT_CLSID:
            return true;

        default:
            return false;
    }
}

inline std::string withoutExtension(const std::string& rPathname)
{
    std::string::size_type nLastDot = rPathname.rfind('.');
    if (nLastDot == std::string::npos)
        return rPathname;
    return rPathname.substr(0, nLastDot);
}

inline std::string moduleName(HMODULE hModule)
{
    const DWORD NFILENAME = 1000;
    static wchar_t sFileName[NFILENAME];

    DWORD nSizeOut = GetModuleFileNameW(hModule, sFileName, NFILENAME);
    if (nSizeOut == 0 || nSizeOut == NFILENAME)
        return to_ullhex((uintptr_t)hModule, sizeof(void*) * 2);
    return convertUTF16ToUTF8(baseName(sFileName));
}

inline std::string prettyCodeAddress(void* pCode)
{
    HMODULE hModule;
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, (wchar_t*)pCode, &hModule))
        return to_ullhex((uintptr_t)pCode);

    return moduleName(hModule) + "!" + to_ullhex((uintptr_t)pCode, sizeof(void*) * 2);
}

class WaitForDebugger
{
    bool mbFlag = true;

public:
    WaitForDebugger(const std::string& sFromWhere)
    {
        std::cerr << "Note from " << sFromWhere << ": Attach process " << GetCurrentProcessId()
                  << " in a debugger NOW!\n"
                  << "Break the process, and set mbFlag to false, and continue (or look around)."
                  << std::endl;
        while (mbFlag)
            ;
    }
};

#define WIDE2(literal) L##literal
#define WIDE1(literal) WIDE2(literal)
#define WIDEFILE WIDE1(__FILE__)

#define WAIT_FOR_DEBUGGER_EACH_TIME                                                                \
    WaitForDebugger aWait(convertUTF16ToUTF8(baseName(WIDEFILE)) + std::string(":")                \
                          + std::to_string(__LINE__))
#define WAIT_FOR_DEBUGGER_ONCE static WAIT_FOR_DEBUGGER_EACH_TIME

#endif // INCLUDED_UTILS_HPP

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

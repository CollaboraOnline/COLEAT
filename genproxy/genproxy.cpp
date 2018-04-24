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
#include <codecvt>
#include <cstdlib>
#include <cwchar>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <map>
#include <set>
#include <sstream>
#include <string>

#include <Windows.h>

#include <comphelper/windowsdebugoutput.hxx>

#pragma warning(pop)

struct FuncTableEntry
{
    FUNCDESC* mpFuncDesc;
    std::vector<BSTR> mvNames;
};

struct Callback
{
    IID maIID;
    std::string msLibName;
    std::string msName;
};

struct DefaultInterface
{
    const std::string msLibName;
    const std::string msName;
    const IID maIID;
};

static std::string sOutputFolder = "generated";

static std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t> aUTF16ToUTF8;
static std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>, wchar_t> aUTF8ToUTF16;

static std::set<IID> aAlreadyHandledIIDs;
static std::set<std::string> aOnlyTheseInterfaces;

static std::set<Callback> aCallbacks;
static std::set<DefaultInterface> aDefaultInterfaces;

static bool bGenerateTracing = false;

static void Generate(const std::string& sLibName, ITypeInfo* pTypeInfo);

inline bool operator<(const Callback& a, const Callback& b)
{
    return ((a.msLibName < b.msLibName) || ((a.msLibName == b.msLibName) && (a.msName < b.msName)));
}

inline bool operator<(const DefaultInterface& a, const DefaultInterface& b)
{
    return ((a.msLibName < b.msLibName) || ((a.msLibName == b.msLibName) && (a.msName < b.msName)));
}

// FIXME: These general functions here in the beginning should really be somwhere else. I seem to
// end up copy-pasting them anyway.

inline bool operator<(const IID& a, const IID& b) { return std::memcmp(&a, &b, sizeof(a)) < 0; }

static std::string to_hex(uint64_t n, int w = 0)
{
    std::stringstream aStringStream;
    aStringStream << std::setfill('0') << std::setw(w) << std::hex << n;
    return aStringStream.str();
}

static std::string IID_initializer(const IID& aIID)
{
    std::string sResult;
    sResult = "{0x" + to_hex(aIID.Data1, 8) + ",0x" + to_hex(aIID.Data2, 4) + ",0x"
              + to_hex(aIID.Data3, 4);
    for (int i = 0; i < 8; ++i)
        sResult += ",0x" + to_hex(aIID.Data4[i], 2);
    sResult += "}";

    return sResult;
}

static std::string WindowsErrorString(DWORD nErrorCode)
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

    std::string sResult = aUTF16ToUTF8.to_bytes(pMsgBuf);
    if (sResult.length() > 2 && sResult.substr(sResult.length() - 2, 2) == "\r\n")
        sResult.resize(sResult.length() - 2);

    HeapFree(GetProcessHeap(), 0, pMsgBuf);

    return sResult;
}

static std::string WindowsErrorStringFromHRESULT(HRESULT nResult)
{
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

static bool IsIgnoredType(ITypeInfo* pTypeInfo)
{
    if (aOnlyTheseInterfaces.size() == 0)
        return false;

    HRESULT nResult;
    BSTR sName;
    nResult = pTypeInfo->GetDocumentation(MEMBERID_NIL, &sName, NULL, NULL, NULL);
    if (SUCCEEDED(nResult))
        return aOnlyTheseInterfaces.count(std::string(aUTF16ToUTF8.to_bytes(sName))) == 0;

    return false;
}

static bool IsIgnoredUserdefinedType(ITypeInfo* pTypeInfo, const TYPEDESC& aTypeDesc)
{
    if (aOnlyTheseInterfaces.size() == 0)
        return false;

    if (aTypeDesc.vt == VT_USERDEFINED)
    {
        HRESULT nResult;
        ITypeInfo* pReferencedTypeInfo;
        nResult = pTypeInfo->GetRefTypeInfo(aTypeDesc.hreftype, &pReferencedTypeInfo);

        if (SUCCEEDED(nResult))
            return IsIgnoredType(pReferencedTypeInfo);
    }
    return false;
}

static std::string VarTypeToString(VARTYPE n)
{
    std::string sResult;

    if (n & (VT_VECTOR | VT_ARRAY | VT_BYREF))
    {
        std::cerr << "Not yet implemented: VECTOR, ARRAY and BYREF\n";
        std::exit(1);
    }

    switch (n & ~(VT_VECTOR | VT_ARRAY | VT_BYREF))
    {
        case VT_I2:
            sResult = "int16_t";
            break;
        case VT_I4:
            sResult = "int32_t";
            break;
        case VT_R4:
            sResult = "float";
            break;
        case VT_R8:
            sResult = "double";
            break;
        case VT_DATE:
            sResult = "DATE";
            break;
        case VT_BSTR:
            sResult = "BSTR";
            break;
        case VT_DISPATCH:
            sResult = "IDispatch*";
            break;
        case VT_BOOL:
            sResult = "VARIANT_BOOL";
            break;
        case VT_VARIANT:
            sResult = "VARIANT";
            break;
        case VT_UNKNOWN:
            sResult = "/* UNKNOWN */ void*";
            break;
        case VT_I1:
            sResult = "int8_t";
            break;
        case VT_UI1:
            sResult = "uint8_t";
            break;
        case VT_UI2:
            sResult = "uint16_t";
            break;
        case VT_UI4:
            sResult = "uint32_t";
            break;
        case VT_I8:
            sResult = "int64_t";
            break;
        case VT_UI8:
            sResult = "uint64_t";
            break;
        case VT_INT:
            sResult = "int";
            break;
        case VT_UINT:
            sResult = "unsigned int";
            break;
        case VT_VOID:
            sResult = "void";
            break;
        case VT_HRESULT:
            sResult = "HRESULT";
            break;
        case VT_PTR:
            sResult = "void*";
            break;
        case VT_SAFEARRAY:
            sResult = "SAFEARRAY";
            break;
        case VT_USERDEFINED:
            sResult = "/* USERDEFINED */ void*";
            break;
        case VT_LPSTR:
            sResult = "LPSTR";
            break;
        case VT_LPWSTR:
            sResult = "LPWSTR";
            break;
        case VT_CLSID:
            sResult = "CLSID";
            break;
        default:
            std::cerr << "Unhandled VARTYPE " << (int)n << "\n";
            std::exit(1);
    }

    return sResult;
}

static std::string TypeToString(const std::string& sLibName, ITypeInfo* pTypeInfo,
                                const TYPEDESC& aTypeDesc, std::string& sReferencedName)
{
    HRESULT nResult;
    std::string sResult;

    sReferencedName = "";

    if (aTypeDesc.vt == VT_PTR)
    {
        if (IsIgnoredUserdefinedType(pTypeInfo, *aTypeDesc.lptdesc))
            sResult = "void*";
        else
            sResult = TypeToString(sLibName, pTypeInfo, *aTypeDesc.lptdesc, sReferencedName) + "*";
    }
    else if (aTypeDesc.vt == VT_USERDEFINED)
    {
        ITypeInfo* pReferencedTypeInfo;
        nResult = pTypeInfo->GetRefTypeInfo(aTypeDesc.hreftype, &pReferencedTypeInfo);

        BSTR sName = NULL;
        if (SUCCEEDED(nResult))
        {
            Generate(sLibName, pReferencedTypeInfo);
            nResult = pReferencedTypeInfo->GetDocumentation(MEMBERID_NIL, &sName, NULL, NULL, NULL);
        }

        TYPEATTR* pTypeAttr = NULL;
        if (SUCCEEDED(nResult))
            nResult = pReferencedTypeInfo->GetTypeAttr(&pTypeAttr);

        if (FAILED(nResult))
            sResult = "/* USERDEFINED:" + to_hex(aTypeDesc.hreftype) + " */ void*";
        else
        {
            if (pTypeAttr->typekind == TKIND_DISPATCH || pTypeAttr->typekind == TKIND_COCLASS)
            {
                sReferencedName = "C" + sLibName + "_" + aUTF16ToUTF8.to_bytes(sName);
                sResult = sReferencedName;
            }
            else if (pTypeAttr->typekind == TKIND_ENUM)
            {
                sReferencedName = "E" + sLibName + "_" + aUTF16ToUTF8.to_bytes(sName);
                sResult = sReferencedName;
            }
            else if (pTypeAttr->typekind == TKIND_ALIAS)
            {
                sResult = VarTypeToString(pTypeAttr->tdescAlias.vt);
            }
            else
            {
                std::cerr << "Huh, unhandled typekind " << pTypeAttr->typekind << "\n";
                std::exit(1);
            }
            pTypeInfo->ReleaseTypeAttr(pTypeAttr);
        }
    }
    else
    {
        sResult = VarTypeToString(aTypeDesc.vt);
    }

    return sResult;
}

static bool IsIntegralType(VARTYPE n)
{
    switch (n)
    {
        case VT_I2:
        case VT_I4:
        case VT_I1:
        case VT_UI1:
        case VT_UI2:
        case VT_UI4:
        case VT_I8:
        case VT_UI8:
        case VT_INT:
        case VT_UINT:
            return true;
        default:
            return false;
    }
}

static int64_t GetIntegralValue(VARIANT& a)
{
    switch (a.vt)
    {
        case VT_I2:
            return a.iVal;
        case VT_I4:
            return a.lVal;
        case VT_I1:
            return a.bVal;
        case VT_UI1:
            return (unsigned int)a.bVal;
        case VT_UI2:
            return (unsigned int)a.iVal;
        case VT_UI4:
            return (unsigned int)a.lVal;
        case VT_I8:
            return a.llVal;
        case VT_UI8:
            return a.llVal;
        case VT_INT:
            return a.lVal;
        case VT_UINT:
            return (unsigned int)a.lVal;
        default:
            std::abort();
    }
}

static void GenerateSink(const std::string& sLibName, ITypeInfo* const pTypeInfo)
{
    HRESULT nResult;

    if (IsIgnoredType(pTypeInfo))
        return;

    BSTR sName;
    nResult = pTypeInfo->GetDocumentation(MEMBERID_NIL, &sName, NULL, NULL, NULL);
    if (FAILED(nResult))
    {
        std::cerr << "GetDocumentation failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
        std::exit(1);
    }

    std::cerr << sLibName << "_" << aUTF16ToUTF8.to_bytes(sName) << " (sink)\n";

    std::string sClass = sLibName + "_" + aUTF16ToUTF8.to_bytes(sName);

    TYPEATTR* pTypeAttr;
    nResult = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(nResult))
    {
        std::cerr << "GetTypeAttr failed for " << sClass << ": "
                  << WindowsErrorStringFromHRESULT(nResult) << "\n";
        std::exit(1);
    }

    if (pTypeAttr->typekind != TKIND_DISPATCH)
    {
        std::cerr << "-" << sClass << " (is not TKIND_DISPATCH)\n";
        return;
    }

    if (aAlreadyHandledIIDs.count(pTypeAttr->guid))
        return;
    aAlreadyHandledIIDs.insert(pTypeAttr->guid);

    aCallbacks.insert({ pTypeAttr->guid, sLibName, aUTF16ToUTF8.to_bytes(sName) });

    const std::string sHeader = sOutputFolder + "/" + sClass + ".hxx";

    std::ofstream aHeader(sHeader);
    if (!aHeader.good())
    {
        std::cerr << "Could not open '" << sHeader << "' for writing\n";
        std::exit(1);
    }

    const std::string sCode = sOutputFolder + "/" + sClass + ".cxx";
    std::ofstream aCode(sCode);
    if (!aCode.good())
    {
        std::cerr << "Could not open '" << sCode << "' for writing\n";
        std::exit(1);
    }

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_C" << sClass << "_HXX\n";
    aHeader << "#define INCLUDED_C" << sClass << "_HXX\n";
    aHeader << "\n";
    aHeader << "#pragma warning(push)\n";
    aHeader << "#pragma warning(disable: 4668 4820 4917)\n";
    aHeader << "#include <Windows.h>\n";
    aHeader << "#pragma warning(pop)\n";
    aHeader << "\n";

    aCode << "// Generated file. Do not edit.\n";
    aCode << "\n";
    aCode << "#include <cstdlib>\n";
    aCode << "#include <iostream>\n";
    aCode << "\n";
    aCode << "#include \"" << sClass << ".hxx\"\n";
    aCode << "\n";

    aCode << "HRESULT " << sClass << "CallbackInvoke(IDispatch* pDispatchToProxy,\n";
    aCode << "            DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,\n";
    aCode << "            DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, "
             "UINT* puArgErr)\n";
    aCode << "{\n";

    aCode << "    HRESULT nResult;\n";
    aCode << "    DISPPARAMS aLocalDispParams = *pDispParams;\n";
    aCode << "    aLocalDispParams.rgvarg = new VARIANTARG[aLocalDispParams.cArgs];\n";

    aCode << "    switch (dispIdMember)\n";
    aCode << "    {\n";

    std::set<std::string> aIncludedHeaders;

    for (UINT nFunc = 0; nFunc < pTypeAttr->cFuncs; ++nFunc)
    {
        FUNCDESC* pFuncDesc;
        nResult = pTypeInfo->GetFuncDesc(nFunc, &pFuncDesc);
        if (FAILED(nResult))
        {
            std::cerr << "GetFuncDesc(" << nFunc
                      << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            std::exit(1);
        }

        if (pFuncDesc->memid & 0x60000000)
        {
            pTypeInfo->ReleaseFuncDesc(pFuncDesc);
            continue;
        }

        BSTR sFuncName;
        nResult = pTypeInfo->GetDocumentation(pFuncDesc->memid, &sFuncName, NULL, NULL, NULL);
        if (FAILED(nResult))
        {
            std::cerr << "GetDocumentation(" << pFuncDesc->memid
                      << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            std::exit(1);
        }

        aCode << "        case " << pFuncDesc->memid << ": /* " << aUTF16ToUTF8.to_bytes(sFuncName)
              << " */\n";
        aCode << "            {\n";

        // Go through parameter list declared for the function in the type library, manipulate
        // corresponding actual run-time parameters. Currently the only special thing needed is to
        // proxy VT_PTR parameters to VT_USERDEFINED interfaces as a pointer the corresponding
        // genproxy-generated class.

        // Note that the parameters in the DISPPARAMS::rgvarg are in right-to-left order, thus all
        // the fuss with '(pFuncDesc->cParams-nParam-1)' below.

        // FIXME: We don't take into account the possibility of optional parameters being left out.
        // But does the outgoing interfaces even have optional parameters? I don't think so.
        for (int nParam = 0; nParam < pFuncDesc->cParams; ++nParam)
        {
            aCode << "                if (" << (pFuncDesc->cParams - nParam - 1)
                  << " < pDispParams->cArgs)\n";
            aCode << "                {\n";
            aCode << "                    aLocalDispParams.rgvarg["
                  << (pFuncDesc->cParams - nParam - 1) << "] = pDispParams->rgvarg["
                  << (pFuncDesc->cParams - nParam - 1) << "];\n";

            // Generate the thing this stuff is needed for
            if (pFuncDesc->lprgelemdescParam[nParam].tdesc.vt == VT_PTR
                && pFuncDesc->lprgelemdescParam[nParam].tdesc.lptdesc->vt == VT_USERDEFINED
                && !IsIgnoredUserdefinedType(pTypeInfo,
                                             *pFuncDesc->lprgelemdescParam[nParam].tdesc.lptdesc))
            {
                ITypeInfo* pReferencedTypeInfo;
                nResult = pTypeInfo->GetRefTypeInfo(
                    pFuncDesc->lprgelemdescParam[nParam].tdesc.lptdesc->hreftype,
                    &pReferencedTypeInfo);
                if (FAILED(nResult))
                {
                    std::cerr << "GetRefTypeInfo failed\n";
                    std::exit(1);
                }
                TYPEATTR* pReferencedTypeAttr;
                nResult = pReferencedTypeInfo->GetTypeAttr(&pReferencedTypeAttr);
                if (FAILED(nResult))
                {
                    std::cerr << "GetTypeAttr failed\n";
                    std::exit(1);
                }
                if (pReferencedTypeAttr->typekind == TKIND_DISPATCH
                    || pReferencedTypeAttr->typekind == TKIND_COCLASS)
                {
                    BSTR sReferencedTypeNameBstr;
                    nResult = pReferencedTypeInfo->GetDocumentation(
                        MEMBERID_NIL, &sReferencedTypeNameBstr, NULL, NULL, NULL);
                    if (FAILED(nResult))
                    {
                        std::cerr << "GetTypeAttr failed\n";
                        std::exit(1);
                    }
                    std::string sReferencedTypeName(aUTF16ToUTF8.to_bytes(sReferencedTypeNameBstr));
                    aCode << "                    aLocalDispParams.rgvarg["
                          << (pFuncDesc->cParams - nParam - 1) << "].pdispVal = new C" << sLibName
                          << "_" << sReferencedTypeName << "(aLocalDispParams.rgvarg["
                          << (pFuncDesc->cParams - nParam - 1) << "].pdispVal);\n";

                    if (!aIncludedHeaders.count(sReferencedTypeName))
                    {
                        aHeader << "#include \"C" << sLibName << "_" << sReferencedTypeName
                                << ".hxx\"\n";
                        aIncludedHeaders.insert(sReferencedTypeName);
                    }
                }
            }
            aCode << "                }\n";
        }

        aCode << "            };\n";
        aCode << "            break;\n";

        pTypeInfo->ReleaseFuncDesc(pFuncDesc);
    }
    aCode << "        default:\n";
    aCode << "            std::cerr << \"Unhandled DISPID \" << dispIdMember << \" in " << sClass
          << "CallbackInvoke!\" << std::endl;\n";
    aCode << "            std::abort();\n";

    aCode << "    }\n";

    aCode << "    nResult = pDispatchToProxy->Invoke(dispIdMember, riid, lcid, wFlags,\n";
    aCode << "        &aLocalDispParams, pVarResult, pExcepInfo, puArgErr);\n";

    aCode << "    delete[] aLocalDispParams.rgvarg;\n";
    aCode << "\n";
    aCode << "    return nResult;\n";
    aCode << "}\n";

    if (aIncludedHeaders.size() > 0)
        aHeader << "\n";
    aHeader << "HRESULT " << sClass << "CallbackInvoke(IDispatch* pDispatchToProxy,\n";
    aHeader << "            DISPID dispIdMember, REFIID riid,LCID lcid, WORD wFlags,\n";
    aHeader << "            DISPPARAMS* pDispParams, VARIANT* pVarResult,\n";
    aHeader << "            EXCEPINFO* pExcepInfo, UINT* puArgErr);\n";
    aHeader << "\n";
    aHeader << "#endif // INCLUDED_C" << sClass << "_HXX\n";

    aHeader.close();
    if (!aHeader.good())
    {
        std::cerr << "Problems writing to '" << sHeader << "'\n";
        std::exit(1);
    }

    aCode.close();
    if (!aCode.good())
    {
        std::cerr << "Problems writing to '" << sCode << "'\n";
        std::exit(1);
    }
}

static void GenerateEnum(const std::string& sLibName, const std::string& sTypeName,
                         ITypeInfo* const pTypeInfo, const TYPEATTR* const pTypeAttr)
{
    HRESULT nResult;

    std::cout << sLibName << "." << sTypeName << " (enum)\n";

    if (pTypeAttr->cFuncs != 0)
    {
        std::cerr << "Huh, enumeration has functions?\n";
        return;
    }

    if (pTypeAttr->cVars == 0)
    {
        std::cerr << "Huh, enumeration has no values?\n";
        return;
    }

    const std::string sClass = sLibName + "_" + sTypeName;

    const std::string sHeader = sOutputFolder + "/E" + sClass + ".hxx";
    std::ofstream aHeader(sHeader);
    if (!aHeader.good())
    {
        std::cerr << "Could not open '" << sHeader << "' for writing\n";
        std::exit(1);
    }

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_E" << sClass << "_HXX\n";
    aHeader << "#define INCLUDED_E" << sClass << "_HXX\n";
    aHeader << "\n";
    aHeader << "enum E" << sClass << " {\n";

    for (UINT i = 0; i < pTypeAttr->cVars; ++i)
    {
        VARDESC* pVarDesc;
        nResult = pTypeInfo->GetVarDesc(i, &pVarDesc);
        if (FAILED(nResult))
        {
            std::cerr << "GetVarDesc(" << i << ") failed:" << WindowsErrorStringFromHRESULT(nResult)
                      << "\n";
            continue;
        }

        if (pVarDesc->varkind != VAR_CONST)
        {
            std::cerr << "Huh, enum has non-const member?\n";
            continue;
        }

        if (!IsIntegralType(pVarDesc->lpvarValue->vt))
        {
            std::cerr << "Huh, enum has non-integral member?\n";
            continue;
        }

        BSTR sVarName;
        UINT nNames;
        nResult = pTypeInfo->GetNames(pVarDesc->memid, &sVarName, 1, &nNames);
        if (FAILED(nResult))
        {
            std::cerr << "GetNames(" << pVarDesc->memid
                      << ") failed:" << WindowsErrorStringFromHRESULT(nResult) << "\n";
            continue;
        }
        aHeader << "    " << aUTF16ToUTF8.to_bytes(sVarName) << " = "
                << GetIntegralValue(*pVarDesc->lpvarValue) << ",\n";
    }

    aHeader << "}\n";
    aHeader << "\n";
    aHeader << "#endif // INCLUDED_E" << sClass << "_HXX\n";
}

static void GenerateCoclass(const std::string& sLibName, const std::string& sTypeName,
                            ITypeInfo* const pTypeInfo, const TYPEATTR* const pTypeAttr)
{
    HRESULT nResult;

    std::cout << sLibName << "." << sTypeName << " (coclass)\n";

    std::string sDefaultInterface;

    for (UINT i = 0; i < pTypeAttr->cImplTypes; ++i)
    {
        HREFTYPE nRefType;
        nResult = pTypeInfo->GetRefTypeOfImplType(i, &nRefType);
        if (FAILED(nRefType))
        {
            std::cerr << "GetRefTypeOfImplType(" << i
                      << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            continue;
        }

        ITypeInfo* pImplTypeInfo;
        nResult = pTypeInfo->GetRefTypeInfo(nRefType, &pImplTypeInfo);
        if (FAILED(nRefType))
        {
            std::cerr << "GetRefTypeInfo(" << nRefType
                      << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            continue;
        }

        INT nImplTypeFlags;
        nResult = pTypeInfo->GetImplTypeFlags(i, &nImplTypeFlags);
        if (FAILED(nRefType))
        {
            std::cerr << "GetImplTypeFlags(" << i
                      << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            continue;
        }

        BSTR sImplTypeName;
        nResult = pImplTypeInfo->GetDocumentation(MEMBERID_NIL, &sImplTypeName, NULL, NULL, NULL);
        if (FAILED(nResult))
        {
            std::cerr << "GetDocumentation failed: " << WindowsErrorStringFromHRESULT(nResult)
                      << "\n";
            std::exit(1);
        }

        TYPEATTR* pImplTypeAttr;
        nResult = pImplTypeInfo->GetTypeAttr(&pImplTypeAttr);
        if (FAILED(nResult))
        {
            std::cerr << "GetTypeAttr failed\n";
            std::exit(1);
        }

        if (nImplTypeFlags & IMPLTYPEFLAG_FSOURCE)
            GenerateSink(sLibName, pImplTypeInfo);
        else
        {
            Generate(sLibName, pImplTypeInfo);
            if (nImplTypeFlags & IMPLTYPEFLAG_FDEFAULT)
            {
                if (!sDefaultInterface.empty())
                {
                    std::cerr << "Multiple default interfaces for '" << sTypeName << "' in '"
                              << sLibName << "'\n";
                    std::exit(1);
                }
                sDefaultInterface = aUTF16ToUTF8.to_bytes(sImplTypeName);
                aDefaultInterfaces.insert({ sLibName, sDefaultInterface, pImplTypeAttr->guid });
            }
        }
    }

    if (sDefaultInterface.empty())
    {
        std::cerr << "No default interface for '" << sTypeName << "' in '" << sLibName << "\n";
        std::exit(1);
    }

    const std::string sClass = sLibName + "_" + sTypeName;

    const std::string sHeader = sOutputFolder + "/C" + sClass + ".hxx";
    std::ofstream aHeader(sHeader);
    if (!aHeader.good())
    {
        std::cerr << "Could not open '" << sHeader << "' for writing\n";
        std::exit(1);
    }

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_C" << sClass << "_HXX\n";
    aHeader << "#define INCLUDED_C" << sClass << "_HXX\n";
    aHeader << "\n";

    aHeader << "#include \"C" << sLibName << "_" << sDefaultInterface << ".hxx\"\n";
    aHeader << "\n";
    aHeader << "class C" << sClass << ": public C" << sLibName << "_" << sDefaultInterface << "\n";
    aHeader << "{\n";
    aHeader << "public:\n";
    aHeader << "    C" << sClass << "(IDispatch* pDispatchToProxy) :\n";
    aHeader << "        C" << sLibName << "_" << sDefaultInterface << "(pDispatchToProxy, "
            << IID_initializer(pTypeAttr->guid) << ")\n";
    aHeader << "    {\n";
    aHeader << "    }\n";
    aHeader << "};\n";

    aHeader << "\n";
    aHeader << "#endif // INCLUDED_C" << sClass << "_HXX\n";
    aHeader.close();

    if (!aHeader.good())
    {
        std::cerr << "Problems writing to '" << sHeader << "'\n";
        std::exit(1);
    }
}

static void CollectFuncInfo(ITypeInfo* const pTypeInfo, const TYPEATTR* pTypeAttr,
                            std::vector<FuncTableEntry>& rVtable)
{
    HRESULT nResult;

    for (UINT nFunc = 0; nFunc < pTypeAttr->cFuncs; ++nFunc)
    {
        nResult = pTypeInfo->GetFuncDesc(nFunc, &rVtable[nFunc].mpFuncDesc);
        if (FAILED(nResult))
        {
            std::cerr << "GetFuncDesc(" << nFunc
                      << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            std::exit(1);
        }
        if (nFunc > 0 && rVtable[nFunc].mpFuncDesc->oVft > 0
            && rVtable[nFunc].mpFuncDesc->oVft
                   != rVtable[nFunc - 1].mpFuncDesc->oVft + (SHORT)sizeof(void*))
        {
            std::cerr << "Huh, functions not in vtable order?\n";
            std::exit(1);
        }
        rVtable[nFunc].mvNames.resize(1u + rVtable[nFunc].mpFuncDesc->cParams);
        UINT nNames;
        nResult
            = pTypeInfo->GetNames(rVtable[nFunc].mpFuncDesc->memid, rVtable[nFunc].mvNames.data(),
                                  1u + rVtable[nFunc].mpFuncDesc->cParams, &nNames);
        if (FAILED(nResult))
        {
            std::cerr << "GetNames(" << nFunc << "," << 1 + rVtable[nFunc].mpFuncDesc->cParams
                      << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            std::exit(1);
        }
        rVtable[nFunc].mvNames.resize(nNames);
    }
}

static void GenerateDispatch(const std::string& sLibName, const std::string& sTypeName,
                             ITypeInfo* const pDispTypeInfo, const TYPEATTR* const pDispTypeAttr)
{
    HRESULT nResult;

    std::cout << sLibName << "." << sTypeName << " (dispatch)\n";

    if (!(pDispTypeAttr->wTypeFlags & TYPEFLAG_FDUAL))
    {
        std::cerr << "Huh, interface is not dual?\n";
        std::exit(1);
    }

    HREFTYPE nVtblRefType;
    nResult = pDispTypeInfo->GetRefTypeOfImplType((UINT)-1, &nVtblRefType);
    if (FAILED(nResult))
    {
        std::cerr << "GetRefTypeOfImplType(-1) failed: " << WindowsErrorStringFromHRESULT(nResult)
                  << "\n";
        std::exit(1);
    }

    ITypeInfo* pVtblTypeInfo;
    nResult = pDispTypeInfo->GetRefTypeInfo(nVtblRefType, &pVtblTypeInfo);
    if (FAILED(nResult))
    {
        std::cerr << "GetRefTypeInfo(" << nVtblRefType
                  << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
        std::exit(1);
    }

    TYPEATTR* pVtblTypeAttr;
    nResult = pVtblTypeInfo->GetTypeAttr(&pVtblTypeAttr);
    if (FAILED(nResult))
    {
        std::cerr << "GetTypeAttr failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
        std::exit(1);
    }

    std::vector<FuncTableEntry> vDispFuncTable(pDispTypeAttr->cFuncs);
    CollectFuncInfo(pDispTypeInfo, pDispTypeAttr, vDispFuncTable);

    std::vector<FuncTableEntry> vVtblFuncTable(pVtblTypeAttr->cFuncs);
    CollectFuncInfo(pVtblTypeInfo, pVtblTypeAttr, vVtblFuncTable);

    if (pDispTypeAttr->cFuncs != pVtblTypeAttr->cFuncs + 7)
    {
        std::cerr << "Huh, IDispatch-based interface has different number of functions than the "
                     "Vtbl-based?\n";
        std::exit(1);
    }

    // Sanity check: Are the IDispatch- and vtbl-based functions similar enough?

    for (UINT nVtblFunc = 0; nVtblFunc < pVtblTypeAttr->cFuncs; ++nVtblFunc)
    {
        const UINT nDispFunc = nVtblFunc + 7;

        if (std::wcscmp(vDispFuncTable[nDispFunc].mvNames[0], vVtblFuncTable[nVtblFunc].mvNames[0])
            != 0)
        {
            std::cerr
                << "Huh, IDispatch-based method has different name than the Vtbl-based one?\n";
            std::cerr << "(" << aUTF16ToUTF8.to_bytes(vDispFuncTable[nDispFunc].mvNames[0])
                      << " vs. " << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nVtblFunc].mvNames[0])
                      << ")\n";
        }

        // The vtbl-based function can have the same number of parameters as the IDispatch-based
        // one, or one more (for the return value or LCID), or two more (for both return value and
        // LCID).
        if (vDispFuncTable[nDispFunc].mpFuncDesc->cParams
                != vVtblFuncTable[nVtblFunc].mpFuncDesc->cParams
            && vDispFuncTable[nDispFunc].mpFuncDesc->cParams
                   != vVtblFuncTable[nVtblFunc].mpFuncDesc->cParams - 1
            && vDispFuncTable[nDispFunc].mpFuncDesc->cParams
                   != vVtblFuncTable[nVtblFunc].mpFuncDesc->cParams - 2)
        {
            std::cerr << "Huh, IDispatch-based method has different number of parameters than the "
                         "Vtbl-based one?\n";
            std::cerr << "(Function: "
                      << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nVtblFunc].mvNames[0]) << ", "
                      << vDispFuncTable[nDispFunc].mpFuncDesc->cParams << " vs. "
                      << vVtblFuncTable[nVtblFunc].mpFuncDesc->cParams << ")\n";
            std::exit(1);
        }

        for (int nVtblParam = 1; nVtblParam < vVtblFuncTable[nVtblFunc].mpFuncDesc->cParams;
             ++nVtblParam)
        {
            if (nVtblParam + 1 >= vDispFuncTable[nDispFunc].mpFuncDesc->cParams)
                break;
            if (std::wcscmp(vDispFuncTable[nDispFunc].mvNames[nVtblParam + 1u],
                            vVtblFuncTable[nVtblFunc].mvNames[nVtblParam + 1u])
                != 0)
            {
                std::cerr << "Huh, IDispatch-based method parameter has different name than that "
                             "in the Vtbl-based one?\n";
                std::cerr
                    << "("
                    << aUTF16ToUTF8.to_bytes(vDispFuncTable[nDispFunc].mvNames[nVtblParam + 1u])
                    << " vs. "
                    << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nVtblFunc].mvNames[nVtblParam + 1u])
                    << ")\n";
                std::exit(1);
            }

            if (vDispFuncTable[nDispFunc].mpFuncDesc->lprgelemdescParam[nVtblParam].tdesc.vt
                != vVtblFuncTable[nVtblFunc].mpFuncDesc->lprgelemdescParam[nVtblParam].tdesc.vt)
            {
                std::cerr << "Huh, IDispatch-based parameter " << nVtblParam << " of "
                          << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nVtblFunc].mvNames[0])
                          << " has different type than that in the Vtbl-based one?\n";
                std::cerr
                    << "("
                    << vDispFuncTable[nDispFunc].mpFuncDesc->lprgelemdescParam[nVtblParam].tdesc.vt
                    << " vs. "
                    << vVtblFuncTable[nVtblFunc].mpFuncDesc->lprgelemdescParam[nVtblParam].tdesc.vt
                    << ")\n";
                std::exit(1);
            }
        }
    }

    const std::string sClass = sLibName + "_" + sTypeName;

    // Open output files

    const std::string sHeader = sOutputFolder + "/C" + sClass + ".hxx";

    std::ofstream aHeader(sHeader);
    if (!aHeader.good())
    {
        std::cerr << "Could not open '" << sHeader << "' for writing\n";
        std::exit(1);
    }

    const std::string sCode = sOutputFolder + "/C" + sClass + ".cxx";
    std::ofstream aCode(sCode);
    if (!aCode.good())
    {
        std::cerr << "Could not open '" << sCode << "' for writing\n";
        std::exit(1);
    }

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_C" << sClass << "_HXX\n";
    aHeader << "#define INCLUDED_C" << sClass << "_HXX\n";
    aHeader << "\n";
    aHeader << "#pragma warning(push)\n";
    aHeader << "#pragma warning(disable: 4668 4820 4917)\n";
    aHeader << "#include <Windows.h>\n";
    aHeader << "#pragma warning(pop)\n";
    aHeader << "\n";
    aHeader << "#include \"CProxiedDispatch.hpp\"\n";
    aHeader << "\n";

    aCode << "// Generated file. Do not edit\n";
    aCode << "\n";
    aCode << "#include \"C" << sClass << ".hxx\"\n";
    aCode << "\n";
    if (bGenerateTracing)
        aCode << "#include <iostream>\n";
    aCode << "#include <vector>\n";
    aCode << "\n";

    // Then the interesting bits. We loop over the functions twice, firt outputting to the code
    // file, then to the header, so that we can add #include statements as needed to the header file
    // while noticing their need while generating the code file.

    // Note that it is the vtbl-based interface we want to output proxy code for. It is after all
    // handling calls to that which is the very reason for this thing.

    std::set<std::string> aIncludedHeaders;
    std::string sReferencedName;

    // Generate cxx file

    for (UINT nFunc = 0; nFunc < pVtblTypeAttr->cFuncs; ++nFunc)
    {
        // Generate the code for one method. They are all of type HRESULT, no?

        aCode << "// vtbl entry " << (vVtblFuncTable[nFunc].mpFuncDesc->oVft / sizeof(void*))
              << "\n";

        aCode << "HRESULT ";
        switch (vVtblFuncTable[nFunc].mpFuncDesc->callconv)
        {
            case CC_CDECL:
                aCode << "__cdecl";
                break;
            case CC_STDCALL:
                aCode << "__stdcall";
                break;
            default:
                std::cerr << "Huh, unhandled call convention "
                          << vVtblFuncTable[nFunc].mpFuncDesc->callconv << "?\n";
                std::exit(1);
        }
        aCode << " C" << sClass << "::";

        std::string sFuncName;
        if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYGET)
            sFuncName = "get";
        else if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYPUT)
            sFuncName = "put";
        else if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYPUTREF)
            sFuncName = "putref";
        sFuncName += aUTF16ToUTF8.to_bytes(vVtblFuncTable[nFunc].mvNames[0]);
        aCode << sFuncName << "(";

        // Parameter list
        for (int nParam = 0; nParam < vVtblFuncTable[nFunc].mpFuncDesc->cParams; ++nParam)
        {
            // FIXME: If a parameter is an enum type that we don't bother generating a C++ enum for,
            // we should just use an integral type for it, not void*. Need to check the type
            // description for the underlying integral type, though.
            if (IsIgnoredUserdefinedType(
                    pVtblTypeInfo,
                    vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc))
                aCode << "void* ";
            else
            {
                aCode << TypeToString(
                             sLibName, pVtblTypeInfo,
                             vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc,
                             sReferencedName)
                      << " ";

                if (sReferencedName != "" && !aIncludedHeaders.count(sReferencedName))
                {
                    // Also output a forward declaration in case of circular dependnecy between headers
                    aHeader << "class " << sReferencedName << ";\n";
                    // aHeader << "#include \"" << sReferencedName << ".hxx\"\n";
                    aIncludedHeaders.insert(sReferencedName);
                }
            }
            if ((size_t)(nParam + 1) < vVtblFuncTable[nFunc].mvNames.size())
                aCode << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nFunc].mvNames[nParam + 1u]);
            else
                aCode << "p" << nParam;
            if (nParam + 1 < vVtblFuncTable[nFunc].mpFuncDesc->cParams)
                aCode << ", ";
        }
        aCode << ")\n";

        // Code block of the function: Package parameters into an array of VARIANTs, and call our
        // magic CProxiedDispatch::Invoke().

        aCode << "{\n";

        if (bGenerateTracing)
            aCode << "    std::wcout << \"C" << sClass << "::" << sFuncName << "(\";\n";

        aCode << "    std::vector<VARIANT> vParams(" << vVtblFuncTable[nFunc].mpFuncDesc->cParams
              << ");\n";
        if (vVtblFuncTable[nFunc].mpFuncDesc->cParams > 0)
        {
            aCode << "    UINT nActualParams = 0;\n";
            aCode << "    bool bGotAll = false;\n";
            aCode << "    (void) bGotAll;\n";
        }

        bool bHadOptional = false;
        bool bHadSomeJustDeclaredAsOptional = false;
        std::string sRetvalName = "nullptr";
        int nRetvalParam = -1;
        for (int nParam = 0; nParam < vVtblFuncTable[nFunc].mpFuncDesc->cParams; ++nParam)
        {
            std::string sParamName;
            if ((size_t)(nParam + 1) < vVtblFuncTable[nFunc].mvNames.size())
                sParamName = aUTF16ToUTF8.to_bytes(vVtblFuncTable[nFunc].mvNames[nParam + 1u]);
            else
                sParamName = "p" + std::to_string(nParam);

            if (bHadOptional
                && !(vVtblFuncTable[nFunc]
                         .mpFuncDesc->lprgelemdescParam[nParam]
                         .paramdesc.wParamFlags
                     & (PARAMFLAG_FOPT | PARAMFLAG_FRETVAL | PARAMFLAG_FLCID)))
            {
                std::cerr << "Huh, Optional parameter "
                          << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nFunc].mvNames[(UINT)nParam])
                          << " to function "
                          << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nFunc].mvNames[0])
                          << " followed by non-optional "
                          << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nFunc].mvNames[nParam + 1u])
                          << "?\n";
                std::exit(1);
            }

            // FIXME: Assume that a VARIANT parameter marked as optional actually isn't optional in
            // the sense that it can be left out completely, but instead the optionality means that
            // it might be "empty" (VT_EMPTY, or perhaps VT_ERROR with an scode of
            // DISP_E_PARAMNOTFOUND, like for VT_PTR). Such optional VT_VARIANT parameters can thus
            // be followed by non-optional ones. Check this with experimentation.
            if ((vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].paramdesc.wParamFlags
                 & PARAMFLAG_FOPT)
                && vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc.vt
                       != VT_VARIANT)
                bHadOptional = true;

            if (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].paramdesc.wParamFlags
                & PARAMFLAG_FOPT)
                bHadSomeJustDeclaredAsOptional = true;

            if (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].paramdesc.wParamFlags
                & PARAMFLAG_FRETVAL)
            {
                if (sRetvalName != "nullptr")
                {
                    std::cerr << "Several return values?\n";
                    std::exit(1);
                }
                sRetvalName = sParamName;
            }

            if (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].paramdesc.wParamFlags
                & PARAMFLAG_FRETVAL)
            {
                nRetvalParam = nParam;
                continue;
            }

            // Ignore parameters marked as LCID. Those are not present in the IDispatch-based
            // function, which is what we proxy to.
            if (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].paramdesc.wParamFlags
                & PARAMFLAG_FLCID)
            {
                aCode << "    (void) "
                      << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nFunc].mvNames[nParam + 1u]) << ";\n";

                if (bGenerateTracing)
                {
                    aCode << "    if(!bGotAll)\n";
                    aCode << "        std::wcout";
                    if (nParam > 0)
                        aCode << " << \",\"";
                    aCode << " << " << sParamName << ";\n";
                }

                continue;
            }

            if (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc.vt == VT_PTR
                && vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc.lptdesc->vt
                       == VT_VARIANT)
            {
                // If this is an optional parameter, check whether we got it
                if (bHadOptional)
                {
                    aCode << "    if (" << sParamName << "->vt == VT_ERROR && " << sParamName
                          << "->scode == DISP_E_PARAMNOTFOUND)\n";
                    aCode << "        bGotAll = true;\n";
                }
            }

            aCode << "    if (!bGotAll)\n";
            aCode << "    {\n";
            switch (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc.vt)
            {
                case VT_I2:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_I2;\n";
                    aCode << "        vParams[nActualParams].iVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_I4:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_I4;\n";
                    aCode << "        vParams[nActualParams].lVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_R4:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_R4;\n";
                    aCode << "        vParams[nActualParams].fltVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_R8:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_R8;\n";
                    aCode << "        vParams[nActualParams].dblVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_BSTR:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_BSTR;\n";
                    aCode << "        vParams[nActualParams].bstrVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_DISPATCH:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_DISPATCH;\n";
                    aCode << "        vParams[nActualParams].pdispVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_BOOL:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_BOOL;\n";
                    aCode << "        vParams[nActualParams].boolVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_VARIANT:
                    // FIXME: Is this correct? Experimentation will show.
                    aCode << "        vParams[nActualParams] = " << sParamName << ";\n";
                    aCode << "        if (vParams[nActualParams].vt == VT_ILLEGAL\n";
                    aCode << "            || (vParams[nActualParams].vt == VT_ERROR\n";
                    aCode << "                && vParams[nActualParams].scode == "
                             "DISP_E_PARAMNOTFOUND))\n";
                    aCode << "            vParams[nActualParams].vt = VT_EMPTY;\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_UI2:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_UI2;\n";
                    aCode << "        vParams[nActualParams].iVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_UI4:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_UI4;\n";
                    aCode << "        vParams[nActualParams].lVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_I8:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_I8;\n";
                    aCode << "        vParams[nActualParams].llVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_UI8:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_UI8;\n";
                    aCode << "        vParams[nActualParams].llVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_INT:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_INT;\n";
                    aCode << "        vParams[nActualParams].lVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_UINT:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_UINT;\n";
                    aCode << "        vParams[nActualParams].lVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_INT_PTR:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_INT_PTR;\n";
                    aCode << "        vParams[nActualParams].plVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_UINT_PTR:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_UINT_PTR;\n";
                    aCode << "        vParams[nActualParams].plVal = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_PTR:
                    if (vVtblFuncTable[nFunc]
                            .mpFuncDesc->lprgelemdescParam[nParam]
                            .tdesc.lptdesc->vt
                        == VT_VARIANT)
                    {
                        aCode << "        vParams[nActualParams] = *" << sParamName << ";\n";
                        aCode << "        nActualParams++;\n";
                    }
                    else
                    {
                        // FIXME: Probably wrong for instance for SAFEARRAY(BSTR)* parameters
                        aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                        aCode << "        vParams[nActualParams].vt = VT_PTR;\n";
                        aCode << "        vParams[nActualParams].byref = " << sParamName << ";\n";
                        aCode << "        nActualParams++;\n";
                    }
                    break;
                case VT_USERDEFINED:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_USERDEFINED;\n";
                    aCode << "        vParams[nActualParams].byref = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_LPSTR:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_LPSTR;\n";
                    aCode << "        vParams[nActualParams].byref = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                case VT_LPWSTR:
                    aCode << "        VariantInit(&vParams[" << nParam << "]);\n";
                    aCode << "        vParams[nActualParams].vt = VT_LPWSTR;\n";
                    aCode << "        vParams[nActualParams].byref = " << sParamName << ";\n";
                    aCode << "        nActualParams++;\n";
                    break;
                default:
                    if (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc.vt
                        & VT_BYREF)
                    {
                        aCode << "        VariantInit(&vParams[nActualParams]);\n";
                        aCode
                            << "        vParams[nActualParams].vt = "
                            << vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc.vt
                            << ";\n";
                        aCode << "        vParams[nActualParams].byref = " << sParamName << ";";
                        aCode << "        nActualParams++;\n";
                    }
                    else
                    {
                        std::cerr
                            << "Unhandled type of parameter " << nParam << ": "
                            << vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc.vt
                            << "\n";
                        std::exit(1);
                    }
            }
            aCode << "    }\n";

            if (bGenerateTracing)
            {
                aCode << "    if(!bGotAll)\n";
                switch (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc.vt)
                {
                    case VT_I2:
                    case VT_I4:
                    case VT_R4:
                    case VT_R8:
                    case VT_BSTR:
                    case VT_DISPATCH:
                    case VT_BOOL:
                    case VT_UI2:
                    case VT_UI4:
                    case VT_I8:
                    case VT_UI8:
                    case VT_INT:
                    case VT_UINT:
                    case VT_INT_PTR:
                    case VT_UINT_PTR:
                    case VT_LPSTR:
                    case VT_LPWSTR:
                        aCode << "        std::wcout";
                        if (nParam > 0)
                            aCode << " << \",\"";
                        aCode << " << " << sParamName << ";\n";
                        break;
                    case VT_VARIANT:
                        aCode << "        std::wcout";
                        if (nParam > 0)
                            aCode << " << \",\"";
                        aCode << " << \"<VARIANT:\" << " << sParamName << ".vt << \">\";\n";
                        break;
                    case VT_PTR:
                        if (vVtblFuncTable[nFunc]
                                .mpFuncDesc->lprgelemdescParam[nParam]
                                .tdesc.lptdesc->vt
                            == VT_VARIANT)
                        {
                            aCode << "        switch(" << sParamName << "->vt)\n";
                            aCode << "        {\n";
                            aCode << "        case VT_I2: std::wcout << " << sParamName
                                  << "->iVal; break;\n";
                            aCode << "        case VT_I4: std::wcout << " << sParamName
                                  << "->lVal; break;\n";
                            aCode << "        case VT_R4: std::wcout << " << sParamName
                                  << "->fltVal; break;\n";
                            aCode << "        case VT_R8: std::wcout << " << sParamName
                                  << "->dblVal; break;\n";
                            aCode << "        case VT_BSTR: std::wcout << " << sParamName
                                  << "->bstrVal; break;\n";
                            aCode << "        case VT_DISPATCH: std::wcout << " << sParamName
                                  << "->pdispVal; break;\n";
                            aCode << "        case VT_BOOL: std::wcout << " << sParamName
                                  << "->boolVal; break;\n";
                            aCode << "        case VT_UI2: std::wcout << " << sParamName
                                  << "->uiVal; break;\n";
                            aCode << "        case VT_UI4: std::wcout << " << sParamName
                                  << "->ulVal; break;\n";
                            aCode << "        case VT_I8: std::wcout << " << sParamName
                                  << "->llVal; break;\n";
                            aCode << "        case VT_UI8: std::wcout << " << sParamName
                                  << "->ullVal; break;\n";
                            aCode << "        case VT_INT: std::wcout << " << sParamName
                                  << "->intVal; break;\n";
                            aCode << "        case VT_UINT: std::wcout << " << sParamName
                                  << "->uintVal; break;\n";
                            aCode << "        case VT_INT_PTR: std::wcout << " << sParamName
                                  << "->pintVal; break;\n";
                            aCode << "        case VT_UINT_PTR: std::wcout << " << sParamName
                                  << "->puintVal; break;\n";
                            aCode << "        case VT_LPSTR: std::wcout << " << sParamName
                                  << "->pcVal; break;\n";
                            aCode << "        case VT_LPWSTR: std::wcout << (LPWSTR)" << sParamName
                                  << "->byref; break;\n";
                            aCode << "        default: std::wcout << " << sParamName
                                  << "->byref; break;\n";
                            aCode << "        }\n";
                        }
                        else
                        {
                            aCode << "        std::wcout << " << sParamName << ";\n";
                        }
                        break;
                    case VT_USERDEFINED:
                        aCode << "        std::wcout";
                        if (nParam > 0)
                            aCode << " << \",\"";
                        aCode << " << \"<USERDEFINED>\";\n";
                        break;
                    default:
                        if (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc.vt
                            & VT_BYREF)
                        {
                            if (nParam > 0)
                                aCode << " << \",\";\n";
                            aCode << " << " << sParamName << ";\n";
                        }
                        break;
                }
            }
        }

        if (bGenerateTracing)
            aCode << "    std::wcout << \")\" << std::endl;\n";

        aCode << "    std::vector<VARIANT> vReverseParams;\n";

        if (vVtblFuncTable[nFunc].mpFuncDesc->cParams > 0)
        {
            // Resize the vParams to match the number of actual ones we have. FIXME: Make it correct
            // size from the start.

            aCode << "    vParams.resize(nActualParams);\n";

            // Reverse the order of parameters. Just create another vector.
            aCode << "    for (auto i = vParams.rbegin(); i != vParams.rend(); ++i)\n";
            aCode << "        vReverseParams.push_back(*i);\n";
        }

        // Call CProxiedDispatch::Invoke()
        aCode << "    HRESULT nResult = Invoke(\""
              << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nFunc].mvNames[0]) << "\", "
              << vVtblFuncTable[nFunc].mpFuncDesc->invkind << ", "
              << "vReverseParams, " << sRetvalName << ");\n";
        if (nRetvalParam >= 0)
        {
            switch (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nRetvalParam].tdesc.vt)
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
                case VT_USERDEFINED:
                    // Nothing necessary.
                    break;
                case VT_DISPATCH:
                    // Unclear what todo. We have no idea what the actual interface of the returned
                    // value is, do we? Or should we call GetTypeInfo at run-time? But that would be
                    // little use either. Let's just punt here too and let the returned IDispatch pointer
                    // be returned as such.
                    break;
                case VT_VARIANT:
                    // This one is unclear, too.
                    break;
                case VT_PTR:
                    if (vVtblFuncTable[nFunc]
                                .mpFuncDesc->lprgelemdescParam[nRetvalParam]
                                .tdesc.lptdesc->vt
                            == VT_PTR
                        && vVtblFuncTable[nFunc]
                                   .mpFuncDesc->lprgelemdescParam[nRetvalParam]
                                   .tdesc.lptdesc->lptdesc->vt
                               == VT_USERDEFINED
                        && !IsIgnoredUserdefinedType(
                               pVtblTypeInfo, *vVtblFuncTable[nFunc]
                                                   .mpFuncDesc->lprgelemdescParam[nRetvalParam]
                                                   .tdesc.lptdesc->lptdesc))
                    {
                        ITypeInfo* pReferencedTypeInfo;
                        nResult = pVtblTypeInfo->GetRefTypeInfo(
                            vVtblFuncTable[nFunc]
                                .mpFuncDesc->lprgelemdescParam[nRetvalParam]
                                .tdesc.lptdesc->lptdesc->hreftype,
                            &pReferencedTypeInfo);
                        if (FAILED(nResult))
                        {
                            std::cerr << "GetRefTypeInfo failed\n";
                            std::exit(1);
                        }
                        TYPEATTR* pReferencedTypeAttr;
                        nResult = pReferencedTypeInfo->GetTypeAttr(&pReferencedTypeAttr);
                        if (FAILED(nResult))
                        {
                            std::cerr << "GetTypeAttr failed\n";
                            std::exit(1);
                        }
                        if (pReferencedTypeAttr->typekind == TKIND_DISPATCH
                            || pReferencedTypeAttr->typekind == TKIND_COCLASS)
                        {
                            aCode << "    if (nResult == S_OK)\n";
                            aCode << "        *"
                                  << aUTF16ToUTF8.to_bytes(
                                         vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                                  << " = new "
                                  << TypeToString(sLibName, pVtblTypeInfo,
                                                  *vVtblFuncTable[nFunc]
                                                       .mpFuncDesc->lprgelemdescParam[nRetvalParam]
                                                       .tdesc.lptdesc->lptdesc,
                                                  sReferencedName)
                                  << "(*"
                                  << aUTF16ToUTF8.to_bytes(
                                         vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                                  << ");\n";
                        }
                        pReferencedTypeInfo->ReleaseTypeAttr(pReferencedTypeAttr);
                    }
                    break;
                default:
                    std::cerr << "Unhandled return value type: "
                              << vVtblFuncTable[nFunc]
                                     .mpFuncDesc->lprgelemdescParam[nRetvalParam]
                                     .tdesc.vt
                              << "\n";
                    std::exit(1);
            }
        }
        aCode << "    return nResult;\n";
        aCode << "}\n";
    }

    // Generate hxx file

    if (aIncludedHeaders.size())
        aHeader << "\n";

    // Two constructors: One that takes an extra IID (in case this is the default interface for a
    // coclass).
    aHeader << "class C" << sClass << ": public CProxiedDispatch\n";
    aHeader << "{\n";
    aHeader << "public:\n";
    aHeader << "    C" << sClass << "(IDispatch* pDispatchToProxy) :\n";
    aHeader << "        CProxiedDispatch(pDispatchToProxy, " << IID_initializer(pVtblTypeAttr->guid)
            << ")\n";
    aHeader << "    {\n";
    aHeader << "    }\n";
    aHeader << "\n";
    aHeader << "    C" << sClass << "(IDispatch* pDispatchToProxy, const IID& aExtraIID) :\n";
    aHeader << "        CProxiedDispatch(pDispatchToProxy, " << IID_initializer(pVtblTypeAttr->guid)
            << ", aExtraIID)\n";
    aHeader << "    {\n";
    aHeader << "    }\n";
    aHeader << "\n";

    for (UINT nFunc = 0; nFunc < pVtblTypeAttr->cFuncs; ++nFunc)
    {
        aHeader << "    virtual HRESULT ";
        switch (vVtblFuncTable[nFunc].mpFuncDesc->callconv)
        {
            case CC_CDECL:
                aHeader << "__cdecl";
                break;
            case CC_STDCALL:
                aHeader << "__stdcall";
                break;
            default:
                std::abort();
        }
        aHeader << " ";

        if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYGET)
            aHeader << "get";
        else if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYPUT)
            aHeader << "put";
        else if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYPUTREF)
            aHeader << "putref";
        aHeader << aUTF16ToUTF8.to_bytes(vVtblFuncTable[nFunc].mvNames[0]) << "(";
        for (int nParam = 0; nParam < vVtblFuncTable[nFunc].mpFuncDesc->cParams; ++nParam)
        {
            if (IsIgnoredUserdefinedType(
                    pVtblTypeInfo,
                    vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc))
                aHeader << "void*";
            else
                aHeader << TypeToString(
                    sLibName, pVtblTypeInfo,
                    vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc,
                    sReferencedName);

            if (nParam + 1 < vVtblFuncTable[nFunc].mpFuncDesc->cParams)
                aHeader << ", ";
        }
        aHeader << ");\n";
    }

    // Finish up

    aHeader << "};\n";
    aHeader << "\n";

    // This is a bit unconventional, to #include other headers at the *end*, but one way to get out
    // of silly Catch-22 where you either get circularity problems or undefinedness problems.
    for (auto i : aIncludedHeaders)
        aHeader << "#include \"" << i << ".hxx\"\n";
    aHeader << "\n";

    aHeader << "#endif // INCLUDED_C" << sClass << "_HXX\n";

    aHeader.close();
    if (!aHeader.good())
    {
        std::cerr << "Problems writing to '" << sHeader << "'\n";
        std::exit(1);
    }

    aCode << "\n";

    aCode.close();
    if (!aCode.good())
    {
        std::cerr << "Problems writing to '" << sCode << "'\n";
        std::exit(1);
    }
}

static void Generate(const std::string& sLibName, ITypeInfo* const pTypeInfo)
{
    HRESULT nResult;

    if (IsIgnoredType(pTypeInfo))
        return;

    BSTR sNameBstr;
    nResult = pTypeInfo->GetDocumentation(MEMBERID_NIL, &sNameBstr, NULL, NULL, NULL);
    if (FAILED(nResult))
    {
        std::cerr << "GetDocumentation failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
        std::exit(1);
    }

    const std::string sName = aUTF16ToUTF8.to_bytes(sNameBstr);

    TYPEATTR* pTypeAttr;
    nResult = pTypeInfo->GetTypeAttr(&pTypeAttr);
    if (FAILED(nResult))
    {
        std::cerr << "GetTypeAttr failed for '" << sName << "' in '" << sLibName << ": "
                  << WindowsErrorStringFromHRESULT(nResult) << "\n";
        std::exit(1);
    }

    if (aAlreadyHandledIIDs.count(pTypeAttr->guid))
        return;
    aAlreadyHandledIIDs.insert(pTypeAttr->guid);

    // Check for a coclass first
    switch (pTypeAttr->typekind)
    {
        case TKIND_COCLASS:
            GenerateCoclass(sLibName, sName, pTypeInfo, pTypeAttr);
            break;
        case TKIND_ENUM:
            GenerateEnum(sLibName, sName, pTypeInfo, pTypeAttr);
            break;
        case TKIND_DISPATCH:
            GenerateDispatch(sLibName, sName, pTypeInfo, pTypeAttr);
            break;
        default:
            std::cerr << "Unexpected type kind " << pTypeAttr->typekind << "?\n";
            std::exit(1);
    }
}

static void GenerateCallbackInvoker()
{
    if (aCallbacks.size() == 0)
        return;

    const std::string sHeader = sOutputFolder + "/CallbackInvoker.hxx";
    std::ofstream aHeader(sHeader);
    if (!aHeader.good())
    {
        std::cerr << "Could not open '" << sHeader << "' for writing\n";
        std::exit(1);
    }

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_CallbackInvoker_HXX\n";
    aHeader << "#define INCLUDED_CallbackInvoker_HXX\n";
    aHeader << "\n";

    aHeader << "#include <codecvt>\n";
    aHeader << "#include <string>\n";
    aHeader << "\n";

    for (auto aCallback : aCallbacks)
    {
        aHeader << "#include \"" << aCallback.msLibName << "_" << aCallback.msName << ".hxx\"\n";
    }
    aHeader << "\n";

    aHeader << "static HRESULT "
            << "ProxiedCallbackInvoke(const IID& aIID, IDispatch* pDispatchToProxy,\n";
    aHeader << "                   DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,\n";
    aHeader << "                   DISPPARAMS* pDispParams, VARIANT* pVarResult,\n";
    aHeader << "                   EXCEPINFO* pExcepInfo, UINT* puArgErr)\n";
    aHeader << "{\n";

    for (auto aCallback : aCallbacks)
    {
        aHeader << "    const IID aIID_" << aCallback.msLibName << "_" << aCallback.msName << " = "
                << IID_initializer(aCallback.maIID) << ";\n";
        aHeader << "    if (IsEqualIID(aIID, aIID_" << aCallback.msLibName << "_"
                << aCallback.msName << "))\n";
        aHeader << "        return " << aCallback.msLibName << "_" << aCallback.msName
                << "CallbackInvoke(pDispatchToProxy,\n";
        aHeader << "                   dispIdMember, riid, lcid, wFlags,\n";
        aHeader << "                   pDispParams, pVarResult,\n";
        aHeader << "                   pExcepInfo, puArgErr);\n";
    }

    aHeader << "\n";
    aHeader << "    wchar_t* sIID = L\"\";\n";
    aHeader << "    StringFromIID(aIID, &sIID);\n";

    aHeader << "    std::cerr << \"ProxiedCallbackInvoke: Not prepared to handle IID \" << "
               "std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t>().to_bytes(sIID) << "
               "std::endl;\n";
    aHeader << "    std::abort();\n";
    aHeader << "}\n";

    aHeader << "\n";
    aHeader << "#endif // INCLUDED_CallbackInvoker_HXX\n";
    aHeader.close();
    if (!aHeader.good())
    {
        std::cerr << "Problems writing to '" << sHeader << "'\n";
        std::exit(1);
    }
}

static void GenerateDefaultInterfaceCreator()
{
    if (aDefaultInterfaces.size() == 0)
        return;

    const std::string sHeader = sOutputFolder + "/DefaultInterfaceCreator.hxx";
    std::ofstream aHeader(sHeader);
    if (!aHeader.good())
    {
        std::cerr << "Could not open '" << sHeader << "' for writing\n";
        std::exit(1);
    }

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_DefaultInterfaceCreator_HXX\n";
    aHeader << "#define INCLUDED_DefaultInterfaceCreator_HXX\n";
    aHeader << "\n";

    aHeader << "#include <codecvt>\n";
    aHeader << "#include <string>\n";
    aHeader << "\n";

    for (auto aDefaultInterface : aDefaultInterfaces)
    {
        aHeader << "#include \"C" << aDefaultInterface.msLibName << "_" << aDefaultInterface.msName
                << ".hxx\"\n";
    }
    aHeader << "\n";

    aHeader << "static bool DefaultInterfaceCreator(const IID& rIID, IDispatch** pPDispatch, "
               "IDispatch* pReplacementAppDispatch, std::wstring& sClass)\n";
    aHeader << "{\n";

    for (auto aDefaultInterface : aDefaultInterfaces)
    {
        aHeader << "    const IID aIID_" << aDefaultInterface.msLibName << "_"
                << aDefaultInterface.msName << " = " << IID_initializer(aDefaultInterface.maIID)
                << ";\n";
        aHeader << "    if (IsEqualIID(rIID, aIID_" << aDefaultInterface.msLibName << "_"
                << aDefaultInterface.msName << "))\n";
        aHeader << "    {\n";
        aHeader << "        *pPDispatch = new C" << aDefaultInterface.msLibName << "_"
                << aDefaultInterface.msName << "(pReplacementAppDispatch);\n";
        aHeader << "        sClass = L\"" << aDefaultInterface.msLibName << "_"
                << aDefaultInterface.msName << "\";\n";
        aHeader << "        return true;\n";
        aHeader << "    }\n";
    }

    aHeader << "    return false;\n";
    aHeader << "}\n";

    aHeader << "\n";
    aHeader << "#endif // INCLUDED_DefaultInterfaceCreator_HXX\n";
    aHeader.close();
    if (!aHeader.good())
    {
        std::cerr << "Problems writing to '" << sHeader << "'\n";
        std::exit(1);
    }
}

static void GenerateOutgoingInterfaceMap(const std::map<IID, IID>& aOutgoingInterfaceMap)
{
    if (aOutgoingInterfaceMap.size() == 0)
        return;

    const std::string sHeader = sOutputFolder + "/OutgoingInterfaceMap.hxx";
    std::ofstream aHeader(sHeader);
    if (!aHeader.good())
    {
        std::cerr << "Could not open '" << sHeader << "' for writing\n";
        std::exit(1);
    }

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_OutgoingInterfaceMap_HXX\n";
    aHeader << "#define INCLUDED_OutgoingInterfaceMap_HXX\n";
    aHeader << "\n";

    aHeader << "const static struct { IID maOtherAppIID, maLibreOfficeIID; } "
               "aOutgoingInterfaceMap[] =\n";
    aHeader << "{\n";

    for (const auto i : aOutgoingInterfaceMap)
    {
        aHeader << "    { " << IID_initializer(i.first) << ", " << IID_initializer(i.second)
                << "},\n";
    }

    aHeader << "};\n";
    aHeader << "\n";
    aHeader << "#endif // INCLUDED__OutgoingInterfaceMap_HXX\n";
    aHeader.close();
    if (!aHeader.good())
    {
        std::cerr << "Problems writing to '" << sHeader << "'\n";
        std::exit(1);
    }
}

static void Usage(char** argv)
{
    const char* const pBackslash = std::strrchr(argv[0], L'\\');
    const char* const pSlash = std::strrchr(argv[0], L'/');
    char* pProgram
        = _strdup((pBackslash && pSlash)
                      ? ((pBackslash > pSlash) ? (pBackslash + 1) : (pSlash + 1))
                      : (pBackslash ? (pBackslash + 1) : (pSlash ? (pSlash + 1) : argv[0])));
    char* const pDot = std::strrchr(pProgram, '.');
    if (pDot && _stricmp(pDot, ".exe") == 0)
        *pDot = '\0';

    std::cerr << "Usage: " << pProgram
              << " [options] typelibrary[:interface] ...\n"
                 "\n"
                 "  Options:\n"
                 "    -d directory                 Directory where to write generated files.\n"
                 "                                 Default: \"generated\".\n"
                 "    -i ifaceA,ifaceB,...         Only output code for those interfaces\n"
                 "    -I file                      Only output code for interfaces listed in file\n"
                 "    -M file                      Use mapping between LibreOffice outgoing "
                 "interface IIDs\n"
                 "                                 and the other application's source interface "
                 "IIDs in file\n"
                 "    -T                           Generate verbose tracing outout.\n"
                 "  For instance: "
              << argv[0]
              << " -i _Application,Documents,Document foo.olb:Application bar.exe:SomeInterface\n";
    std::exit(1);
}

int main(int argc, char** argv)
{
    int argi = 1;

    std::map<IID, IID> aOutgoingInterfaceMap;

    while (argi < argc && argv[argi][0] == '-')
    {
        switch (argv[argi][1])
        {
            case 'd':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                sOutputFolder = argv[argi + 1];
                argi++;
                break;
            }
            case 'i':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                char* p = argv[argi + 1];
                char* q;
                while ((q = std::strchr(p, ',')) != NULL)
                {
                    aOnlyTheseInterfaces.insert(std::string(p, (unsigned)(q - p)));
                    p = q + 1;
                }
                aOnlyTheseInterfaces.insert(std::string(p));
                argi++;
                break;
            }
            case 'I':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                std::ifstream aIFile(argv[argi + 1]);
                if (!aIFile.good())
                {
                    std::cerr << "Could not open " << argv[argi + 1] << " for reading\n";
                    std::exit(1);
                }
                while (!aIFile.eof())
                {
                    std::string sIface;
                    aIFile >> sIface;
                    if (aIFile.good())
                        aOnlyTheseInterfaces.insert(sIface);
                }
                argi++;
                break;
            }
            case 'M':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                std::ifstream aMFile(argv[argi + 1]);
                if (!aMFile.good())
                {
                    std::cerr << "Could not open " << argv[argi + 1] << " for reading\n";
                    std::exit(1);
                }
                while (!aMFile.eof())
                {
                    std::string sLibreOfficeOutgoingInterface;
                    aMFile >> sLibreOfficeOutgoingInterface;
                    if (sLibreOfficeOutgoingInterface.length() > 0
                        && sLibreOfficeOutgoingInterface[0] == '#')
                    {
                        std::string sRestOfLine;
                        std::getline(aMFile, sRestOfLine);
                        continue;
                    }
                    std::string sOtherAppSourceInterface;
                    aMFile >> sOtherAppSourceInterface;
                    if (aMFile.good())
                    {
                        IID aLibreOfficeOutgoingInterface;
                        if (FAILED(IIDFromString(
                                aUTF8ToUTF16.from_bytes(sLibreOfficeOutgoingInterface.c_str())
                                    .data(),
                                &aLibreOfficeOutgoingInterface)))
                        {
                            std::cerr << "Could not interpret " << sLibreOfficeOutgoingInterface
                                      << " in " << argv[argi + 1] << " as an IID\n";
                            break;
                        }
                        IID aOtherAppSourceInterface;
                        if (FAILED(IIDFromString(
                                aUTF8ToUTF16.from_bytes(sOtherAppSourceInterface.c_str()).data(),
                                &aOtherAppSourceInterface)))
                        {
                            std::cerr << "Count not interpret " << sOtherAppSourceInterface
                                      << " in " << argv[argi + 1] << " as an IID\n";
                            break;
                        }
                        aOutgoingInterfaceMap[aLibreOfficeOutgoingInterface]
                            = aOtherAppSourceInterface;
                    }
                }
                argi++;
                break;
            }
            case 'T':
                bGenerateTracing = true;
                break;
            default:
                Usage(argv);
        }
        argi++;
    }

    if (argc - argi < 1)
        Usage(argv);

    CoInitialize(NULL);

    for (; argi < argc; ++argi)
    {
        char* const pColon = std::strchr(argv[argi], ':');
        const char* pInterface = "Application";
        if (pColon)
        {
            *pColon = '\0';
            pInterface = pColon + 1;
        }

        HRESULT nResult;
        ITypeLib* pTypeLib;
        // FIXME: argv is not in UTF-8 but in system codepage. But no big deal, only developers run this
        // program and they should know to not use non-ASCII characters in their folder and file names,
        // one hopes.
        nResult
            = LoadTypeLibEx(aUTF8ToUTF16.from_bytes(argv[argi]).data(), REGKIND_NONE, &pTypeLib);
        if (FAILED(nResult))
        {
            std::cerr << "Could not load '" << argv[argi]
                      << "' as a type library: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            std::exit(1);
        }

        UINT nTypeInfoCount = pTypeLib->GetTypeInfoCount();
        if (nTypeInfoCount == 0)
        {
            std::cerr << "No type information in type library " << argv[argi] << "?\n";
            std::exit(1);
        }

        BSTR sLibNameBstr;
        nResult = pTypeLib->GetDocumentation(-1, &sLibNameBstr, NULL, NULL, NULL);
        if (FAILED(nResult))
        {
            std::cerr << "GetDocumentation(-1) of " << argv[argi]
                      << " failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            std::exit(1);
        }
        std::string sLibName = aUTF16ToUTF8.to_bytes(sLibNameBstr);

        bool bFound = false;
        for (UINT i = 0; i < nTypeInfoCount; ++i)
        {
            TYPEKIND nTypeKind;
            nResult = pTypeLib->GetTypeInfoType(i, &nTypeKind);
            if (FAILED(nResult))
            {
                std::cerr << "GetTypeInfoType(" << i << ") of " << argv[argi]
                          << " failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
                std::exit(1);
            }

            if (nTypeKind != TKIND_ENUM && nTypeKind != TKIND_INTERFACE
                && nTypeKind != TKIND_DISPATCH && nTypeKind != TKIND_COCLASS)
                continue;

            BSTR sTypeName;
            nResult = pTypeLib->GetDocumentation((INT)i, &sTypeName, NULL, NULL, NULL);
            if (FAILED(nResult))
            {
                std::cerr << "GetDocumentation(" << i << ") of " << argv[argi]
                          << " failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
                std::exit(1);
            }

            if (std::wcscmp(sTypeName, aUTF8ToUTF16.from_bytes(pInterface).data()) != 0)
                continue;

            ITypeInfo* pTypeInfo;
            nResult = pTypeLib->GetTypeInfo(i, &pTypeInfo);
            if (FAILED(nResult))
            {
                std::cerr << "GetTypeInfo(" << i << ") of " << argv[argi]
                          << " failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
                std::exit(1);
            }

            bFound = true;

            Generate(sLibName, pTypeInfo);

            break;
        }

        if (!bFound)
        {
            std::cerr << "Could not find " << pInterface << " in " << argv[argi] << "\n";
            std::exit(1);
        }
    }

    GenerateOutgoingInterfaceMap(aOutgoingInterfaceMap);

    GenerateCallbackInvoker();

    GenerateDefaultInterfaceCreator();

    CoUninitialize();

    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

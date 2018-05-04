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
#include <set>
#include <sstream>
#include <string>
#include <vector>

#include <Windows.h>

#pragma warning(pop)

#include "interfacemap.hpp"
#include "outgoingmap.hpp"
#include "utils.hpp"

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

struct Dispatch
{
    const std::string msLibName;
    const std::string msName;
    const IID maIID;
};

static std::string sOutputFolder = "generated";

static std::set<IID> aAlreadyHandledIIDs;
static std::set<std::string> aOnlyTheseInterfaces;

static std::set<Callback> aCallbacks;
static std::set<DefaultInterface> aDefaultInterfaces;
static std::vector<Dispatch> aDispatches;
static std::vector<InterfaceMapping> aInterfaceMap;
static std::vector<OutgoingInterfaceMapping> aOutgoingInterfaceMap;

static void Generate(const std::string& sLibName, ITypeInfo* pTypeInfo);

inline bool operator<(const Callback& a, const Callback& b)
{
    return ((a.msLibName < b.msLibName) || ((a.msLibName == b.msLibName) && (a.msName < b.msName)));
}

inline bool operator<(const DefaultInterface& a, const DefaultInterface& b)
{
    return ((a.msLibName < b.msLibName) || ((a.msLibName == b.msLibName) && (a.msName < b.msName)));
}

class OutputFile : public std::ofstream
{
private:
    const std::string msFilename;

    std::string readAll(const std::string& sFilename, std::ifstream& rStream)
    {
        rStream.seekg(0, std::ios::end);
        std::size_t nLength = (std::size_t)rStream.tellg();
        rStream.seekg(0, std::ios::beg);
        std::vector<char> vBuf(nLength + 1);
        rStream.read(vBuf.data(), (std::streamsize)nLength);
        vBuf[nLength] = '\0';
        rStream.close();
        if (!rStream.good())
        {
            std::cerr << "Could not read from '" << sFilename << "'\n";
            std::exit(1);
        }
        return std::string(vBuf.data());
    }

public:
    OutputFile(const std::string& sFilename)
        : std::ofstream(sFilename + ".temp")
        , msFilename(sFilename)
    {
        if (!good())
        {
            std::cerr << "Could not open '" << sFilename << ".temp' for writing\n";
            std::exit(1);
        }
    }

    OutputFile(const OutputFile&) = delete;

    void close()
    {
        std::ofstream::close();
        if (!good())
        {
            std::cerr << "Problems writing to '" << msFilename << ".temp'\n";
            std::exit(1);
        }

        const std::string sTempFilename = msFilename + ".temp";
        std::ifstream aJustWritten(sTempFilename, std::ios::binary);
        if (aJustWritten.good())
        {
            std::string sJustWritten = readAll(sTempFilename, aJustWritten);
            std::ifstream aOld(msFilename, std::ios::binary);
            if (aOld.good())
            {
                std::string sOld = readAll(msFilename, aOld);
                if (sJustWritten == sOld)
                {
                    if (std::remove(sTempFilename.c_str()) != 0)
                        std::cerr << "Could not remove '" << sTempFilename << "'\n";
                    return;
                }
            }
        }
        std::remove(msFilename.c_str());
        if (std::rename(sTempFilename.c_str(), msFilename.c_str()) != 0)
        {
            std::cerr << "Could not rename '" << sTempFilename << "' to '" << msFilename << "'\n";
            std::exit(1);
        }
    }
};

static bool IsIgnoredType(const std::string& sLibName, ITypeInfo* pTypeInfo)
{
    if (aOnlyTheseInterfaces.size() == 0)
        return false;

    HRESULT nResult;
    BSTR sName;
    nResult = pTypeInfo->GetDocumentation(MEMBERID_NIL, &sName, NULL, NULL, NULL);
    if (SUCCEEDED(nResult))
        return aOnlyTheseInterfaces.count(sLibName + "." + std::string(convertUTF16ToUTF8(sName)))
               == 0;

    return false;
}

static bool IsIgnoredUserdefinedType(const std::string& sLibName, ITypeInfo* pTypeInfo,
                                     const TYPEDESC& aTypeDesc)
{
    if (aOnlyTheseInterfaces.size() == 0)
        return false;

    if (aTypeDesc.vt == VT_USERDEFINED)
    {
        HRESULT nResult;
        ITypeInfo* pReferencedTypeInfo;
        nResult = pTypeInfo->GetRefTypeInfo(aTypeDesc.hreftype, &pReferencedTypeInfo);

        if (SUCCEEDED(nResult))
            return IsIgnoredType(sLibName, pReferencedTypeInfo);
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
            sResult = "IUnknown*";
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
        if (IsIgnoredUserdefinedType(sLibName, pTypeInfo, *aTypeDesc.lptdesc))
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
            sResult = "/* USERDEFINED:" + to_uhex(aTypeDesc.hreftype) + " */ void*";
        else
        {
            if (pTypeAttr->typekind == TKIND_DISPATCH || pTypeAttr->typekind == TKIND_COCLASS)
            {
                sReferencedName = "C" + sLibName + "_" + convertUTF16ToUTF8(sName);
                sResult = sReferencedName;
            }
            else if (pTypeAttr->typekind == TKIND_ENUM)
            {
                sReferencedName = "E" + sLibName + "_" + convertUTF16ToUTF8(sName);
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

    if (IsIgnoredType(sLibName, pTypeInfo))
        return;

    BSTR sName;
    nResult = pTypeInfo->GetDocumentation(MEMBERID_NIL, &sName, NULL, NULL, NULL);
    if (FAILED(nResult))
    {
        std::cerr << "GetDocumentation failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
        std::exit(1);
    }

    std::cerr << sLibName << "_" << convertUTF16ToUTF8(sName) << " (sink)\n";

    std::string sClass = sLibName + "_" + convertUTF16ToUTF8(sName);

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

    aCallbacks.insert({ pTypeAttr->guid, sLibName, convertUTF16ToUTF8(sName) });

    const std::string sHeader = sOutputFolder + "/" + sClass + ".hxx";
    OutputFile aHeader(sHeader);

    const std::string sCode = sOutputFolder + "/" + sClass + ".cxx";
    OutputFile aCode(sCode);

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_" << sClass << "_HXX\n";
    aHeader << "#define INCLUDED_" << sClass << "_HXX\n";
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
    aCode << "#include \"CProxiedUnknown.hpp\"\n";
    aCode << "\n";
    aCode << "#include \"" << sClass << ".hxx\"\n";
    aCode << "\n";

    aCode << "HRESULT " << sClass << "CallbackInvoke(IDispatch* pDispatchToProxy,\n";
    aCode << "            DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,\n";
    aCode << "            DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, "
             "UINT* puArgErr)\n";
    aCode << "{\n";

    aCode << "    if (CProxiedUnknown::getParam()->mbVerbose || "
             "CProxiedUnknown::getParam()->mbTraceOnly)\n";
    aCode << "    {\n";
    aCode << "        if (!CProxiedUnknown::mbIsAtBeginningOfLine)\n";
    aCode << "            std::cout << \"\\n\" << CProxiedUnknown::indent();\n";
    aCode << "        std::cout << \"" << sLibName << "." << convertUTF16ToUTF8(sName) << ".\";\n";
    aCode << "    }\n";

    aCode << "    HRESULT nResult;\n";
    aCode << "    DISPPARAMS aLocalDispParams = *pDispParams;\n";
    aCode << "    aLocalDispParams.rgvarg = new VARIANTARG[aLocalDispParams.cArgs];\n";

    aCode << "    switch (dispIdMember)\n";
    aCode << "    {\n";

    // Also fill in the NameToMemberIdMapping in the corresponding OutgoingInterfaceMapping
    OutgoingInterfaceMapping* pOutgoing = nullptr;
    auto pOutgoingNameToId = new std::vector<NameToMemberIdMapping>;
    for (auto& i : aOutgoingInterfaceMap)
    {
        if (IsEqualIID(i.maSourceInterfaceInProxiedApp, pTypeAttr->guid))
        {
            pOutgoing = &i;
            break;
        }
    }

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

        pOutgoingNameToId->push_back(
            { _strdup(convertUTF16ToUTF8(sFuncName).data()), pFuncDesc->memid });

        aCode << "        case " << pFuncDesc->memid << ": // " << convertUTF16ToUTF8(sFuncName)
              << "\n";
        aCode << "            {\n";
        aCode << "                if (CProxiedUnknown::getParam()->mbVerbose || "
                 "CProxiedUnknown::getParam()->mbTraceOnly)\n";
        aCode << "                {\n";
        aCode << "                    std::cout << \"" << convertUTF16ToUTF8(sFuncName)
              << "\" << std::endl;\n";
        aCode << "                    CProxiedUnknown::mbIsAtBeginningOfLine = true;\n";
        aCode << "                }\n";

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
                && !IsIgnoredUserdefinedType(sLibName, pTypeInfo,
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
                    std::string sReferencedTypeName(convertUTF16ToUTF8(sReferencedTypeNameBstr));
                    aCode << "                    aLocalDispParams.rgvarg["
                          << (pFuncDesc->cParams - nParam - 1)
                          << "].pdispVal = reinterpret_cast<IDispatch*>(C" << sLibName << "_"
                          << sReferencedTypeName << "::get(nullptr, aLocalDispParams.rgvarg["
                          << (pFuncDesc->cParams - nParam - 1) << "].pdispVal));\n";

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

    pOutgoingNameToId->push_back({ nullptr, 0 });
    if (pOutgoing != nullptr)
        pOutgoing->maNameToId = pOutgoingNameToId->data();

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
    aHeader << "#endif // INCLUDED_" << sClass << "_HXX\n";

    aHeader.close();
    aCode.close();
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
    OutputFile aHeader(sHeader);
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
        aHeader << "    " << convertUTF16ToUTF8(sVarName) << " = "
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

    // Fill in the information in aInterfaceMap.
    for (auto& i : aInterfaceMap)
    {
        if (IsEqualIID(i.maFromCoclass, pTypeAttr->guid))
        {
            i.msFromLibName = _strdup(sLibName.c_str());

            BSTR sCoclassName;
            nResult = pTypeInfo->GetDocumentation(MEMBERID_NIL, &sCoclassName, NULL, NULL, NULL);
            if (FAILED(nResult))
                i.msFromCoclassName = _strdup(IID_to_string(i.maFromCoclass).c_str());
            else
                i.msFromCoclassName = _strdup(convertUTF16ToUTF8(sCoclassName).c_str());
            break;
        }
    }

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
        if (FAILED(nResult))
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
                sDefaultInterface = convertUTF16ToUTF8(sImplTypeName);
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
    OutputFile aHeader(sHeader);

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
    aHeader << "    C" << sClass
            << "(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy) :\n";
    aHeader << "        C" << sLibName << "_" << sDefaultInterface
            << "(pBaseClassUnknown, pDispatchToProxy, " << IID_initializer(pTypeAttr->guid)
            << ")\n";
    aHeader << "    {\n";
    aHeader << "    }\n";
    aHeader << "\n";
    aHeader << "    static C" << sClass
            << "* get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy)\n";
    aHeader << "    {\n";
    aHeader << "        CProxiedUnknown* pExisting = find(pDispatchToProxy);\n";
    aHeader << "        if (pExisting != nullptr)\n";
    aHeader << "            return static_cast<C" << sClass << "*>(pExisting);";
    aHeader << "\n";
    aHeader << "        return new C" << sClass << "(pBaseClassUnknown, pDispatchToProxy);\n";
    aHeader << "    }\n";
    aHeader << "};\n";

    aHeader << "\n";
    aHeader << "#endif // INCLUDED_C" << sClass << "_HXX\n";

    aHeader.close();
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
            std::cerr << "(" << convertUTF16ToUTF8(vDispFuncTable[nDispFunc].mvNames[0]) << " vs. "
                      << convertUTF16ToUTF8(vVtblFuncTable[nVtblFunc].mvNames[0]) << ")\n";
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
            std::cerr << "(Function: " << convertUTF16ToUTF8(vVtblFuncTable[nVtblFunc].mvNames[0])
                      << ", " << vDispFuncTable[nDispFunc].mpFuncDesc->cParams << " vs. "
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
                std::cerr << "("
                          << convertUTF16ToUTF8(vDispFuncTable[nDispFunc].mvNames[nVtblParam + 1u])
                          << " vs. "
                          << convertUTF16ToUTF8(vVtblFuncTable[nVtblFunc].mvNames[nVtblParam + 1u])
                          << ")\n";
                std::exit(1);
            }

            if (vDispFuncTable[nDispFunc].mpFuncDesc->lprgelemdescParam[nVtblParam].tdesc.vt
                != vVtblFuncTable[nVtblFunc].mpFuncDesc->lprgelemdescParam[nVtblParam].tdesc.vt)
            {
                std::cerr << "Huh, IDispatch-based parameter " << nVtblParam << " of "
                          << convertUTF16ToUTF8(vVtblFuncTable[nVtblFunc].mvNames[0])
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

    aDispatches.push_back({ sLibName, sTypeName, pVtblTypeAttr->guid });

    const std::string sClass = sLibName + "_" + sTypeName;

    // Open output files

    const std::string sHeader = sOutputFolder + "/C" + sClass + ".hxx";
    OutputFile aHeader(sHeader);

    const std::string sCode = sOutputFolder + "/C" + sClass + ".cxx";
    OutputFile aCode(sCode);

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

    // Generate bulk of cxx file

    // First constructors. Two constructors: One that takes an extra IID (in case this is the
    // default interface for a coclass).

    aCode << "C" << sClass << "::C" << sClass
          << "(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy) :\n";
    aCode << "    CProxiedDispatch(pBaseClassUnknown, pDispatchToProxy, "
          << IID_initializer(pVtblTypeAttr->guid) << ", \"" << sLibName << "\")\n";
    aCode << "{\n";
    aCode << "}\n";
    aCode << "\n";
    aCode << "C" << sClass << "::C" << sClass
          << "(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy, const IID& aIID) :\n";
    aCode << "    CProxiedDispatch(pBaseClassUnknown, pDispatchToProxy, "
          << IID_initializer(pVtblTypeAttr->guid) << ", aIID, \"" << sLibName << "\")\n";
    aCode << "{\n";
    aCode << "}\n";
    aCode << "\n";

    // The the get() functions
    aCode << "C" << sClass << "* C" << sClass
          << "::get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy)\n";
    aCode << "{\n";
    aCode << "    return get(pBaseClassUnknown, pDispatchToProxy, IID_NULL);\n";
    aCode << "}\n";
    aCode << "\n";
    aCode << "C" << sClass << "* C" << sClass
          << "::get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy, const IID& aIID)\n";
    aCode << "{\n";
    aCode << "    CProxiedUnknown* pExisting = find(pDispatchToProxy);\n";
    aCode << "    if (pExisting != nullptr)\n";
    aCode << "        return static_cast<C" << sClass << "*>(pExisting);";
    aCode << "\n";
    aCode << "    return new C" << sClass << "(pBaseClassUnknown, pDispatchToProxy, aIID);\n";
    aCode << "}\n";
    aCode << "\n";

    // Then the interface member functions

    for (UINT nFunc = 0; nFunc < pVtblTypeAttr->cFuncs; ++nFunc)
    {
        // Generate the code for one method. They are all of type HRESULT, no?

        aCode << "// vtbl entry " << (vVtblFuncTable[nFunc].mpFuncDesc->oVft / sizeof(void*))
              << ", member id " << vVtblFuncTable[nFunc].mpFuncDesc->memid << "\n";

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
        sFuncName += convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[0]);
        aCode << sFuncName << "(";

        // Parameter list
        for (int nParam = 0; nParam < vVtblFuncTable[nFunc].mpFuncDesc->cParams; ++nParam)
        {
            // FIXME: If a parameter is an enum type that we don't bother generating a C++ enum for,
            // we should just use an integral type for it, not void*. Need to check the type
            // description for the underlying integral type, though.
            if (IsIgnoredUserdefinedType(
                    sLibName, pVtblTypeInfo,
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
                    // Just output a forward declaration in case of circular dependnecy between headers
                    aHeader << "class " << sReferencedName << ";\n";
                    aIncludedHeaders.insert(sReferencedName);
                }
            }
            if ((size_t)(nParam + 1) < vVtblFuncTable[nFunc].mvNames.size())
                aCode << convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[nParam + 1u]);
            else
                aCode << "p" << nParam;
            if (nParam + 1 < vVtblFuncTable[nFunc].mpFuncDesc->cParams)
                aCode << ", ";
        }
        aCode << ")\n";

        // Code block of the function: Package parameters into an array of VARIANTs, and call our
        // magic CProxiedDispatch::Invoke().

        aCode << "{\n";

        aCode << "    if (getParam()->mbVerbose || getParam()->mbTraceOnly)\n";
        aCode << "    {\n";
        aCode << "        std::cout << indent() << \"" << sLibName << "." << sTypeName
              << "<\" << (mpBaseClassUnknown ? mpBaseClassUnknown : this) << \">."
              << convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[0]);
        if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYGET)
            aCode << "\";\n";
        else if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYPUT
                 || vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYPUTREF)
            aCode << " = \";\n";
        else if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_FUNC)
            aCode << "(\";\n";
        else
            assert(!"Unexpected invkind");
        aCode << "        mbIsAtBeginningOfLine = false;\n";
        aCode << "    }\n";

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
                sParamName = convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[nParam + 1u]);
            else
                sParamName = "p" + std::to_string(nParam);

            if (bHadOptional
                && !(vVtblFuncTable[nFunc]
                         .mpFuncDesc->lprgelemdescParam[nParam]
                         .paramdesc.wParamFlags
                     & (PARAMFLAG_FOPT | PARAMFLAG_FRETVAL | PARAMFLAG_FLCID)))
            {
                std::cerr << "Huh, Optional parameter "
                          << convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[(UINT)nParam - 1u])
                          << " to function " << convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[0])
                          << " followed by non-optional "
                          << convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[(UINT)nParam])
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
                      << convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[nParam + 1u]) << ";\n";

                aCode
                    << "    if (!bGotAll && (getParam()->mbVerbose || getParam()->mbTraceOnly))\n";
                aCode << "        std::cout";
                if (nParam > 0)
                    aCode << " << \",\"";
                aCode << " << " << sParamName << ";\n";
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
                    aCode << "            VariantInit(&vParams[nActualParams]);\n";
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

            aCode << "    if (!bGotAll && (getParam()->mbVerbose || getParam()->mbTraceOnly))\n";
            aCode << "    {\n";
            aCode << "        std::cout";
            if (nParam > 0)
                aCode << " << \",\"";
            switch (vVtblFuncTable[nFunc].mpFuncDesc->lprgelemdescParam[nParam].tdesc.vt)
            {
                case VT_I2:
                case VT_I4:
                case VT_R4:
                case VT_R8:
                case VT_BSTR:
                case VT_DISPATCH:
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
                    aCode << " << " << sParamName << ";\n";
                    break;
                case VT_BOOL:
                    aCode << " << (" << sParamName << " ? \"True\" : \"False\");\n";
                    break;
                case VT_VARIANT:
                    aCode << " << \"\";\n";
                    aCode << "        if (" << sParamName << ".vt == VT_ERROR\n";
                    aCode << "            && " << sParamName << ".scode == DISP_E_PARAMNOTFOUND)\n";
                    aCode << "            std::cout << \"(empty)\";\n";
                    aCode << "        else\n";
                    aCode << "            std::cout << " << sParamName << ";\n";
                    break;
                case VT_PTR:
                    if (vVtblFuncTable[nFunc]
                            .mpFuncDesc->lprgelemdescParam[nParam]
                            .tdesc.lptdesc->vt
                        == VT_VARIANT)
                    {
                        aCode << " << *" << sParamName << ";\n";
                    }
                    else
                    {
                        aCode << " << " << sParamName << ";\n";
                    }
                    break;
                case VT_USERDEFINED:
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
            aCode << "    }\n";
        }

        if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_FUNC)
        {
            aCode << "    if (getParam()->mbVerbose || getParam()->mbTraceOnly)\n";
            aCode << "        std::cout << \")";
            aCode << "\";\n";
        }

        aCode << "    std::vector<VARIANT> vReverseParams;\n";

        if (vVtblFuncTable[nFunc].mpFuncDesc->cParams > 0)
        {
            // Drop trailing empty parameters
            aCode << "    while (nActualParams > 0 && vParams[nActualParams-1].vt == VT_EMPTY)\n";
            aCode << "      nActualParams--;\n";

            // Resize the vParams to match the number of actual ones we have. FIXME: Make it correct
            // size from the start.
            aCode << "    vParams.resize(nActualParams);\n";

            // Reverse the order of parameters. Just create another vector.
            aCode << "    for (auto i = vParams.rbegin(); i != vParams.rend(); ++i)\n";
            aCode << "        vReverseParams.push_back(*i);\n";
        }

        // Call CProxiedDispatch::genericInvoke()
        aCode << "    increaseIndent();\n";
        aCode << "    HRESULT nResult = genericInvoke(\""
              << convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[0]) << "\", "
              << vVtblFuncTable[nFunc].mpFuncDesc->invkind << ", "
              << "vReverseParams, " << sRetvalName << ");\n";
        aCode << "    decreaseIndent();\n";
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
                    // Unclear
                    break;
                case VT_DISPATCH:
                    // Unclear what todo. We have no idea what the actual interface of the returned
                    // value is, do we? Or should we call GetTypeInfo at run-time? But that would be
                    // little use either. Let's just punt here too and let the returned IDispatch pointer
                    // be returned as such.
                    if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYGET
                        || vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_FUNC)
                    {
                        aCode << "        if (getParam()->mbVerbose || "
                                 "getParam()->mbTraceOnly)\n";
                        aCode << "            std::cout << *"
                              << convertUTF16ToUTF8(
                                     vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                              << ";\n";
                    }
                    break;
                case VT_VARIANT:
                    // This one is unclear, too.
                    if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYGET
                        || vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_FUNC)
                    {
                        aCode << "        if (getParam()->mbVerbose || "
                                 "getParam()->mbTraceOnly)\n";
                        aCode << "            std::cout << *"
                              << convertUTF16ToUTF8(
                                     vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                              << ";\n";
                    }
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
                               sLibName, pVtblTypeInfo,
                               *vVtblFuncTable[nFunc]
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
                                  << convertUTF16ToUTF8(
                                         vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                                  << " = "
                                  << TypeToString(sLibName, pVtblTypeInfo,
                                                  *vVtblFuncTable[nFunc]
                                                       .mpFuncDesc->lprgelemdescParam[nRetvalParam]
                                                       .tdesc.lptdesc->lptdesc,
                                                  sReferencedName)
                                  << "::get(nullptr, reinterpret_cast<IDispatch*>(*"
                                  << convertUTF16ToUTF8(
                                         vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                                  << "));\n";

                            // FIXME: The code snippet below is copy-pasted right below, and the
                            // inner part of it once more.
                            if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYGET
                                || vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_FUNC)
                            {
                                aCode << "    if (getParam()->mbVerbose || "
                                         "getParam()->mbTraceOnly)\n";
                                aCode << "    {\n";
                                aCode << "        if (nResult == S_OK)\n";
                                aCode << "            std::cout << \" -> \" << *"
                                      << convertUTF16ToUTF8(
                                             vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                                      << " << std::endl;\n";
                                aCode << "        else\n";
                                aCode << "            std::cout << \": \" << "
                                         "HRESULT_to_string(nResult) << std::endl;\n";
                                aCode << "        mbIsAtBeginningOfLine = true;\n";
                                aCode << "    }\n";
                            }
                        }
                        pReferencedTypeInfo->ReleaseTypeAttr(pReferencedTypeAttr);
                    }
                    else if (vVtblFuncTable[nFunc].mpFuncDesc->memid == DISPID_NEWENUM
                             && vVtblFuncTable[nFunc]
                                        .mpFuncDesc->lprgelemdescParam[nRetvalParam]
                                        .tdesc.vt
                                    == VT_PTR
                             && vVtblFuncTable[nFunc]
                                        .mpFuncDesc->lprgelemdescParam[nRetvalParam]
                                        .tdesc.lptdesc->vt
                                    == VT_UNKNOWN)
                    {
                        if (!aIncludedHeaders.count("CProxiedEnumVARIANT"))
                            aHeader << "#include \"CProxiedEnumVARIANT.hpp\"\n";

                        aCode << "    if (nResult == S_OK)\n";
                        aCode
                            << "        *"
                            << convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                            << " = new CProxiedEnumVARIANT(*"
                            << convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                            << ", \"" << sLibName << "\");\n";

                        aCode << "    if (getParam()->mbVerbose || "
                                 "getParam()->mbTraceOnly)\n";
                        aCode << "    {\n";
                        aCode << "        if (nResult == S_OK)\n";
                        aCode << "            std::cout << *"
                              << convertUTF16ToUTF8(
                                     vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                              << " << std::endl;\n";
                        aCode << "        else\n";
                        aCode << "            std::cout << \": \" << HRESULT_to_string(nResult) << "
                                 "std::endl;\n";
                        aCode << "        mbIsAtBeginningOfLine = true;\n";
                        aCode << "    }\n";
                    }
                    else if (isDirectlyPrintableType(
                                 vVtblFuncTable[nFunc]
                                     .mpFuncDesc->lprgelemdescParam[nRetvalParam]
                                     .tdesc.lptdesc->vt))
                    {
                        if (vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_PROPERTYGET
                            || vVtblFuncTable[nFunc].mpFuncDesc->invkind == INVOKE_FUNC)
                        {
                            aCode << "    if (getParam()->mbVerbose || "
                                     "getParam()->mbTraceOnly)\n";
                            aCode << "    {\n";
                            aCode << "        if (nResult == S_OK)\n";
                            aCode << "            std::cout << \" -> \" << *"
                                  << convertUTF16ToUTF8(
                                         vVtblFuncTable[nFunc].mvNames[nRetvalParam + 1u])
                                  << " << std::endl;\n";
                            aCode << "        else\n";
                            aCode << "            std::cout << \": \" << "
                                     "HRESULT_to_string(nResult) << "
                                     "std::endl;\n";
                            aCode << "        mbIsAtBeginningOfLine = true;\n";
                            aCode << "    }\n";
                        }
                    }
                    else
                    {
                        aCode << "    if (getParam()->mbVerbose || "
                                 "getParam()->mbTraceOnly)\n";
                        aCode << "    {\n";
                        aCode << "        if (nResult == S_OK)\n";
                        aCode << "            std::cout << \"?\\n\";\n";
                        aCode << "        else\n";
                        aCode << "            std::cout << \": \" << HRESULT_to_string(nResult) << "
                                 "std::endl;\n";
                        aCode << "        mbIsAtBeginningOfLine = true;\n";
                        aCode << "    }\n";
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
        else
        {
            aCode << "    if (getParam()->mbVerbose || getParam()->mbTraceOnly)\n";
            aCode << "    {\n";
            aCode << "        std::cout << std::endl;\n";
            aCode << "        mbIsAtBeginningOfLine = true;\n";
            aCode << "    }\n";
        }
        aCode << "    return nResult;\n";
        aCode << "}\n";
    }

    // Generate hxx file

    if (aIncludedHeaders.size())
        aHeader << "\n";

    aHeader << "class C" << sClass << ": public CProxiedDispatch\n";
    aHeader << "{\n";
    aHeader << "public:\n";
    aHeader << "    C" << sClass << "(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy);\n";
    aHeader << "\n";
    aHeader << "    C" << sClass
            << "(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy, const IID& aIID);\n";
    aHeader << "\n";
    aHeader << "    static C" << sClass
            << "* get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy);\n";
    aHeader << "\n";
    aHeader
        << "    static C" << sClass
        << "* get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy, const IID& aIID);\n";
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
        aHeader << convertUTF16ToUTF8(vVtblFuncTable[nFunc].mvNames[0]) << "(";
        for (int nParam = 0; nParam < vVtblFuncTable[nFunc].mpFuncDesc->cParams; ++nParam)
        {
            if (IsIgnoredUserdefinedType(
                    sLibName, pVtblTypeInfo,
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
    aCode.close();
}

static void Generate(const std::string& sLibName, ITypeInfo* const pTypeInfo)
{
    HRESULT nResult;

    if (IsIgnoredType(sLibName, pTypeInfo))
        return;

    BSTR sNameBstr;
    nResult = pTypeInfo->GetDocumentation(MEMBERID_NIL, &sNameBstr, NULL, NULL, NULL);
    if (FAILED(nResult))
    {
        std::cerr << "GetDocumentation failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
        std::exit(1);
    }

    const std::string sName = convertUTF16ToUTF8(sNameBstr);

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
    OutputFile aHeader(sHeader);

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_CallbackInvoker_HXX\n";
    aHeader << "#define INCLUDED_CallbackInvoker_HXX\n";
    aHeader << "\n";

    aHeader << "#include \"utils.hpp\"\n";
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

    aHeader << "    std::cerr << \"ProxiedCallbackInvoke: Not prepared to handle IID \" << "
               "IID_to_string(aIID) << std::endl;\n";
    aHeader << "    std::abort();\n";
    aHeader << "}\n";

    aHeader << "\n";
    aHeader << "#endif // INCLUDED_CallbackInvoker_HXX\n";

    aHeader.close();
}

static void GenerateDefaultInterfaceCreator()
{
    if (aDefaultInterfaces.size() == 0)
        return;

    const std::string sHeader = sOutputFolder + "/DefaultInterfaceCreator.hxx";
    OutputFile aHeader(sHeader);

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_DefaultInterfaceCreator_HXX\n";
    aHeader << "#define INCLUDED_DefaultInterfaceCreator_HXX\n";
    aHeader << "\n";

    aHeader << "#include <string>\n";
    aHeader << "\n";

    for (auto aDefaultInterface : aDefaultInterfaces)
    {
        aHeader << "#include \"C" << aDefaultInterface.msLibName << "_" << aDefaultInterface.msName
                << ".hxx\"\n";
    }
    aHeader << "\n";

    aHeader << "static bool DefaultInterfaceCreator(IUnknown* pBaseClassUnknown, const IID& rIID, "
               "IDispatch** pPDispatch, "
               "IDispatch* pReplacementAppDispatch, std::string& sClass)\n";
    aHeader << "{\n";

    // Here we use the constructors, not the get() functions, as we know that we should always be
    // creating new proxy object. This is called from CProxiedCoclass::QueryInterface().

    // Or should we? What if some app calls the coclass object's QueryInterface() for the same IID
    // several times? Oh well, let's fix that if such a situation arises.

    for (auto aDefaultInterface : aDefaultInterfaces)
    {
        aHeader << "    const IID aIID_" << aDefaultInterface.msLibName << "_"
                << aDefaultInterface.msName << " = " << IID_initializer(aDefaultInterface.maIID)
                << ";\n";
        aHeader << "    if (IsEqualIID(rIID, aIID_" << aDefaultInterface.msLibName << "_"
                << aDefaultInterface.msName << "))\n";
        aHeader << "    {\n";
        aHeader << "        *pPDispatch = reinterpret_cast<IDispatch*>(new C"
                << aDefaultInterface.msLibName << "_" << aDefaultInterface.msName
                << "(pBaseClassUnknown, pReplacementAppDispatch));\n";
        aHeader << "        sClass = \"" << aDefaultInterface.msLibName << "_"
                << aDefaultInterface.msName << "\";\n";
        aHeader << "        return true;\n";
        aHeader << "    }\n";
    }

    aHeader << "    return false;\n";
    aHeader << "}\n";

    aHeader << "\n";
    aHeader << "#endif // INCLUDED_DefaultInterfaceCreator_HXX\n";

    aHeader.close();
}

static void GenerateOutgoingInterfaceMap()
{
    if (aOutgoingInterfaceMap.size() == 0)
        return;

    const std::string sHeader = sOutputFolder + "/OutgoingInterfaceMap.hxx";
    OutputFile aHeader(sHeader);

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_OutgoingInterfaceMap_HXX\n";
    aHeader << "#define INCLUDED_OutgoingInterfaceMap_HXX\n";
    aHeader << "\n";
    aHeader << "#include \"outgoingmap.hpp\"\n";
    aHeader << "\n";

    int nMappingIndex = 0;
    for (const auto i : aOutgoingInterfaceMap)
    {
        if (i.maNameToId == nullptr)
        {
            std::cerr << "Information was not filled in for all requested outgoing interfaces. Did "
                         "you perhaps leave\n"
                      << "the interface name for " << IID_to_string(i.maSourceInterfaceInProxiedApp)
                      << " out from the -I file?\n";
            std::exit(1);
        }

        aHeader << "const static NameToMemberIdMapping a" << nMappingIndex << "[] =\n";
        aHeader << "{\n";
        for (const NameToMemberIdMapping* p = i.maNameToId; p->mpName != nullptr; ++p)
        {
            aHeader << "    { \"" << p->mpName << "\", " << p->mnMemberId << " },\n";
        }
        aHeader << "    { nullptr, 0 }\n";
        aHeader << "};\n";
        aHeader << "\n";

        ++nMappingIndex;
    }

    aHeader << "const static OutgoingInterfaceMapping "
               "aOutgoingInterfaceMap[] =\n";
    aHeader << "{\n";

    nMappingIndex = 0;
    for (const auto i : aOutgoingInterfaceMap)
    {
        aHeader << "    { " << IID_initializer(i.maSourceInterfaceInProxiedApp) << ", "
                << IID_initializer(i.maOutgoingInterfaceInReplacement) << ", "
                << "a" << nMappingIndex << " },\n";
        ++nMappingIndex;
    }

    aHeader << "};\n";
    aHeader << "\n";
    aHeader << "#endif // INCLUDED__OutgoingInterfaceMap_HXX\n";

    aHeader.close();
}

static void GenerateProxyCreator()
{
    if (aDispatches.size() == 0)
        return;

    const std::string sHeader = sOutputFolder + "/ProxyCreator.hxx";
    OutputFile aHeader(sHeader);

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_ProxyCreator_HXX\n";
    aHeader << "#define INCLUDED_ProxyCreator_HXX\n";
    aHeader << "\n";

    for (const auto i : aDispatches)
    {
        aHeader << "#include \"C" << i.msLibName << "_" << i.msName << ".hxx\"\n";
    }

    aHeader << "\n";

    aHeader << "static IDispatch* "
            << "ProxyCreator(const IID& aIID, IDispatch* pDispatchToProxy)\n";
    aHeader << "{\n";

    for (const auto i : aDispatches)
    {
        aHeader << "    const IID aIID_" << i.msLibName << "_" << i.msName << " = "
                << IID_initializer(i.maIID) << ";\n";
        aHeader << "    if (IsEqualIID(aIID, aIID_" << i.msLibName << "_" << i.msName << "))\n";
        aHeader << "        return reinterpret_cast<IDispatch*>(C" << i.msLibName << "_" << i.msName
                << "::get(nullptr, pDispatchToProxy));\n";
    }

    aHeader << "    return pDispatchToProxy;\n";
    aHeader << "};\n";
    aHeader << "\n";
    aHeader << "#endif // INCLUDED_ProxyCreator_HXX\n";

    aHeader.close();
}

static void GenerateInterfaceMapping()
{
    if (aInterfaceMap.size() == 0)
        return;

    const std::string sHeader = sOutputFolder + "/InterfaceMapping.hxx";
    OutputFile aHeader(sHeader);

    aHeader << "// Generated file. Do not edit.\n";
    aHeader << "\n";
    aHeader << "#ifndef INCLUDED_InterfaceMapping_HXX\n";
    aHeader << "#define INCLUDED_InterfaceMapping_HXX\n";
    aHeader << "\n";
    aHeader << "#include \"interfacemap.hpp\"\n";
    aHeader << "\n";

    aHeader << "const static InterfaceMapping aInterfaceMap[] =\n";
    aHeader << "{\n";

    for (const auto i : aInterfaceMap)
    {
        aHeader << "    { " << IID_initializer(i.maFromCoclass) << ",\n"
                << "      " << IID_initializer(i.maFromDefault) << ",\n"
                << "      " << IID_initializer(i.maReplacementCoclass) << ",\n"
                << "      \"" << i.msFromLibName << "\",\n"
                << "      \"" << i.msFromCoclassName << "\" },\n";
    }

    aHeader << "};\n";
    aHeader << "\n";
    aHeader << "#endif // INCLUDED_InterfaceMapping_HXX\n";

    aHeader.close();
}

static void Usage(wchar_t** argv)
{
    std::cerr << "Usage: " << convertUTF16ToUTF8(programName(argv[0]))
              << " [options] typelibrary[:interface] ...\n"
                 "\n"
                 "  Options:\n"
                 "    -d directory                 Directory where to write generated files.\n"
                 "                                 Default: \"generated\".\n"
                 "    -i app.iface,app.iface,...   Only output code for those interfaces\n"
                 "    -I file                      Only output code for interfaces listed in file\n"
                 "    -M file                      file contains interface mappings\n"
                 "    -O file                      Use mapping between replacement app's outgoing "
                 "interface IIDs\n"
                 "                                 and the proxied application's source interface "
                 "IIDs in file\n"
                 "  If no -M option is given, does not do any COM server redirection.\n"
                 "  For instance: "
              << convertUTF16ToUTF8(programName(argv[0]))
              << " -i _Application,Documents,Document foo.olb:Application bar.exe:SomeInterface\n";
    std::exit(1);
}

static bool parseMapping(const char* pLine, InterfaceMapping& rMapping)
{
    wchar_t* pWLine = _wcsdup(convertUTF8ToUTF16(pLine).data());

    wchar_t* pColon1 = std::wcschr(pWLine, L':');
    if (!pColon1)
        return false;
    *pColon1 = L'\0';

    wchar_t* pColon2 = std::wcschr(pColon1 + 1, L':');
    if (!pColon2)
        return false;
    *pColon2 = L'\0';

    if (FAILED(IIDFromString(pWLine, &rMapping.maFromCoclass))
        || FAILED(IIDFromString(pColon1 + 1, &rMapping.maFromDefault))
        || FAILED(IIDFromString(pColon2 + 1, &rMapping.maReplacementCoclass)))
        return false;

    rMapping.msFromLibName = "";
    rMapping.msFromCoclassName = "";

    return true;
}

int wmain(int argc, wchar_t** argv)
{
    int argi = 1;

    while (argi < argc && argv[argi][0] == L'-')
    {
        switch (argv[argi][1])
        {
            case L'd':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                sOutputFolder = convertUTF16ToUTF8(argv[argi + 1]);
                argi++;
                break;
            }
            case L'i':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                wchar_t* p = argv[argi + 1];
                wchar_t* q;
                while ((q = std::wcschr(p, L',')) != NULL)
                {
                    aOnlyTheseInterfaces.insert(
                        std::string(convertUTF16ToUTF8(p), (unsigned)(q - p)));
                    p = q + 1;
                }
                aOnlyTheseInterfaces.insert(std::string(convertUTF16ToUTF8(p)));
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
                    std::cerr << "Could not open " << convertUTF16ToUTF8(argv[argi + 1])
                              << " for reading\n";
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
                    std::cerr << "Could not open " << convertUTF16ToUTF8(argv[argi + 1])
                              << " for reading\n";
                    std::exit(1);
                }
                int nLine = 0;
                while (!aMFile.eof())
                {
                    std::string sLine;
                    std::getline(aMFile, sLine);
                    nLine++;
                    if (sLine.length() == 0 || sLine[0] == '#')
                        continue;
                    char* pLine = _strdup(sLine.data());
                    InterfaceMapping aMapping;
                    if (!parseMapping(pLine, aMapping))
                    {
                        std::cerr << "Invalid IIDs on line " << nLine << " in "
                                  << convertUTF16ToUTF8(argv[argi + 1]) << "\n";
                        Usage(argv);
                    }
                    aInterfaceMap.push_back(aMapping);
                    std::free(pLine);
                }
                aMFile.close();
                argi++;
                break;
            }
            case 'O':
            {
                if (argi + 1 >= argc)
                    Usage(argv);
                std::ifstream aOFile(argv[argi + 1]);
                if (!aOFile.good())
                {
                    std::cerr << "Could not open " << convertUTF16ToUTF8(argv[argi + 1])
                              << " for reading\n";
                    std::exit(1);
                }
                while (!aOFile.eof())
                {
                    std::string sProxiedAppSourceInterface;
                    aOFile >> sProxiedAppSourceInterface;
                    if (sProxiedAppSourceInterface.length() > 0
                        && sProxiedAppSourceInterface[0] == '#')
                    {
                        std::string sRestOfLine;
                        std::getline(aOFile, sRestOfLine);
                        continue;
                    }
                    std::string sReplacementOutgoingInterface;
                    aOFile >> sReplacementOutgoingInterface;
                    if (aOFile.good())
                    {
                        IID aProxiedAppSourceInterface;
                        if (FAILED(IIDFromString(
                                convertUTF8ToUTF16(sProxiedAppSourceInterface.c_str()).data(),
                                &aProxiedAppSourceInterface)))
                        {
                            std::cerr << "Count not interpret " << sProxiedAppSourceInterface
                                      << " in '" << convertUTF16ToUTF8(argv[argi + 1])
                                      << "' as an IID\n";
                            break;
                        }
                        IID aReplacementOutgoingInterface;
                        if (FAILED(IIDFromString(
                                convertUTF8ToUTF16(sReplacementOutgoingInterface.c_str()).data(),
                                &aReplacementOutgoingInterface)))
                        {
                            std::cerr << "Could not interpret " << sReplacementOutgoingInterface
                                      << " in " << convertUTF16ToUTF8(argv[argi + 1])
                                      << " as an IID\n";
                            break;
                        }
                        // We leave the maNameToId as null and fill in once we have the type library.
                        aOutgoingInterfaceMap.push_back(
                            { aProxiedAppSourceInterface, aReplacementOutgoingInterface, nullptr });
                    }
                }
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

    CoInitialize(NULL);

    for (; argi < argc; ++argi)
    {
        wchar_t* const pColon = std::wcschr(argv[argi], L':');
        const wchar_t* pInterface = L"Application";
        if (pColon)
        {
            *pColon = L'\0';
            pInterface = pColon + 1;
        }

        HRESULT nResult;
        ITypeLib* pTypeLib;
        nResult = LoadTypeLibEx(argv[argi], REGKIND_NONE, &pTypeLib);
        if (FAILED(nResult))
        {
            std::cerr << "Could not load '" << convertUTF16ToUTF8(argv[argi])
                      << "' as a type library: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            std::exit(1);
        }

        UINT nTypeInfoCount = pTypeLib->GetTypeInfoCount();
        if (nTypeInfoCount == 0)
        {
            std::cerr << "No type information in '" << convertUTF16ToUTF8(argv[argi]) << "'?\n";
            std::exit(1);
        }

        BSTR sLibNameBstr;
        nResult = pTypeLib->GetDocumentation(-1, &sLibNameBstr, NULL, NULL, NULL);
        if (FAILED(nResult))
        {
            std::cerr << "GetDocumentation(-1) of '" << argv[argi]
                      << "' failed: " << WindowsErrorStringFromHRESULT(nResult) << "\n";
            std::exit(1);
        }
        std::string sLibName = convertUTF16ToUTF8(sLibNameBstr);

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

            if (std::wcscmp(sTypeName, pInterface) != 0)
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

    GenerateProxyCreator();

    GenerateInterfaceMapping();

    GenerateOutgoingInterfaceMap();

    GenerateCallbackInvoker();

    GenerateDefaultInterfaceCreator();

    CoUninitialize();

    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

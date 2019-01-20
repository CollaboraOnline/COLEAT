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
#include <iostream>
#include <string>
#include <vector>

#include <Windows.h>
#include <OleCtl.h>

#pragma warning(pop)

#include "utils.hpp"

#include "CProxiedDispatch.hpp"
#include "CProxiedEnumVARIANT.hpp"

#include "ProxyCreator.hxx"

CProxiedDispatch::CProxiedDispatch(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                   const char* sLibName, const char* sPropName)
    : CProxiedDispatch(pBaseClassUnknown, pDispatchToProxy, IID_NULL, sLibName, sPropName)
{
}

CProxiedDispatch::CProxiedDispatch(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                   const IID& rIID, const char* sLibName, const char* sPropName)
    : CProxiedDispatch(pBaseClassUnknown, pDispatchToProxy, IID_IDispatch, rIID, sLibName,
                       sPropName)
{
}

CProxiedDispatch::CProxiedDispatch(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                   const IID& rIID1, const IID& rIID2, const char* sLibName,
                                   const char* sPropName)
    : CProxiedUnknown(pBaseClassUnknown, pDispatchToProxy, rIID1, rIID2, sLibName)
    , mpDispatchToProxy(pDispatchToProxy)
    , mpDispIdToName(new std::map<DISPID, std::map<DISPID, std::string>>)
    , msPropName(sPropName)
{
    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::CTOR(" << pBaseClassUnknown << ", "
                  << pDispatchToProxy << ", " << rIID1 << ", " << rIID2 << ")" << std::endl;
}

CProxiedDispatch* CProxiedDispatch::get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                        const char* sLibName, const char* sPropName)
{
    return get(pBaseClassUnknown, pDispatchToProxy, IID_NULL, sLibName, sPropName);
}

CProxiedDispatch* CProxiedDispatch::get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                        const IID& rIID, const char* sLibName,
                                        const char* sPropName)
{
    CProxiedUnknown* pExisting = find(pDispatchToProxy);
    if (pExisting != nullptr)
        return static_cast<CProxiedDispatch*>(pExisting);

    return new CProxiedDispatch(pBaseClassUnknown, pDispatchToProxy, rIID, sLibName, sPropName);
}

CProxiedDispatch* CProxiedDispatch::get(IUnknown* pBaseClassUnknown, IDispatch* pDispatchToProxy,
                                        const IID& rIID1, const IID& rIID2, const char* sLibName,
                                        const char* sPropName)
{
    return new CProxiedDispatch(pBaseClassUnknown, pDispatchToProxy, rIID1, rIID2, sLibName,
                                sPropName);
}

HRESULT CProxiedDispatch::genericInvoke(const std::string& rFuncName, int nInvKind,
                                        std::vector<VARIANT>& rParameters, void* pRetval)
{
    if (getParam()->mbVerbose)
    {
        std::cout << std::endl;
        std::cout << this << "@CProxiedDispatch::genericInvoke(" << rFuncName << ")..."
                  << std::endl;
    }

    HRESULT nResult = S_OK;

    std::wstring sFuncNameWstr = convertUTF8ToUTF16(rFuncName.c_str());
    LPOLESTR pFuncNameWstr = (LPOLESTR)sFuncNameWstr.data();

    MEMBERID nMemberId;

    nResult = mpDispatchToProxy->GetIDsOfNames(IID_NULL, &pFuncNameWstr, 1, LOCALE_USER_DEFAULT,
                                               &nMemberId);
    if (nResult == DISP_E_UNKNOWNNAME)
    {
        if (getParam()->mbVerbose)
            std::cout << this << "@CProxiedDispatch::genericInvoke(" << rFuncName
                      << "): Not implemented in the replacement app" << std::endl;
        return E_NOTIMPL;
    }

    if (FAILED(nResult))
    {
        if (getParam()->mbVerbose)
            std::cout << this << "@CProxiedDispatch::genericInvoke(" << rFuncName
                      << "): GetIDsOfNames failed: " << WindowsErrorStringFromHRESULT(nResult)
                      << std::endl;
        return nResult;
    }

    WORD nFlags;
    switch (nInvKind)
    {
        case INVOKE_FUNC:
            nFlags = DISPATCH_METHOD;
            break;
        case INVOKE_PROPERTYGET:
            nFlags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;
            break;
        case INVOKE_PROPERTYPUT:
            nFlags = DISPATCH_PROPERTYPUT;
            break;
        case INVOKE_PROPERTYPUTREF:
            nFlags = DISPATCH_PROPERTYPUTREF;
            break;
        default:
            std::cout << this << "@CProxiedDispatch::genericInvoke(" << rFuncName
                      << "): Unhandled nInvKind: " << nInvKind << std::endl;
            std::abort();
    }

    DISPPARAMS aDispParams;
    aDispParams.rgvarg = rParameters.data();
    aDispParams.rgdispidNamedArgs = NULL;
    aDispParams.cArgs = (UINT)rParameters.size();
    aDispParams.cNamedArgs = 0;

    VARIANT aResult;
    VariantInit(&aResult);

    UINT nArgErr;

    // See https://msdn.microsoft.com/en-us/library/windows/desktop/ms221486(v=vs.85).aspx
    DISPID nDispid = DISPID_PROPERTYPUT;
    if (nFlags == DISPATCH_PROPERTYPUT)
    {
        aDispParams.rgdispidNamedArgs = &nDispid;
        aDispParams.cNamedArgs = 1;
    }

    nResult = mpDispatchToProxy->Invoke(nMemberId, IID_NULL, LOCALE_USER_DEFAULT, nFlags,
                                        &aDispParams, &aResult, NULL, &nArgErr);
    if (FAILED(nResult))
    {
        if (getParam()->mbVerbose)
            std::cout << "..." << this << "@CProxiedDispatch::genericInvoke(" << rFuncName
                      << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }

    if (aResult.vt != VT_EMPTY && aResult.vt != VT_VOID && pRetval != nullptr)
    {
        switch (aResult.vt)
        {
            case VT_I2:
                *(int16_t*)pRetval = aResult.iVal;
                break;
            case VT_I4:
                *(int32_t*)pRetval = aResult.lVal;
                break;
            case VT_R4:
                *(float*)pRetval = aResult.fltVal;
                break;
            case VT_R8:
                *(double*)pRetval = aResult.dblVal;
                break;
            case VT_BSTR:
                *(BSTR*)pRetval = aResult.bstrVal;
                break;
            case VT_DISPATCH:
                *(IDispatch**)pRetval = aResult.pdispVal;
                break;
            default:
                std::cout << this << "@CProxiedDispatch::genericInvoke(" << rFuncName
                          << "): Unhandled vt: " << aResult.vt << std::endl;
                std::abort();
        }
    }
    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedDispatch::genericInvoke(" << rFuncName << "): S_OK"
                  << std::endl;

    return S_OK;
}

// IDispatch

HRESULT STDMETHODCALLTYPE CProxiedDispatch::GetTypeInfoCount(UINT* pctinfo)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::GetTypeInfoCount..." << std::endl;

    nResult = mpDispatchToProxy->GetTypeInfoCount(pctinfo);

    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedDispatch::GetTypeInfoCount:"
                  << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedDispatch::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::GetTypeInfo(" << iTInfo << ")..." << std::endl;

    nResult = mpDispatchToProxy->GetTypeInfo(iTInfo, lcid, ppTInfo);

    if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedDispatch::GetTypeInfo(" << iTInfo
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedDispatch::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames,
                                                          UINT cNames, LCID lcid, DISPID* rgDispId)
{
    HRESULT nResult;

    if (getParam()->mbVerbose)
    {
        std::cout << this << "@CProxiedDispatch::GetIDsOfNames(" << cNames << ":[";
        for (UINT i = 0; i < cNames; ++i)
        {
            if (i > 0)
                std::cout << ",";
            std::cout << "\"" << convertUTF16ToUTF8(rgszNames[i]) << "\"";
        }
        std::cout << "])..." << std::endl;
    }
    nResult = mpDispatchToProxy->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);

    if (rgszNames && rgDispId && nResult == S_OK && getParam()->mbTrace)
    {
        std::string sName = convertUTF16ToUTF8(rgszNames[0]);
        if (mpDispIdToName->count(rgDispId[0]))
            assert((*mpDispIdToName)[rgDispId[0]][0] == sName);
        else
            (*mpDispIdToName)[rgDispId[0]][0] = sName;
    }

    if (getParam()->mbVerbose)
    {
        std::cout << "..." << this << "@CProxiedDispatch::GetIDsOfNames: ";
        if (nResult == S_OK)
        {
            std::cout << "[";
            for (UINT i = 0; i < cNames; ++i)
            {
                if (i > 0)
                    std::cout << ",";
                std::cout << "\"" << convertUTF16ToUTF8(rgszNames[i]) << "\":0x"
                          << to_hex(rgDispId[i]);
                if (i > 0)
                    (*mpDispIdToName)[rgDispId[0]][rgDispId[i]] = convertUTF16ToUTF8(rgszNames[i]);
            }
            std::cout << "]" << std::endl;
        }
        else
            std::cout << WindowsErrorStringFromHRESULT(nResult) << std::endl;
    }
    else if (getParam()->mbTrace)
    {
        if (nResult == S_OK)
        {
            for (UINT i = 1; i < cNames; ++i)
            {
                (*mpDispIdToName)[rgDispId[0]][rgDispId[i]] = convertUTF16ToUTF8(rgszNames[i]);
            }
        }
    }


    return nResult;
}

HRESULT STDMETHODCALLTYPE CProxiedDispatch::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
                                                   WORD wFlags, DISPPARAMS* pDispParams,
                                                   VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                                                   UINT* puArgErr)
{
    HRESULT nResult;

    ITypeInfo* pTI = NULL;
    FUNCDESC* pFuncDesc = NULL;

    nResult = mpDispatchToProxy->GetTypeInfo(0, LOCALE_USER_DEFAULT, &pTI);

    TYPEATTR* pTA = NULL;
    if (!FAILED(nResult))
        nResult = pTI->GetTypeAttr(&pTA);

    if (!FAILED(nResult))
    {
        for (UINT i = 0; i < pTA->cFuncs; ++i)
        {
            nResult = pTI->GetFuncDesc(i, &pFuncDesc);
            if (FAILED(nResult))
                break;

            if (pFuncDesc->memid == dispIdMember
                && (((wFlags & DISPATCH_METHOD) && pFuncDesc->invkind == INVOKE_FUNC)
                    || ((wFlags & DISPATCH_PROPERTYGET) && pFuncDesc->invkind == INVOKE_PROPERTYGET)
                    || ((wFlags & DISPATCH_PROPERTYPUT) && pFuncDesc->invkind == INVOKE_PROPERTYPUT)
                    || ((wFlags & DISPATCH_PROPERTYPUTREF)
                        && pFuncDesc->invkind == INVOKE_PROPERTYPUTREF)))
                break;
            pTI->ReleaseFuncDesc(pFuncDesc);
            pFuncDesc = NULL;
        }
    }

    if (getParam()->mbTrace)
    {
        std::cout << msLibName << "." << std::flush;

        if (pTI == NULL)
            if (msPropName != nullptr)
                std::cout << msPropName;
            else
                std::cout << "?";
        else
        {
            BSTR sTypeName = NULL;
            nResult = pTI->GetDocumentation(MEMBERID_NIL, &sTypeName, NULL, NULL, NULL);

            if (FAILED(nResult))
                std::cout << "?";
            else
            {
                std::cout << convertUTF16ToUTF8(sTypeName);
                SysFreeString(sTypeName);
            }
        }

        std::cout << "<" << (mpBaseClassUnknown ? mpBaseClassUnknown : this) << ">.";

        if (pTI == NULL)
            if (mpDispIdToName->count(dispIdMember))
                std::cout << (*mpDispIdToName)[dispIdMember][0];
            else
                std::cout << dispIdMember;
        else
        {
            BSTR sFuncName = NULL;
            nResult = pTI->GetDocumentation(dispIdMember, &sFuncName, NULL, NULL, NULL);

            if (FAILED(nResult))
                std::cout << "?";
            else
            {
                std::cout << convertUTF16ToUTF8(sFuncName);
                SysFreeString(sFuncName);
            }
        }

        // Print in and inout parameters. If we don't have type information, we don't know which
        // ones are just out. If this is a property assignment, and there is just one parameter,
        // that is the new property value, and we print the call like a pseudo-code assignment
        // below.

        if (pFuncDesc != NULL
            && !(pFuncDesc->invkind == INVOKE_PROPERTYGET && pDispParams->cArgs == 0)
            && !((pFuncDesc->invkind == INVOKE_PROPERTYPUT
                  || pFuncDesc->invkind == INVOKE_PROPERTYPUTREF)
                 && pDispParams->cArgs == 1))
        {
            std::cout << "(";
            for (UINT n = 0; n < pDispParams->cArgs; ++n)
            {
                if (n > 0)
                    std::cout << ",";
                if (pFuncDesc != NULL && (SHORT)n < pFuncDesc->cParams
                    && !(pFuncDesc->lprgelemdescParam[n].paramdesc.wParamFlags & PARAMFLAG_FIN))
                    std::cout << "<OUT>";
                else
                    std::cout << pDispParams->rgvarg[n];
            }
            std::cout << ")";
        }
        else if (pFuncDesc == NULL)
        {
            std::cout << "(";
            for (UINT n = 0; n < pDispParams->cArgs; ++n)
            {
                if (n > 0)
                    std::cout << ",";
                if (n < pDispParams->cNamedArgs && pTI == NULL)
                    if (mpDispIdToName->count(dispIdMember) &&
                        (*mpDispIdToName)[dispIdMember].count(pDispParams->rgdispidNamedArgs[n]))
                        std::cout << (*mpDispIdToName)[dispIdMember][pDispParams->rgdispidNamedArgs[n]] << ":=";

                std::cout << pDispParams->rgvarg[n];
            }
            std::cout << ")";
        }

        if (getParam()->mbVerbose)
            std::cout << std::endl;
        else
            mbIsAtBeginningOfLine = false;
    }
    else if (getParam()->mbVerbose)
        std::cout << this << "@CProxiedDispatch::Invoke(0x" << to_hex(dispIdMember) << ")..."
                  << std::endl;

    increaseIndent();
    nResult = mpDispatchToProxy->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult,
                                        pExcepInfo, puArgErr);
    decreaseIndent();

    std::string sPrettyResultTypeName;

    if (nResult == S_OK && pFuncDesc != NULL)
    {
        if (pFuncDesc->invkind == INVOKE_FUNC || pFuncDesc->invkind == INVOKE_PROPERTYGET)
        {
            if (pVarResult != NULL && pVarResult->vt == VT_DISPATCH
                && pFuncDesc->elemdescFunc.tdesc.vt == VT_PTR)
            {
                pVarResult->pdispVal = ProxyCreator(pVarResult->pdispVal, sPrettyResultTypeName);
            }
            else if (pVarResult != NULL && pVarResult->vt == VT_DISPATCH
                     && pFuncDesc->elemdescFunc.tdesc.vt == VT_DISPATCH)
            {
                IDispatch* pNewDispVal = ProxyCreator(pVarResult->pdispVal, sPrettyResultTypeName);
                if (pNewDispVal != pVarResult->pdispVal)
                    pVarResult->pdispVal = pNewDispVal;
                else
                    pVarResult->pdispVal = reinterpret_cast<IDispatch*>(
                        CProxiedDispatch::get(nullptr, pVarResult->pdispVal, "Unknown"));
            }
            else if (dispIdMember == DISPID_NEWENUM && pVarResult != NULL
                     && pVarResult->vt == VT_UNKNOWN
                     && pFuncDesc->elemdescFunc.tdesc.vt == VT_UNKNOWN)
            {
                pVarResult->punkVal = new CProxiedEnumVARIANT(pVarResult->punkVal, msLibName);
            }
        }
    }

    if (nResult == S_OK && getParam()->mbTrace)
    {
        // FIXME: Print inout and out parameters here.

        if (pFuncDesc != NULL)
        {
            if (pFuncDesc->invkind == INVOKE_FUNC || pFuncDesc->invkind == INVOKE_PROPERTYGET)
            {
                if (pVarResult != NULL)
                {
                    std::cout << " -> ";
                    if (sPrettyResultTypeName != "")
                        std::cout << sPrettyResultTypeName << "<" << *pVarResult << ">";
                    else
                        std::cout << *pVarResult;
                    mbIsAtBeginningOfLine = false;
                }

                if (!mbIsAtBeginningOfLine)
                {
                    std::cout << std::endl;
                    mbIsAtBeginningOfLine = true;
                }
            }
            else if (pFuncDesc->invkind == INVOKE_PROPERTYPUT
                     || pFuncDesc->invkind == INVOKE_PROPERTYPUTREF)
            {
                // FIXME: Why do we print the value being put here and not above, before the
                // mpDispatchToProxy->Invoke() call?
                if (pDispParams->cArgs > 0)
                {
                    std::cout << " = " << pDispParams->rgvarg[pDispParams->cArgs - 1];
                    mbIsAtBeginningOfLine = false;
                }

                if (!mbIsAtBeginningOfLine)
                {
                    std::cout << std::endl;
                    mbIsAtBeginningOfLine = true;
                }
            }
        }
        else
        {
            if (pVarResult != NULL)
            {
                std::cout << " -> " << *pVarResult;
                mbIsAtBeginningOfLine = false;
            }

            if (!mbIsAtBeginningOfLine)
            {
                std::cout << std::endl;
                mbIsAtBeginningOfLine = true;
            }
        }

        if (pFuncDesc != NULL)
            pTI->ReleaseFuncDesc(pFuncDesc);
    }
    else if (getParam()->mbTrace)
    {
        std::cout << ": " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        mbIsAtBeginningOfLine = true;
    }
    else if (getParam()->mbVerbose)
        std::cout << "..." << this << "@CProxiedDispatch::Invoke(0x" << to_hex(dispIdMember)
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

    return nResult;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

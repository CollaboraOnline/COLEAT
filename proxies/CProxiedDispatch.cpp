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

#include <codecvt>
#include <cstdlib>
#include <iostream>
#include <map>
#include <sstream>
#include <string>
#include <vector>

#include <Windows.h>
#include <OleCtl.h>

#include <comphelper/windowsdebugoutput.hxx>

#pragma warning(pop)

#include "CProxiedDispatch.hpp"
#include "CallbackInvoker.hxx"
#include "OutgoingInterfaceMap.hxx"

static ThreadProcParam* pGlobalParam;

// FIXME: Factor out these three functions into some include file. Now copy-pasted from
// genproxy.cpp. (Originally from LibreOffice's <comphelper/windowserrorstring.hxx>, but that uses
// OUString and can't thus be used as such outside LibreOffice.)

static std::string to_hex(uint64_t n)
{
    std::stringstream aStringStream;
    aStringStream << std::hex << n;
    return aStringStream.str();
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

    std::string sResult
        = std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t>().to_bytes(pMsgBuf);
    if (sResult.length() > 2 && sResult.substr(sResult.length() - 2, 2) == "\r\n")
        sResult.resize(sResult.length() - 2);

    HeapFree(GetProcessHeap(), 0, pMsgBuf);

    return sResult;
}

static std::string WindowsErrorStringFromHRESULT(HRESULT nResult)
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

class CProxiedSink : public IDispatch
{
private:
    IDispatch* mpDispatchToProxy;
    ITypeInfo* mpTypeInfoOfOutgoingInterface;
    IID maIID;

public:
    CProxiedSink(IDispatch* pDispatchToProxy, ITypeInfo* pTypeInfoOfOutgoingInterface, IID aIID)
        : mpDispatchToProxy(pDispatchToProxy)
        , mpTypeInfoOfOutgoingInterface(pTypeInfoOfOutgoingInterface)
        , maIID(aIID)
    {
        std::cout << this << "@CProxiedSink::CTOR" << std::endl;
    }

    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (!ppvObject)
            return E_POINTER;

        if (IsEqualIID(riid, IID_IUnknown))
        {
            *ppvObject = this;
            AddRef();
            return S_OK;
        }

        if (IsEqualIID(riid, IID_IDispatch))
        {
            *ppvObject = this;
            AddRef();
            return S_OK;
        }

        std::cout << this << "@CProxiedSink::QueryInterface(" << riid << "): E_NOINTERFACE"
                  << std::endl;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        ULONG nRetval = mpDispatchToProxy->AddRef();
        std::cout << this << "@CProxiedSink::AddRef: " << nRetval << "" << std::endl;

        return nRetval;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG nRetval = mpDispatchToProxy->Release();
        std::cout << this << "@CProxiedSink::Release: " << nRetval << "" << std::endl;

        return nRetval;
    }

    // IDispatch
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override
    {
        if (!pctinfo)
            return E_POINTER;

        HRESULT nResult = mpDispatchToProxy->GetTypeInfoCount(pctinfo);
        *pctinfo = 0;
        std::cout << this << "@CProxiedSink::GetTypeInfoCount: " << *pctinfo << ": "
                  << WindowsErrorStringFromHRESULT(nResult) << "" << std::endl;
        return nResult;
    }

    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override
    {
        HRESULT nResult = mpDispatchToProxy->GetTypeInfo(iTInfo, lcid, ppTInfo);
        std::cout << this
                  << "@CProxiedSink::GetTypeInfo: " << WindowsErrorStringFromHRESULT(nResult) << ""
                  << std::endl;
        return nResult;
    }

    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
                                            LCID lcid, DISPID* rgDispId) override
    {
        HRESULT nResult = mpDispatchToProxy->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
        std::cout << this
                  << "@CProxiedSink::GetIDsOfNames: " << WindowsErrorStringFromHRESULT(nResult)
                  << "" << std::endl;
        return nResult;
    }

    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                                     DISPPARAMS* pDispParams, VARIANT* pVarResult,
                                     EXCEPINFO* pExcepInfo, UINT* puArgErr) override
    {
        HRESULT nResult;

        std::cout << this << "@CProxiedSink::Invoke(" << dispIdMember << ")..." << std::endl;

        BSTR sName;
        UINT nNames;
        nResult = mpTypeInfoOfOutgoingInterface->GetNames(dispIdMember, &sName, 1, &nNames);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedSink::Invoke(" << dispIdMember
                      << "): GetNames failed: " << WindowsErrorStringFromHRESULT(nResult) << ""
                      << std::endl;
            return nResult;
        }

        DISPID nDispIdMemberInClient;
        nResult
            = mpDispatchToProxy->GetIDsOfNames(IID_NULL, &sName, 1, lcid, &nDispIdMemberInClient);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedSink::Invoke(" << dispIdMember
                      << "): GetIDsOfNames("
                      << std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t>().to_bytes(sName)
                      << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << "" << std::endl;
            SysFreeString(sName);
            return nResult;
        }

        SysFreeString(sName);

        nResult = ProxiedCallbackInvoke(maIID, mpDispatchToProxy, nDispIdMemberInClient, riid, lcid,
                                        wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

        std::cout << "..." << this << "@CProxiedSink::Invoke(" << dispIdMember
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << "" << std::endl;

        return nResult;
    }
};

class CProxiedEnumConnections : public IEnumConnections
{
private:
    IEnumConnections* mpEnumConnnectionsToProxy;

public:
    CProxiedEnumConnections(IEnumConnections* pEnumConnnectionsToProxy)
        : mpEnumConnnectionsToProxy(pEnumConnnectionsToProxy)
    {
        std::cout << this << "@CProxiedEnumConnections::CTOR" << std::endl;
    }

    // IUnknown

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (IsEqualIID(riid, IID_IUnknown))
        {
            *ppvObject = this;
            AddRef();
            return S_OK;
        }

        if (IsEqualIID(riid, IID_IEnumConnections))
        {
            *ppvObject = this;
            AddRef();
            return S_OK;
        }

        std::cout << this << "@CProxiedEnumConnections::QueryInterface(" << riid
                  << "): E_NOINTERFACE" << std::endl;

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        ULONG nRetval = mpEnumConnnectionsToProxy->AddRef();

        std::cout << this << "@CProxiedEnumConnections::AddRef: " << nRetval << "" << std::endl;

        return nRetval;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG nRetval = mpEnumConnnectionsToProxy->Release();

        std::cout << this << "@CProxiedEnumConnections::Release: " << nRetval << "" << std::endl;

        return nRetval;
    }

    // IEnumConnections

    HRESULT STDMETHODCALLTYPE Next(ULONG cConnections, LPCONNECTDATA rgcd,
                                   ULONG* pcFetched) override
    {
        HRESULT nResult;

        std::cout << this << "@CProxiedEnumConnections::Next(" << cConnections << ")..."
                  << std::endl;

        nResult = mpEnumConnnectionsToProxy->Next(cConnections, rgcd, pcFetched);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedEnumConnections::Next(" << cConnections
                      << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
            return nResult;
        }
        std::cout << "..." << this << "@CProxiedEnumConnections::Next(" << cConnections
                  << "): " << WindowsErrorStringFromHRESULT(nResult);
        if (pcFetched)
            std::cout << *pcFetched;
        std::cout << std::endl;

        return nResult;
    }

    HRESULT STDMETHODCALLTYPE Skip(ULONG cConnections) override
    {
        HRESULT nResult;

        std::cout << this << "@CProxiedEnumConnections::Skip(" << cConnections << ")..."
                  << std::endl;

        nResult = mpEnumConnnectionsToProxy->Skip(cConnections);

        std::cout << "..." << this << "@CProxiedEnumConnections::Skip(" << cConnections
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;

        return nResult;
    }

    HRESULT STDMETHODCALLTYPE Reset() override
    {
        HRESULT nResult;

        std::cout << this << "@CProxiedEnumConnections::Reset..." << std::endl;

        nResult = mpEnumConnnectionsToProxy->Reset();

        std::cout << "..." << this
                  << "@CProxiedEnumConnections::Reset: " << WindowsErrorStringFromHRESULT(nResult)
                  << std::endl;

        return nResult;
    }

    HRESULT STDMETHODCALLTYPE Clone(IEnumConnections** /* ppEnum */) override
    {
        std::cout << this << "@CProxiedEnumConnections::CLone: E_NOTIMPL" << std::endl;

        return E_NOTIMPL;
    }
};

class CProxiedConnectionPoint : public IConnectionPoint
{
private:
    IConnectionPointContainer* mpContainer;
    IConnectionPoint* mpConnectionPointToProxy;
    IID maIID;
    ITypeInfo* mpTypeInfoOfOutgoingInterface;
    std::map<DWORD, IDispatch*> maAdvisedSinks;

public:
    CProxiedConnectionPoint(IConnectionPointContainer* pContainer,
                            IConnectionPoint* pConnectionPointToProxy, IID aIID,
                            ITypeInfo* pTypeInfoOfOutgoingInterface)
        : mpContainer(pContainer)
        , mpConnectionPointToProxy(pConnectionPointToProxy)
        , maIID(aIID)
        , mpTypeInfoOfOutgoingInterface(pTypeInfoOfOutgoingInterface)
    {
        std::cout << this << "@CProxiedConnectionPoint::CTOR" << std::endl;
    }

    virtual ~CProxiedConnectionPoint() = default;

    // IUnknown

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (IsEqualIID(riid, IID_IUnknown))
        {
            *ppvObject = this;
            AddRef();
            return S_OK;
        }

        if (IsEqualIID(riid, IID_IDispatch))
        {
            *ppvObject = this;
            AddRef();
            return S_OK;
        }

        if (IsEqualIID(riid, IID_IConnectionPoint))
        {
            *ppvObject = this;
            AddRef();
            return S_OK;
        }

        std::cout << this << "@CProxiedConnectionPoint::QueryInterface(" << riid
                  << "): E_NOINTERFACE" << std::endl;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        ULONG nRetval = mpConnectionPointToProxy->AddRef();
        std::cout << this << "@CProxiedConnectionPoint::AddRef: " << nRetval << "" << std::endl;

        return nRetval;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG nRetval = mpConnectionPointToProxy->Release();
        std::cout << this << "@CProxiedConnectionPoint::Release: " << nRetval << "" << std::endl;

        return nRetval;
    }

    // IConnectionPoint

    HRESULT STDMETHODCALLTYPE GetConnectionInterface(IID* /*pIID*/) override { return E_NOTIMPL; }

    HRESULT STDMETHODCALLTYPE
    GetConnectionPointContainer(IConnectionPointContainer** ppCPC) override
    {
        if (!ppCPC)
            return E_POINTER;

        *ppCPC = mpContainer;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Advise(IUnknown* pUnkSink, DWORD* pdwCookie) override
    {
        HRESULT nResult;

        if (!pdwCookie)
            return E_POINTER;

        std::cout << this << "@CProxiedConnectionPoint::Advise(" << pUnkSink << ")..." << std::endl;

        IDispatch* pSinkAsDispatch;
        nResult = pUnkSink->QueryInterface(IID_IDispatch, (void**)&pSinkAsDispatch);
        if (FAILED(nResult))
        {
            std::cerr << "..." << this << "@CProxiedSink::Advise: Sink is not an IDispatch"
                      << std::endl;
            return E_NOTIMPL;
        }

        IDispatch* pDispatch
            = new CProxiedSink(pSinkAsDispatch, mpTypeInfoOfOutgoingInterface, maIID);

        *pdwCookie = 0;
        nResult = mpConnectionPointToProxy->Advise(pDispatch, pdwCookie);

        std::cout << "..." << this << "@CProxiedConnectionPoint::Advise(" << pUnkSink << "): ";

        if (FAILED(nResult))
        {
            std::cout << WindowsErrorStringFromHRESULT(nResult) << "" << std::endl;
            pSinkAsDispatch->Release();
            delete pDispatch;
            return nResult;
        }

        std::cout << *pdwCookie << ": S_OK" << std::endl;

        maAdvisedSinks[*pdwCookie] = pDispatch;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Unadvise(DWORD dwCookie) override
    {
        HRESULT nResult;

        std::cout << this << "@CProxiedConnectionPoint::Unadvise(" << dwCookie << ")..."
                  << std::endl;

        nResult = mpConnectionPointToProxy->Unadvise(dwCookie);
        if (maAdvisedSinks.count(dwCookie) == 0)
        {
            std::cout << "..." << this << "@CProxiedConnectionPoint::Unadvise(" << dwCookie
                      << "): E_POINTER" << std::endl;
            return E_POINTER;
        }
        delete maAdvisedSinks[dwCookie];
        maAdvisedSinks.erase(dwCookie);

        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedConnectionPoint::Unadvise(" << dwCookie
                      << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
            return nResult;
        }

        std::cout << "..." << this << "@CProxiedConnectionPoint::Unadvise(" << dwCookie << "): S_OK"
                  << std::endl;

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE EnumConnections(IEnumConnections** ppEnum) override
    {
        HRESULT nResult;

        if (!ppEnum)
            return E_POINTER;

        std::cout << this << "@CProxiedConnectionPoint::EnumConnections..." << std::endl;

        nResult = mpConnectionPointToProxy->EnumConnections(ppEnum);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedConnectionPoint::EnumConnections: "
                      << WindowsErrorStringFromHRESULT(nResult) << std::endl;
            return nResult;
        }

        *ppEnum = new CProxiedEnumConnections(*ppEnum);

        std::cout << "..." << this << "@CProxiedConnectionPoint::EnumConnections: S_OK"
                  << std::endl;

        return S_OK;
    }
};

class CProxiedEnumConnectionPoints : public IEnumConnectionPoints
{
private:
    CProxiedDispatch* mpProxiedDispatch;
    IEnumConnectionPoints* mpEnumConnectionPointsToProxy;

public:
    CProxiedEnumConnectionPoints(CProxiedDispatch* pProxiedDispatch,
                                 IEnumConnectionPoints* pEnumConnectionPointsToProxy)
        : mpProxiedDispatch(pProxiedDispatch)
        , mpEnumConnectionPointsToProxy(pEnumConnectionPointsToProxy)
    {
        std::cout << this << "@CProxiedEnumConnectionPoints::CTOR" << std::endl;
    }

    // IUnknown

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (IsEqualIID(riid, IID_IUnknown))
        {
            *ppvObject = this;
            AddRef();
            return S_OK;
        }

        if (IsEqualIID(riid, IID_IDispatch))
        {
            *ppvObject = this;
            AddRef();
            return S_OK;
        }

        if (IsEqualIID(riid, IID_IEnumConnectionPoints))
        {
            *ppvObject = this;
            AddRef();
            return S_OK;
        }

        std::cout << this << "@CProxiedEnumConnectionPoints::QueryInterface(" << riid
                  << "): E_NOINTERFACE" << std::endl;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        ULONG nRetval = mpEnumConnectionPointsToProxy->AddRef();
        std::cout << this << "@CProxiedEnumConnectionPoints::AddRef: " << nRetval << ""
                  << std::endl;

        return nRetval;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG nRetval = mpEnumConnectionPointsToProxy->Release();
        std::cout << this << "@CProxiedEnumConnectionPoints::Release: " << nRetval << ""
                  << std::endl;

        return nRetval;
    }

    // IEnumConnectionPoints

    HRESULT STDMETHODCALLTYPE Next(ULONG cConnections, LPCONNECTIONPOINT* ppCP,
                                   ULONG* pcFetched) override
    {
        if (!ppCP || !pcFetched)
            return E_POINTER;

        if (cConnections == 0)
            return E_INVALIDARG;

        std::vector<IConnectionPoint*> vCP(cConnections);
        HRESULT nResult = mpEnumConnectionPointsToProxy->Next(cConnections, vCP.data(), pcFetched);

        std::cout << this << "@CProxiedEnumConnectionPoints::Next(" << cConnections
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << "" << std::endl;

        if (FAILED(nResult))
            return nResult;

        for (ULONG i = 0; i < *pcFetched; ++i)
        {
            IID aIID;
            nResult = vCP[i]->GetConnectionInterface(&aIID);
            if (FAILED(nResult))
                return nResult;

            std::cout << "  " << i << ": " << aIID << "" << std::endl;

            bool bFound = false;
            for (const auto aMapEntry : aOutgoingInterfaceMap)
            {
                const IID aProxiedOrReplacementIID
                    = (CProxiedDispatch::getParam()->mbTraceOnly ? aMapEntry.maOtherAppIID
                                                                 : aMapEntry.maLibreOfficeIID);

                if (IsEqualIID(aIID, aProxiedOrReplacementIID))
                {
                    ITypeInfo* pTI;
                    nResult = mpProxiedDispatch->FindTypeInfo(aProxiedOrReplacementIID, &pTI);
                    if (FAILED(nResult))
                        return nResult;
                    ppCP[i] = new CProxiedConnectionPoint(mpProxiedDispatch, vCP[i],
                                                          aMapEntry.maOtherAppIID, pTI);
                    bFound = true;
                    break;
                }
            }
            if (!bFound)
            {
                std::cerr << "Unhandled outgoing interface returned in connection point: " << aIID
                          << "" << std::endl;
                std::abort();
            }
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Skip(ULONG cConnections) override
    {
        return mpEnumConnectionPointsToProxy->Skip(cConnections);
    }

    HRESULT STDMETHODCALLTYPE Reset() override { return mpEnumConnectionPointsToProxy->Reset(); }

    HRESULT STDMETHODCALLTYPE Clone(IEnumConnectionPoints** ppEnum) override
    {
        IEnumConnectionPoints* pEnumConnectionPoints;
        HRESULT nResult = mpEnumConnectionPointsToProxy->Clone(&pEnumConnectionPoints);
        if (FAILED(nResult))
            return nResult;
        *ppEnum = new CProxiedEnumConnectionPoints(mpProxiedDispatch, pEnumConnectionPoints);
        return S_OK;
    }
};

CProxiedDispatch::CProxiedDispatch(IDispatch* pDispatchToProxy, const IID& aIID)
    : maIID(aIID)
    , maCoclassIID(IID_NULL)
    , mpDispatchToProxy(pDispatchToProxy)
    , mpCPMap(new ConnectionPointMapHolder())
{
    std::cout << this << "@CProxiedDispatch::CTOR(" << maIID << ", " << pDispatchToProxy << ")"
              << std::endl;
}

CProxiedDispatch::CProxiedDispatch(IDispatch* pDispatchToProxy, const IID& aIID,
                                   const IID& aCoclassIID)
    : maIID(aIID)
    , maCoclassIID(aCoclassIID)
    , mpDispatchToProxy(pDispatchToProxy)
    , mpCPMap(new ConnectionPointMapHolder())
{
    std::cout << this << "@CProxiedDispatch::CTOR(" << maIID << ", " << maCoclassIID << ", "
              << pDispatchToProxy << ")" << std::endl;
}

void CProxiedDispatch::setParam(ThreadProcParam* pParam) { pGlobalParam = pParam; }

ThreadProcParam* CProxiedDispatch::getParam() { return pGlobalParam; }

HRESULT CProxiedDispatch::Invoke(std::string sFuncName, int nInvKind,
                                 std::vector<VARIANT>& rParameters, void* pRetval)
{
    HRESULT nResult = S_OK;

    std::wstring sFuncNameWstr
        = std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>, wchar_t>().from_bytes(
            sFuncName.c_str());
    LPOLESTR pFuncNameWstr = (LPOLESTR)sFuncNameWstr.data();

    MEMBERID nMemberId;

    nResult = mpDispatchToProxy->GetIDsOfNames(IID_NULL, &pFuncNameWstr, 1, LOCALE_USER_DEFAULT,
                                               &nMemberId);
    if (nResult == DISP_E_UNKNOWNNAME)
    {
        std::cerr << "GetIDsOfNames(" << sFuncName << "): Not implemented in LibreOffice"
                  << std::endl;
        return E_NOTIMPL;
    }

    if (FAILED(nResult))
    {
        std::cerr << "GetIDsOfNames(" << sFuncName
                  << ") failed: " << WindowsErrorStringFromHRESULT(nResult) << "" << std::endl;
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
            std::cerr << "Unhandled nInvKind in CProxiedDispatch::Invoke: " << nInvKind << ""
                      << std::endl;
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

    std::cout << this << "@CProxiedDispatch::Invoke(" << sFuncName << ")..." << std::endl;

    nResult = mpDispatchToProxy->Invoke(nMemberId, IID_NULL, LOCALE_USER_DEFAULT, nFlags,
                                        &aDispParams, &aResult, NULL, &nArgErr);
    if (FAILED(nResult))
    {
        std::cout << "..." << this << "@CProxiedDispatch::Invoke(" << sFuncName
                  << "): " << WindowsErrorStringFromHRESULT(nResult) << "" << std::endl;
        return nResult;
    }

    if (aResult.vt != VT_EMPTY && aResult.vt != VT_VOID && pRetval != nullptr)
    {
        switch (aResult.vt)
        {
            case VT_BSTR:
                *(BSTR*)pRetval = aResult.bstrVal;
                break;
            case VT_DISPATCH:
                *(IDispatch**)pRetval = aResult.pdispVal;
                break;
            default:
                std::cerr << "Unhandled vt in CProxiedDispatch::Invoke: " << aResult.vt << ""
                          << std::endl;
                std::abort();
        }
    }
    std::cout << "..." << this << "@CProxiedDispatch::Invoke(" << sFuncName << "): S_OK"
              << std::endl;

    return nResult;
}

// IUnknown

HRESULT STDMETHODCALLTYPE CProxiedDispatch::QueryInterface(REFIID riid, void** ppvObject)
{
    HRESULT nResult;

    std::cout << this << "@CProxiedDispatch::QueryInterface(" << riid << "): ";

    if (IsEqualIID(riid, IID_IUnknown))
    {
        std::cout << "self" << std::endl;
        *ppvObject = this;
        AddRef();
        return S_OK;
    }

    if (IsEqualIID(riid, IID_IDispatch))
    {
        std::cout << "self" << std::endl;
        *ppvObject = this;
        AddRef();
        return S_OK;
    }

    if (IsEqualIID(riid, maIID))
    {
        std::cout << "self" << std::endl;
        *ppvObject = this;
        AddRef();
        return S_OK;
    }

    if (!IsEqualIID(maCoclassIID, IID_NULL) && IsEqualIID(riid, maCoclassIID))
    {
        std::cout << "self" << std::endl;
        *ppvObject = this;
        AddRef();
        return S_OK;
    }

    nResult = mpDispatchToProxy->QueryInterface(riid, ppvObject);
    std::cout << WindowsErrorStringFromHRESULT(nResult) << "" << std::endl;

    // We need to special-case IConnectionPointContainer
    if (nResult == S_OK && IsEqualIID(riid, IID_IConnectionPointContainer))
    {
        // We must proxy it so that we can reverse-proxy the calls to advised sinks
        mpConnectionPointContainerToProxy = (IConnectionPointContainer*)*ppvObject;
        *ppvObject = (IConnectionPointContainer*)this;
        return S_OK;
    }

    nResult = mpDispatchToProxy->QueryInterface(riid, ppvObject);
    return nResult;
}

ULONG STDMETHODCALLTYPE CProxiedDispatch::AddRef()
{
    ULONG nRetval = mpDispatchToProxy->AddRef();
    std::cout << this << "@CProxiedDispatch::AddRef: " << nRetval << "" << std::endl;

    return nRetval;
}

ULONG STDMETHODCALLTYPE CProxiedDispatch::Release()
{
    ULONG nRetval = mpDispatchToProxy->Release();
    std::cout << this << "@CProxiedDispatch::Release: " << nRetval << "" << std::endl;

    return nRetval;
}

// IDispatch

HRESULT STDMETHODCALLTYPE CProxiedDispatch::GetTypeInfoCount(UINT* pctinfo)
{
    return mpDispatchToProxy->GetTypeInfoCount(pctinfo);
}

HRESULT STDMETHODCALLTYPE CProxiedDispatch::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    return mpDispatchToProxy->GetTypeInfo(iTInfo, lcid, ppTInfo);
}

HRESULT STDMETHODCALLTYPE CProxiedDispatch::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames,
                                                          UINT cNames, LCID lcid, DISPID* rgDispId)
{
    return mpDispatchToProxy->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
}

HRESULT STDMETHODCALLTYPE CProxiedDispatch::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
                                                   WORD wFlags, DISPPARAMS* pDispParams,
                                                   VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                                                   UINT* puArgErr)
{
    return mpDispatchToProxy->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult,
                                     pExcepInfo, puArgErr);
}

HRESULT CProxiedDispatch::FindTypeInfo(const IID& rIID, ITypeInfo** pTypeInfo)
{
    std::cout << this << "@CProxiedDispatch::FindTypeInfo(" << rIID << ")..." << std::endl;
    HRESULT nResult;

    ITypeInfo* pTI;

    nResult = GetTypeInfo(0, LOCALE_USER_DEFAULT, &pTI);
    if (FAILED(nResult))
    {
        std::cout << "..." << this << "@CProxiedDispatch::FindTypeInfo(" << rIID << ") ("
                  << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }

    TYPEATTR* pTA;
    nResult = pTI->GetTypeAttr(&pTA);
    if (FAILED(nResult))
    {
        std::cout << "..." << this << "@CProxiedDispatch::FindTypeInfo(" << rIID << ") ("
                  << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
        return nResult;
    }

    bool bFound = false;
    for (WORD i = 0; i < pTA->cImplTypes; ++i)
    {
        HREFTYPE nHrefType;
        nResult = pTI->GetRefTypeOfImplType(i, &nHrefType);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedDispatch::FindTypeInfo(" << rIID << ") ("
                      << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
            return nResult;
        }

        nResult = pTI->GetRefTypeInfo(nHrefType, pTypeInfo);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedDispatch::FindTypeInfo(" << rIID << ") ("
                      << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
            return nResult;
        }

        TYPEATTR* pRefTA;
        nResult = (*pTypeInfo)->GetTypeAttr(&pRefTA);
        if (FAILED(nResult))
        {
            std::cout << "..." << this << "@CProxiedDispatch::FindTypeInfo(" << rIID << ") ("
                      << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult) << std::endl;
            return nResult;
        }

        std::cout << "..." << this << "@CProxiedDispatch::FindTypeInfo(" << rIID << ") ("
                  << __LINE__ << "): " << rIID << ":" << pRefTA->guid << "\n";

        if (IsEqualIID(rIID, pRefTA->guid))
        {
            (*pTypeInfo)->ReleaseTypeAttr(pRefTA);
            bFound = true;
            break;
        }
        (*pTypeInfo)->ReleaseTypeAttr(pRefTA);
    }
    pTI->ReleaseTypeAttr(pTA);

    if (!bFound)
    {
        std::cout << "..." << this << "@CProxiedDispatch::FindTypeInfo(" << rIID << ") ("
                  << __LINE__ << "): E_NOTIMPL" << std::endl;
        return E_NOTIMPL;
    }

    std::cout << "..." << this << "@CProxiedDispatch::FindTypeInfo(" << rIID << "): S_OK"
              << std::endl;
    return S_OK;
}

// IConnectionPointContainer

HRESULT STDMETHODCALLTYPE CProxiedDispatch::EnumConnectionPoints(IEnumConnectionPoints** ppEnum)
{
    HRESULT nResult;

    if (!ppEnum)
        return E_POINTER;

    std::cout << "CProxiedDispatch::EnumConnectionPoints" << std::endl;

    IEnumConnectionPoints* pEnumConnectionPoints;
    nResult = mpConnectionPointContainerToProxy->EnumConnectionPoints(&pEnumConnectionPoints);

    if (FAILED(nResult))
        return nResult;

    *ppEnum = new CProxiedEnumConnectionPoints(this, pEnumConnectionPoints);

    return S_OK;
}

HRESULT STDMETHODCALLTYPE CProxiedDispatch::FindConnectionPoint(REFIID riid,
                                                                IConnectionPoint** ppCP)
{
    HRESULT nResult;

    if (!ppCP)
        return E_POINTER;

    if (mpCPMap->maConnectionPoints.count(riid))
    {
        std::cout << this << "@CProxiedDispatch::FindConnectionPoint(" << riid << "): S_OK"
                  << std::endl;
        *ppCP = mpCPMap->maConnectionPoints[riid];
        (*ppCP)->AddRef();
        return S_OK;
    }

    for (const auto aMapEntry : aOutgoingInterfaceMap)
    {
        if (IsEqualIID(riid, aMapEntry.maOtherAppIID))
        {
            ITypeInfo* pTI;
            const IID aProxiedOrReplacementIID
                = (getParam()->mbTraceOnly ? aMapEntry.maOtherAppIID : aMapEntry.maLibreOfficeIID);
            nResult = FindTypeInfo(aProxiedOrReplacementIID, &pTI);
            if (FAILED(nResult))
            {
                std::cout << this << "@CProxiedDispatch::FindConnectionPoint(" << riid << ") ("
                          << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult)
                          << std::endl;
                return nResult;
            }

            IConnectionPoint* pCP;
            nResult = mpConnectionPointContainerToProxy->FindConnectionPoint(
                aProxiedOrReplacementIID, &pCP);
            if (FAILED(nResult))
            {
                std::cout << this << "@CProxiedDispatch::FindConnectionPoint(" << riid << ") ("
                          << __LINE__ << "): " << WindowsErrorStringFromHRESULT(nResult) << ""
                          << std::endl;
                return nResult;
            }

            std::cout << this << "@CProxiedDispatch::FindConnectionPoint(" << riid << ")..."
                      << std::endl;

            *ppCP = new CProxiedConnectionPoint(this, pCP, aMapEntry.maOtherAppIID, pTI);

            std::cout << "..." << this << "@CProxiedDispatch::FindConnectionPoint(" << riid
                      << "): S_OK" << std::endl;

            mpCPMap->maConnectionPoints[riid] = *ppCP;
            (*ppCP)->AddRef();

            return S_OK;
        }
    }

    std::cout << this << "@CProxiedDispatch::FindConnectionPoint(" << riid << ") (" << __LINE__
              << "): E_NOTIMPL" << std::endl;
    return CONNECT_E_NOCONNECTION;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */

//
// AddObj.cpp
// Contains the implementation of IUnknown and IAdd interfaces
//
#include "..\common\pch.h"
#include "AddObj.h"
#include "AddObject_i.c" // Contains the interface IIDs
#include "utilities.h"

#define DEBUG_LOGGER_ENABLED
#define FILE_LOGGER_ENABLED
#define LOG_PREFIX "[ADDOBJ-FASTADD]"
#include "logger.h"

//
// CAddObj constructor/destructor
//
CAddObj::CAddObj() : m_nRefCount(0)
{
    InterlockedIncrement(&g_nComObjectsInUse);
}

CAddObj::~CAddObj()
{
    InterlockedDecrement(&g_nComObjectsInUse);
}

//
// IUnknown interface implementation
//
HRESULT __stdcall CAddObj::QueryInterface(REFIID riid, void** ppvObject)
{
    LOG("Entering");

    WCHAR wsIID[GUID_STRING_LENGTH] = { 0 };
    BOOL bRes = GuidToWideString(riid, wsIID, GUID_STRING_LENGTH);
    if (!bRes || !ppvObject)
        return E_INVALIDARG;

    LOG("IID: %ws, ppvObject: 0x%p, pvObject: 0x%p", wsIID, ppvObject, *ppvObject);

    if (riid == IID_IUnknown) {
        *ppvObject = static_cast<void*>(this);
        AddRef();
        LOG("Query for IUnknown, refCount: %d", m_nRefCount);
        return S_OK;
    }

    if (riid == IID_IAdd) {
        *ppvObject = static_cast<void*>(this);
        AddRef();
        LOG("Query for IAdd, refCount: %d", m_nRefCount);
        return S_OK;
    }

    WCHAR wsIIDName[MAX_PATH] = { 0 };
    GetInterfaceName(riid, wsIIDName, MAX_PATH);
    LOG("!!! Not supported interface: %ws, %ws", wsIID, wsIIDName);

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG __stdcall CAddObj::AddRef() {
    LOG("On instance 0x%p", this);
    return InterlockedIncrement(&m_nRefCount);
}

ULONG __stdcall CAddObj::Release() {
    LOG("On instance 0x%p", this);
    long nRefCount = 0;
    nRefCount = InterlockedDecrement(&m_nRefCount);
    if (0 == nRefCount) {
        LOG("Cleanup instance 0x%p", this);
        delete this;
    }
    return nRefCount;
}

//
// IAdd interface implementation
//
HRESULT __stdcall CAddObj::SetFirstNumber(long nX1)
{
    m_nX1 = nX1;
    LOG("X1: %d", m_nX1);
    return S_OK;
}

HRESULT __stdcall CAddObj::SetSecondNumber(long nX2)
{
    m_nX2 = nX2;
    LOG("X2: %d", m_nX2);
    return S_OK;
}

HRESULT __stdcall CAddObj::PerformAddition(long* pSum)
{
    if (!pSum) {
        LOG("Invalid argument. Exiting");
        return E_INVALIDARG;
    }

    *pSum = m_nX1 + m_nX2;
    LOG("SuperFast addition algorithm produced the result: %d", *pSum);

    return S_OK;
}

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

extern ULONG g_ObjectCount;

extern PCWSTR g_wsMessageBoxTitle;

//
// CAddObj constructor/destructor
//
CAddObj::CAddObj() : m_refCount(0)
{
    LOG("Constructing instance: 0x%p", this);
    InterlockedIncrement(&g_ObjectCount);
}

CAddObj::~CAddObj()
{
#ifdef _DEBUG
    MessageBox(
        NULL,
        L"CAddObj is being destructed.\nMake sure you see this message. If not, you might have memory leak!",
        g_wsMessageBoxTitle,
        MB_OK | MB_SETFOREGROUND
    );
#endif
    InterlockedDecrement(&g_ObjectCount);
    LOG("Destructing instance: 0x%p", this);
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

    WCHAR wsIIDName[MAX_PATH] = { 0 };
    GetInterfaceName(riid, wsIIDName, MAX_PATH);

    LOG("IID: %ws [%ws], ppvObject: 0x%p", wsIID, wsIIDName, ppvObject);

    if (riid == IID_IUnknown) {
        *ppvObject = static_cast<void*>(this);
        AddRef();
        LOG("Query for IUnknown, refCount: %d", m_refCount);
        return S_OK;
    }

    if (riid == IID_IAdd) {
        *ppvObject = static_cast<void*>(this);
        AddRef();
        LOG("Query for IAdd, refCount: %d", m_refCount);
        return S_OK;
    }

    LOG("WARNING! Not supported interface: %ws [%ws], refCount: %d", wsIID, wsIIDName, m_refCount);

    *ppvObject = NULL;
    return E_NOINTERFACE;
}

ULONG __stdcall CAddObj::AddRef() {
    ULONG refCount = InterlockedIncrement(&m_refCount);
    LOG("On instance 0x%p, refCount: %d", this, refCount);
    return refCount;
}

ULONG __stdcall CAddObj::Release() {
    ULONG refCount = InterlockedDecrement(&m_refCount);
    LOG("On instance 0x%p, refCount: %d", this, refCount);
    if (0 == refCount) {
        LOG("Cleanup instance 0x%p", this);
        delete this;
    }
    return refCount;
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

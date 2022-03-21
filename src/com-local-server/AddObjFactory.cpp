//
// AddObjFactory.cpp
// Contains the implementation of IUnknown and IClassFactory interfaces
//
#include "..\common\pch.h"
#include "AddObj.h"
#include "AddObjFactory.h"
#include "utilities.h"

#define DEBUG_LOGGER_ENABLED
#define FILE_LOGGER_ENABLED
#define LOG_PREFIX "[ADDOBJ-FACTORY]"
#include "logger.h"

extern ULONG g_ObjectCount;
extern ULONG g_LockCount;

extern PCWSTR g_wsMessageBoxTitle;

//
// CAddObjFactory ctor / dtor
//
CAddObjFactory::CAddObjFactory() : m_refCount(0)
{
    LOG("Constructing instance: 0x%p", this);
    InterlockedIncrement(&g_ObjectCount);
}

CAddObjFactory::~CAddObjFactory() {
#ifdef _DEBUG
    MessageBox(
        NULL,
        L"CAddObjFactory is being destructed.\nMake sure you see this message. If not, you might have memory leak!",
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
HRESULT __stdcall CAddObjFactory::QueryInterface(REFIID riid, void** ppvObject)
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

    if (riid == IID_IClassFactory) {
        *ppvObject = static_cast<void*>(this);
        AddRef();
        LOG("Query for IClassFactory, refCount: %d", m_refCount);
        return S_OK;
    }

    LOG("WARNING! Not supported interface: %ws [%ws], refCount: %d", wsIID, wsIIDName, m_refCount);

    *ppvObject = NULL;
    return E_NOINTERFACE;
}

ULONG __stdcall CAddObjFactory::AddRef() {
    ULONG refCount = InterlockedIncrement(&m_refCount);
    LOG("On instance 0x%p, refCount: %d", this, refCount);
    return refCount;
}

ULONG __stdcall CAddObjFactory::Release() {
    ULONG refCount = InterlockedDecrement(&m_refCount);
    LOG("On instance 0x%p, refCount: %d", this, refCount);
    if (0 == refCount) {
        LOG("Cleanup instance 0x%p", this);
        delete this;
    }
    return refCount;
}

//
// IClassFactory interface implementation
//
HRESULT __stdcall CAddObjFactory::CreateInstance(_In_opt_ IUnknown* pUnkOuter, _In_ REFIID riid, _COM_Outptr_ void** ppvObject)
{
    LOG("Entering");

    WCHAR wsIID[GUID_STRING_LENGTH] = { 0 };
    BOOL bRes = GuidToWideString(riid, wsIID, GUID_STRING_LENGTH);
    if (!bRes || !ppvObject)
        return E_INVALIDARG;

    LOG("IID: %ws, pUnkOuter: 0x%p, ppvObject: 0x%p, pvObject: 0x%p", wsIID, pUnkOuter, ppvObject, *ppvObject);

    if (pUnkOuter != NULL)
        return CLASS_E_NOAGGREGATION; // Cannot aggregate

    //
    // Create and instance of the component
    //
    CAddObj* pObject = new CAddObj;
    LOG("Created CAddObj instance: 0x%p", pObject);
    if (!pObject)
        return E_OUTOFMEMORY;

    //
    // Get the requested interface
    //
    HRESULT hr = pObject->QueryInterface(riid, ppvObject);
    LOG("CAddObj->QueryInterface(%ws) returned 0x%08x", wsIID, hr);
    if (FAILED(hr))
        delete pObject;

    return hr;
}

HRESULT __stdcall CAddObjFactory::LockServer(_In_ BOOL fLock)
{
    LOG("fLock: %d", fLock);

    if (fLock) {
        long lockCount = InterlockedDecrement(&g_LockCount);

        if (0 == lockCount) {
            LOG("PostQuitMessage(EXIT_SUCCESS)");
            PostQuitMessage(EXIT_SUCCESS);
        }
    }
    else
        InterlockedDecrement(&g_LockCount);

    return S_OK;
}
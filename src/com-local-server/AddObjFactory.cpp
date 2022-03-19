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

//
// CAddObjFactory constructor
//
CAddObjFactory::CAddObjFactory() : m_nRefCount(0) {}

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

    LOG("IID: %ws, ppvObject: 0x%p, pvObject: 0x%p", wsIID, ppvObject, *ppvObject);

    if (riid == IID_IUnknown) {
        *ppvObject = static_cast<void*>(this);
        AddRef();
        LOG("Query for IUnknown, refCount: %d", m_nRefCount);
        return S_OK;
    }

    if (riid == IID_IClassFactory) {
        *ppvObject = static_cast<void*>(this);
        AddRef();
        LOG("Query for IClassFactory, refCount: %d", m_nRefCount);
        return S_OK;
    }

    WCHAR wsIIDName[MAX_PATH] = { 0 };
    GetInterfaceName(riid, wsIIDName, MAX_PATH);
    LOG("!!! Not supported interface: %ws, %ws", wsIID, wsIIDName);

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG __stdcall CAddObjFactory::AddRef() {
    LOG("On instance 0x%p", this);
    return InterlockedIncrement(&m_nRefCount);
}

ULONG __stdcall CAddObjFactory::Release() {
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
// IClassFactory interface implementation
//
HRESULT __stdcall CAddObjFactory::CreateInstance(_In_opt_ IUnknown* pUnkOuter, _In_ REFIID riid, _COM_Outptr_ void** ppvObject)
{
    //
    // This method lets the client manufacture components en masse.
    // The class factory provides a mechanism to control the way
    // the component is created. Within the class factory the
    // author of the component may decide to selectivey enable
    // or disable creation as per license agreements
    //

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
    if (!pObject)
        return E_OUTOFMEMORY;

    //
    // Get the requested interface
    //
    LOG("Created CAddObj instance: 0x%p, now requesting interface %ws", pObject, wsIID);
    return pObject->QueryInterface(riid, ppvObject);
}

HRESULT __stdcall CAddObjFactory::LockServer(_In_ BOOL fLock)
{
    LOG("Not implemented");
    return E_NOTIMPL;
}
//
// Exports.cpp
// Contains the implementation of the exported functions that are
// called by the COM framework to create / destroy the COM component.
//
#include "..\common\pch.h"
#include "AddObj.h"
#include "AddObjFactory.h"
#include "IAdd_i.c" // Contains the interface IIDs
#include "utilities.h"

#define DEBUG_LOGGER_ENABLED
#define FILE_LOGGER_ENABLED
#define LOG_PREFIX "[ADDOBJ-EXPORTS]"
#include "logger.h"

STDAPI DllGetClassObject(const CLSID& clsid, const IID& riid, void** ppvObject)
{
    //
    // Check if the requested COM object is implemented in this DLL.
    // There can be more than 1 COM object implemented in a DLL.
    //

    LOG("Entering");

    WCHAR wsIID[GUID_STRING_LENGTH] = { 0 };
    BOOL bRes = GuidToWideString(riid, wsIID, GUID_STRING_LENGTH);
    if (!bRes || !ppvObject)
        return E_INVALIDARG;

    WCHAR wsCLSID[GUID_STRING_LENGTH] = { 0 };
    bRes = GuidToWideString(clsid, wsCLSID, GUID_STRING_LENGTH);
    if (!bRes)
        return E_INVALIDARG;

    LOG("IID: %ws, CLSID: %ws, ppvObject: 0x%p, pvObject: 0x%p", wsIID, wsCLSID, ppvObject, *ppvObject);

    //
    // If control reaches here then that implies that the object
    // specified by the user is not implemented in this DLL
    //
    if (clsid != CLSID_AddObject) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    //
    // iid specifies the requested interface for the factory object
    // The client can request for IUnknown, IClassFactory, IClassFactory2
    //
    CAddObjFactory* pAddFactory = new CAddObjFactory;
    if (!pAddFactory)
        return E_OUTOFMEMORY;

    //
    // Get the requested interface
    //
    LOG("Created CAddObjFactory instance: 0x%p, now requesting interface %ws", pAddFactory, wsIID);
    return pAddFactory->QueryInterface(riid, ppvObject);
}

STDAPI DllCanUnloadNow()
{
    //
    // A DLL is no longer in use when it is not managing any existing objects
    // (the reference count on all of its objects is 0).
    // We will examine the value of g_nComObjectsInUse
    //

    LOG("COM objects in use: %d", g_nComObjectsInUse);

    return 0 == g_nComObjectsInUse ? S_OK : S_FALSE;
}
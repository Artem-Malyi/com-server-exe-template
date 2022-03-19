//
// AddObjFactory.h
// Contains the C++ class declaration for IClassFactory implementation
//
#pragma once
#include "unknwn.h"

class CAddObjFactory : public IClassFactory
{
public:
    //CAddFactory::CAddFactory() : m_nRefCount(0) {}
    CAddObjFactory();

    // IUnknown interface
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppObj);
    ULONG   __stdcall AddRef();
    ULONG   __stdcall Release();

    // IClassFactory interface
    HRESULT __stdcall CreateInstance(_In_opt_ IUnknown* pUnkOuter, _In_ REFIID riid, _COM_Outptr_ void** ppvObject);
    HRESULT __stdcall LockServer(_In_ BOOL fLock);

private:
    long m_nRefCount;
};
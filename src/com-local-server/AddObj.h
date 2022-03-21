//
// AddObj.h
// Contains the C++ class declaration for the IAdd interface
//
#pragma once
#include "AddObject_h.h" // Contains the C++ style interface declarations.

class CAddObj: public IAdd
{
public:
    CAddObj();
    virtual ~CAddObj();

    // IUnknown interface
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppObj);
    ULONG   __stdcall AddRef();
    ULONG   __stdcall Release();

    // IAdd interface
    HRESULT __stdcall SetFirstNumber(long nX1);
    HRESULT __stdcall SetSecondNumber(long nX2);
    HRESULT __stdcall PerformAddition(long* pSum);

private:
    long m_nX1, m_nX2; // operands for addition
    ULONG m_refCount;  // for managing the reference count
};

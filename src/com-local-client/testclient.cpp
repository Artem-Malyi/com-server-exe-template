// com-local-client.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

//#define BUILD_CLIENT_WITH_TYPELIB
#define BUILD_CLIENT_WITH_IDL_HEADERS

#include <iostream>
#include <objbase.h>
#include "utilities.h"

#define DEBUG_LOGGER_ENABLED
#define FILE_LOGGER_ENABLED
#define LOG_PREFIX "[ADDOBJ-CLIENT]"
#include "logger.h"

int testComServer();

#ifdef BUILD_CLIENT_WITH_IDL_HEADERS
//
// To make this work enable the Post-Build event in com-local-server project
//
#include "AddObject_h.h"
#include "AddObject_i.c"
#endif

#ifdef BUILD_CLIENT_WITH_TYPELIB
//
// Here we do a #import on the DLL, you can also do a #import on the .TLB
// The #import directive generates two files (.tlh/.tli) in the output folders.
//
//#import "AddObject.tlb"
#ifdef _WIN64
    #ifdef _DEBUG
        #import "com-local-server64d.exe"
    #else
        #import "com-local-server64.dll"
    #endif
#else
    #ifdef _DEBUG
        #import "com-local-server32d.dll"
    #else
        #import "com-local-server32.dll"
    #endif
#endif
#endif // BUILD_CLIENT_WITH_TYPELIB

int main()
{
    return testComServer();
}

#ifdef BUILD_CLIENT_WITH_TYPELIB
int testComServer()
{
    LOG("Entering");

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
        return -1;

    WCHAR wsMessageBuffer[512] = { 0 };

    {
        SuperFastMathLib::IAddPtr pFastAddAlgorithm;
        //
        // IAddPtr is not the actual interface IAdd, but a template C++ class (_com_ptr_t)
        // that contains an embedded instance of the raw IAdd pointer.
        // While destructing, the destructor makes sure to invoke Release() on the internal
        // raw interface pointer. Further, the operator -> has been overloaded to direct all
        // method invocations to the internal raw interface pointer.
        //
        hr = pFastAddAlgorithm.CreateInstance("SuperFast.AddObj");
        ErrorDescription(hr, wsMessageBuffer, _countof(wsMessageBuffer));
        LOG("CoCreateInstance() returned 0x%08x: %ws", hr, wsMessageBuffer);
        if (FAILED(hr))
            return -2;

        long n1 = 100, n2 = 200;
        pFastAddAlgorithm->SetFirstNumber(n1); //"->" overloading in action
        pFastAddAlgorithm->SetSecondNumber(n2);

        long nSum = 0;
        nSum = pFastAddAlgorithm->PerformAddition();

        std::cout << "Output after adding " << n1 << " & " << n2 << " is " << nSum << "\n";
        LOG("SuperFast addition algorithm returned: %d + %d = %d", n1, n2, nSum);
    }

    CoUninitialize();

    return 0;
}
#endif // BUILD_CLIENT_WITH_TYPELIB

#ifdef BUILD_CLIENT_WITH_IDL_HEADERS
int testComServer()
{
    LOG("Entering");
    WCHAR wsMessageBuffer[MAX_PATH] = { 0 };

    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr))
        return -1;

    IClassFactory* pICF = nullptr;
    hr = CoGetClassObject(CLSID_AddObject, CLSCTX_LOCAL_SERVER, nullptr, IID_IClassFactory, (void**)&pICF);
    ErrorDescription(hr, wsMessageBuffer, _countof(wsMessageBuffer));
    LOG("CoGetClassObject() returned 0x%08x: %ws", hr, wsMessageBuffer);
    if (FAILED(hr)) {
        CoUninitialize();
        exit(-2);
    }

    IAdd* pIAdd = nullptr;
    hr = pICF->CreateInstance(nullptr, IID_IAdd, (void**)&pIAdd);
    ErrorDescription(hr, wsMessageBuffer, _countof(wsMessageBuffer));
    LOG("IClassFactory->CreateInstance() returned 0x%08x: %ws", hr, wsMessageBuffer);
    if (FAILED(hr)) {
        pICF->Release();
        CoUninitialize();
        exit(-3);
    }

    long n1 = 100, n2 = 200;
    pIAdd->SetFirstNumber(n1);
    pIAdd->SetSecondNumber(n2);

    long nSum = 0;
    hr = pIAdd->PerformAddition(&nSum);
    ErrorDescription(hr, wsMessageBuffer, _countof(wsMessageBuffer));
    LOG("IAdd->PerformAddition() returned 0x%08x: %ws", hr, wsMessageBuffer);
    if (FAILED(hr)) {
        pICF->Release();
        pIAdd->Release();
        CoUninitialize();
        return -4;
    }

    std::cout << "Output after adding " << n1 << " & " << n2 << " is " << nSum << "\n";
    LOG("SuperFast addition algorithm returned: %d + %d = %d", n1, n2, nSum);

    //
    // Release the reference to the COM object when we're done
    //
    pICF->Release();
    pIAdd->Release();

    CoUninitialize();

    return 0;
}
#endif // BUILD_CLIENT_WITH_IDL_HEADERS

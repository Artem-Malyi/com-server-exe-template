// main.cpp : Defines the entry point for the DLL application.
#include "..\common\pch.h"

#include "AddObjFactory.h"
#include "AddObject_h.h"
#include "AddObject_i.c"
#include "utilities.h"

#define DEBUG_LOGGER_ENABLED
#define FILE_LOGGER_ENABLED
#define LOG_PREFIX "[ADDOBJ-WINMAIN]"
#include "logger.h"

STDAPI DllRegisterServer(void);
STDAPI DllUnregisterServer(void);

// Global COM server variables.
ULONG g_LockCount = 0;		// Represents the number of COM server locks.
ULONG g_ObjectCount = 0;	// Represents the number of living COM objects.

PCWSTR g_wsMessageBoxTitle = L"SuperFast.AddObj COM LocalServer Message";

int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	LOG("Entering, hInstance: 0x%p, lpCmdLine: %s, nCmdShow: %d", hInstance, lpCmdLine, nCmdShow);

	WCHAR wsMessageBuffer[MAX_PATH] = { 0 }, wsMessage[MAX_PATH] = { 0 };
	HRESULT hr = E_FAIL;
    WCHAR wsImageName[MAX_PATH] = { 0 };
    GetModuleFileNameW(GetModuleHandleW(NULL), wsImageName, MAX_PATH);

    // Handle [un]registration of COM server
    if (strstr(lpCmdLine, "/r") || strstr(lpCmdLine, "-r")) {
        hr = DllRegisterServer();
        ErrorDescription(hr, wsMessageBuffer, _countof(wsMessageBuffer));
        LOG("DllRegisterServer() returned 0x%08x: %ws", hr, wsMessageBuffer);
        if (FAILED(hr)) {
            StringCbPrintf(wsMessage, sizeof(wsMessage), L"The call to DllRegisterServer failed with error code 0x08x.", hr);
            MessageBox(NULL, wsMessage, g_wsMessageBoxTitle, MB_OK | MB_ICONERROR);
            return -1;
        }
        else {
            StringCbPrintf(wsMessage, sizeof(wsMessage), L"DllRegisterServer in %s succeeded.", wsImageName);
            MessageBox(NULL, wsMessage, g_wsMessageBoxTitle, MB_OK | MB_ICONINFORMATION);
            return 0;
        }
    }
    else if (strstr(lpCmdLine, "/u") || strstr(lpCmdLine, "-u")) {
        hr = DllUnregisterServer();
        ErrorDescription(hr, wsMessageBuffer, _countof(wsMessageBuffer));
        LOG("DllUnregisterServer() returned 0x%08x: %ws", hr, wsMessageBuffer);
        if (FAILED(hr)) {
            StringCbPrintf(wsMessage, sizeof(wsMessage), L"The call to DllUnregisterServer failed with error code 0x08x.", hr);
            MessageBox(NULL, wsMessage, g_wsMessageBoxTitle, MB_OK | MB_ICONERROR);
            return -1;
        }
        else {
            StringCbPrintf(wsMessage, sizeof(wsMessage), L"DllUnregisterServer in %s succeeded.", wsImageName);
            MessageBox(NULL, wsMessage, g_wsMessageBoxTitle, MB_OK | MB_ICONINFORMATION);
            return 0;
        }
    }
    else if (!(strstr(lpCmdLine, "/Embedding") || strstr(lpCmdLine, "-Embedding"))) {
        LOG("Not started by SCM, hence exiting.");
        return 0;
    }

    // Started by SCM.
#ifdef _DEBUG
    MessageBox(NULL, L"AddObj Local Server is registering.", g_wsMessageBoxTitle, MB_OK | MB_SETFOREGROUND);
#endif
    hr = CoInitialize(NULL);
    ErrorDescription(hr, wsMessageBuffer, _countof(wsMessageBuffer));
    LOG("CoInitialize() returned 0x%08x: %ws", hr, wsMessageBuffer);
    if (FAILED(hr))
        return -2;

    // Create the AddObj factory class object.
    // Note, that it will manage its reference count and delete itself when refCount == 0.
    // So, this object MUST be created on the heap.
    CAddObjFactory* pFactory = new CAddObjFactory();

    // Register the AddObj Factory.
    DWORD regId = 0;
    hr = CoRegisterClassObject(CLSID_AddObject, pFactory, CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE, &regId);
    if (FAILED(hr))
    {
        ErrorDescription(hr, wsMessageBuffer, _countof(wsMessageBuffer));
        LOG("CoRegisterClassObject() returned 0x%08x: %ws", hr, wsMessageBuffer);
        CoUninitialize();
        return -3;
    }

    // Now just run until a quit message is sent, in responce to the final release.
    LOG("Entering main thread message loop");
    MSG message = { 0 };
    while (GetMessage(&message, 0, 0, 0))
    {
        TranslateMessage(&message);
        DispatchMessage(&message);
    }
    LOG("Exiting main thread message loop");

    // All done, so remove class object.
    hr = CoRevokeClassObject(regId);
    ErrorDescription(hr, wsMessageBuffer, _countof(wsMessageBuffer));
    LOG("CoRevokeClassObject() returned 0x%08x: %ws", hr, wsMessageBuffer);

    // Terminate COM.
    CoUninitialize();

#ifdef _DEBUG
		MessageBox(NULL, L"AddObj Local Server is exiting", g_wsMessageBoxTitle, MB_OK | MB_SETFOREGROUND);
#endif

	return 0;
}

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

// Global COM server variables.
ULONG g_LockCount = 0;		// Represents the number of COM server locks.
ULONG g_ObjectCount = 0;	// Represents the number of living COM objects.

const WCHAR* g_wsMessageBoxTitle = L"SuperFast.AddObj COM LocalServer Message";

int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	LOG("Entering, hInstance: 0x%p, lpCmdLine: %s, nCmdShow: %d", hInstance, lpCmdLine, nCmdShow);

	WCHAR wsMessageBuffer[MAX_PATH] = { 0 };

	HRESULT hr = CoInitialize(nullptr);
	ErrorDescription(hr, wsMessageBuffer, _countof(wsMessageBuffer));
	LOG("CoInitialize() returned 0x%08x: %ws", hr, wsMessageBuffer);
	if (FAILED(hr))
		return -1;

	// Let's see if we were started by SCM.
	if (strstr(lpCmdLine, "/Embedding") || strstr(lpCmdLine, "-Embedding"))
	{
#ifdef _DEBUG
		MessageBox(nullptr, L"AddObj Local Server is registering.", g_wsMessageBoxTitle, MB_OK | MB_SETFOREGROUND);
#endif

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
			return -2;
		}

		// Now just run until a quit message is sent,
		// in responce to the final release.
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
	}

	// Terminate COM.
	CoUninitialize();
	MessageBox(nullptr, L"AddObj Local Server is exiting", g_wsMessageBoxTitle, MB_OK | MB_SETFOREGROUND);
	return 0;
}

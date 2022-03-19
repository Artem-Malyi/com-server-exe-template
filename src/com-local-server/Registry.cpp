//
// Registry.cpp
// Contains the implementation of the exported functions that are
// called by regsvr32.exe to register / unregister the COM component.
//
#include "..\common\pch.h"
#include "AddObj.h"
#include "AddObjFactory.h"
#include "AddObject_i.c" // Contains the interface IIDs
#include "tchar.h"

#define DEBUG_LOGGER_ENABLED
#define FILE_LOGGER_ENABLED
#define LOG_PREFIX "[ADDOBJ-EXPORTS]"
#include "logger.h"

static const TCHAR g_szClassDescription[] = TEXT("SuperFast Addition COM Object");
static const TCHAR g_szProgId[] = TEXT("SuperFast.AddObj");
static const TCHAR g_szBoth[] = TEXT("Both");
static const TCHAR g_szThreadingModel[] = TEXT("ThreadingModel");
static const TCHAR g_szInProcServer[] = TEXT("InProcServer32");
static const TCHAR g_szTypeLib[] = TEXT("TypeLib");
static const TCHAR g_szVersion[] = TEXT("Version");
static const TCHAR g_szVersionValue[] = TEXT("1.0");
static const TCHAR g_szTypelibDescription[] = TEXT("Interfaces for SuperFast math lib");
static const TCHAR g_szInterfaceName[] = TEXT("IAdd");

#define OnErrorJump(hr, lRet, label)   \
    if (lRet != ERROR_SUCCESS) {       \
        hr = HRESULT_FROM_WIN32(lRet); \
        goto label;                    \
    }                                  \

STDAPI AddProgIdEntry(LPOLESTR szClassId)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    TCHAR szProgIdRegKey[MAX_PATH] = { 0 };
    HKEY hProgIdKey = nullptr;

    hr = StringCbPrintf(szProgIdRegKey, sizeof(szProgIdRegKey), TEXT("Software\\Classes\\%s\\CLSID"), g_szProgId);
    if (FAILED(hr))
        goto done;

    // HKCU\Software\Classes\ProgId\CLSID
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szProgIdRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hProgIdKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\ProgId\CLSID @=ClassGUID
    lRet = RegSetValueEx(hProgIdKey, NULL, 0, REG_SZ, (PBYTE)szClassId, static_cast<DWORD>(_tcslen(szClassId) * sizeof(TCHAR)));
    OnErrorJump(hr, lRet, done);

done:
    if (hProgIdKey)
        RegCloseKey(hProgIdKey);

    return hr;
}

STDAPI AddClassIdEntry(LPOLESTR szClassId, LPOLESTR szTypelibId, LPTSTR szFileName)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    TCHAR szClassIdRegKey[MAX_PATH] = { 0 };
    HKEY hClsidKey = nullptr, hInProcKey = nullptr, hTypelibKey = nullptr, hVersionKey = nullptr;

    hr = StringCbPrintf(szClassIdRegKey, sizeof(szClassIdRegKey), TEXT("Software\\Classes\\CLSID\\%s"), szClassId);
    if (FAILED(hr))
        goto done;

    // HKCU\Software\Classes\CLSID\ClassGUID
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szClassIdRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hClsidKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\CLSID\ClassGUID @=ClassDescription
    lRet = RegSetValueEx(hClsidKey, NULL, 0, REG_SZ, (PBYTE)g_szClassDescription, sizeof(g_szClassDescription));
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\CLSID\ClassGUID\InprocServer32
    lRet = RegCreateKeyEx(hClsidKey, g_szInProcServer, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hInProcKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\CLSID\ClassGUID\InprocServer32 @=PathToDll
    lRet = RegSetValueEx(hInProcKey, NULL, 0, REG_SZ, (PBYTE)szFileName, static_cast<DWORD>(_tcslen(szFileName) * sizeof(TCHAR)));
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\CLSID\ClassGUID\InprocServer32 ThreadingModel=Both
    lRet = RegSetValueEx(hInProcKey, g_szThreadingModel, 0, REG_SZ, (PBYTE)g_szBoth, sizeof(g_szBoth));
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\CLSID\ClassGUID\TypeLib
    lRet = RegCreateKeyEx(hClsidKey, g_szTypeLib, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hTypelibKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\CLSID\ClassGUID\TypeLib @=TypelibGUID
    lRet = RegSetValueEx(hTypelibKey, NULL, 0, REG_SZ, (PBYTE)szTypelibId, static_cast<DWORD>(_tcslen(szTypelibId) * sizeof(TCHAR)));
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\CLSID\ClassGUID\Version
    lRet = RegCreateKeyEx(hClsidKey, g_szVersion, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hVersionKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\CLSID\ClassGUID\Version @=1.0
    lRet = RegSetValueEx(hVersionKey, NULL, 0, REG_SZ, (PBYTE)g_szVersionValue, sizeof(g_szVersionValue));
    OnErrorJump(hr, lRet, done);

done:
    if (hInProcKey)
        RegCloseKey(hInProcKey);

    if (hClsidKey)
        RegCloseKey(hClsidKey);

    if (hTypelibKey)
        RegCloseKey(hTypelibKey);

    if (hVersionKey)
        RegCloseKey(hVersionKey);

    return hr;
}

STDAPI AddTypeLibIdEntry(LPOLESTR szTypelibId)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    TCHAR szTypelibIdRegKey[MAX_PATH] = { 0 };
    TCHAR szTypelib10RegKey[MAX_PATH] = { 0 };
    HKEY hTypelibIdKey = nullptr, hTypelib10Key = nullptr;

    hr = StringCbPrintf(szTypelibIdRegKey, sizeof(szTypelibIdRegKey), TEXT("Software\\Classes\\TypeLib\\%s"), szTypelibId);
    if (FAILED(hr))
        goto done;

    // HKCU\Software\Classes\TypeLib\TypeLibGUID
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szTypelibIdRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hTypelibIdKey, NULL);
    OnErrorJump(hr, lRet, done);

    hr = StringCbPrintf(szTypelib10RegKey, sizeof(szTypelib10RegKey), TEXT("%s\\1.0"), szTypelibIdRegKey);
    if (FAILED(hr))
        goto done;

    // HKCU\Software\Classes\TypeLib\TypeLibGUID\1.0
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szTypelib10RegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hTypelib10Key, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\TypeLib\TypeLibGUID\1.0 @=TypelibDescription
    lRet = RegSetValueEx(hTypelib10Key, NULL, 0, REG_SZ, (PBYTE)g_szTypelibDescription, sizeof(g_szTypelibDescription));
    OnErrorJump(hr, lRet, done);

done:
    if (hTypelibIdKey)
        RegCloseKey(hTypelibIdKey);

    if (hTypelib10Key)
        RegCloseKey(hTypelib10Key);

    return hr;
}

STDAPI AddInterfaceIdEntry(LPOLESTR szInterfaceId, LPOLESTR szTypelibId)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    TCHAR szInterfaceIdRegKey[MAX_PATH] = { 0 };
    TCHAR szTypelibRegKey[MAX_PATH] = { 0 };
    HKEY hInterfaceIdKey = nullptr, hTypelibKey = nullptr;

    hr = StringCbPrintf(szInterfaceIdRegKey, sizeof(szInterfaceIdRegKey), TEXT("Software\\Classes\\Interface\\%s"), szInterfaceId);
    if (FAILED(hr))
        goto done;

    // HKCU\Software\Classes\Interface\InterfaceGUID
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szInterfaceIdRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hInterfaceIdKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\Interface\InterfaceGUID @=InterfaceName
    lRet = RegSetValueEx(hInterfaceIdKey, NULL, 0, REG_SZ, (PBYTE)g_szInterfaceName, sizeof(g_szInterfaceName));
    OnErrorJump(hr, lRet, done);

    hr = StringCbPrintf(szTypelibRegKey, sizeof(szTypelibRegKey), TEXT("%s\\TypeLib"), szInterfaceIdRegKey);
    if (FAILED(hr))
        goto done;

    // HKCU\Software\Classes\Interface\InterfaceGUID\TypeLib
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szTypelibRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hTypelibKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\Interface\InterfaceGUID\TypeLib @=TypelibGuid
    lRet = RegSetValueEx(hTypelibKey, NULL, 0, REG_SZ, (PBYTE)szTypelibId, static_cast<DWORD>(_tcslen(szTypelibId) * sizeof(TCHAR)));
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\Interface\InterfaceGUID\TypeLib Version=1.0
    lRet = RegSetValueEx(hTypelibKey, g_szVersion, 0, REG_SZ, (PBYTE)g_szVersionValue, sizeof(g_szVersionValue));
    OnErrorJump(hr, lRet, done);

done:
    if (hInterfaceIdKey)
        RegCloseKey(hInterfaceIdKey);

    if (hTypelibKey)
        RegCloseKey(hTypelibKey);

    return hr;
}

STDAPI DllRegisterServer(void)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    LPOLESTR szClassId = nullptr, szTypelibId = nullptr, szInterfaceId = nullptr;
    TCHAR szFileName[MAX_PATH] = { 0 };

    LOG("Entering");

    CoInitialize(NULL);

    lRet = GetModuleFileName(nullptr, szFileName, sizeof(szFileName));
    if (lRet == 0) {
        hr = HRESULT_FROM_WIN32(GetLastError());
        goto done;
    }

    hr = StringFromCLSID(CLSID_AddObject, &szClassId);
    if (FAILED(hr))
        goto done;

    hr = StringFromCLSID(LIBID_SuperFastMathLib, &szTypelibId);
    if (FAILED(hr))
        goto done;

    hr = StringFromCLSID(IID_IAdd, &szInterfaceId);
    if (FAILED(hr))
        goto done;

    //
    // ProgID
    //
    hr = AddProgIdEntry(szClassId);
    if (FAILED(hr))
        goto done;

    //
    // CLSID
    //
    hr = AddClassIdEntry(szClassId, szTypelibId, szFileName);
    if (FAILED(hr))
        goto done;

    //
    // TypeLib ID
    //
    hr = AddTypeLibIdEntry(szTypelibId);
    if (FAILED(hr))
        goto done;

    //
    // Interface ID
    //
    hr = AddInterfaceIdEntry(szInterfaceId, szTypelibId);
    if (FAILED(hr))
        goto done;

done:
    if (szClassId)
        CoTaskMemFree(szClassId);

    if (szTypelibId)
        CoTaskMemFree(szTypelibId);

    if (szInterfaceId)
        CoTaskMemFree(szInterfaceId);

    CoUninitialize();

    LOG("Exiting with 0x%08x", hr);

    return hr;
}

STDAPI RemoveProgIdEntry()
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    HKEY hProgIdKey = nullptr;

    // HKCU\Software\Classes
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Classes"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hProgIdKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\ProgID
    lRet = RegDeleteTree(hProgIdKey, g_szProgId);
    OnErrorJump(hr, lRet, done);

done:
    if (hProgIdKey)
        RegCloseKey(hProgIdKey);

    return hr;
}

STDAPI RemoveClassIdEntry(LPOLESTR szClassId)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    HKEY hClsidKey = nullptr;

    // HKCU\Software\Classes\CLSID
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Classes\\CLSID"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hClsidKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\CLSID\ClassGUID
    lRet = RegDeleteTree(hClsidKey, szClassId);
    OnErrorJump(hr, lRet, done);

done:
    if (hClsidKey)
        RegCloseKey(hClsidKey);

    return hr;
}

STDAPI RemoveTypeLibIdEntry(LPOLESTR szTypelibId)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    HKEY hTypelibKey = nullptr;

    // HKCU\Software\Classes\TypeLib
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Classes\\TypeLib"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hTypelibKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\TypeLib\TypeLibGUID
    lRet = RegDeleteTree(hTypelibKey, szTypelibId);
    OnErrorJump(hr, lRet, done);

done:
    if (hTypelibKey)
        RegCloseKey(hTypelibKey);

    return hr;
}

STDAPI RemoveInterfaceIdEntry(LPOLESTR szInterfaceId)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    HKEY hInterfaceKey = nullptr;

    // HKCU\Software\Classes\Interface
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Classes\\Interface"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hInterfaceKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKCU\Software\Classes\Interface\InterfaceGUID
    lRet = RegDeleteTree(hInterfaceKey, szInterfaceId);
    OnErrorJump(hr, lRet, done);

done:
    if (hInterfaceKey)
        RegCloseKey(hInterfaceKey);

    return hr;
}

STDAPI DllUnregisterServer(void)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    LPOLESTR szClassId = nullptr, szTypelibId = nullptr, szInterfaceId = nullptr;

    LOG("Entering");

    CoInitialize(NULL);

    hr = StringFromCLSID(CLSID_AddObject, &szClassId);
    if (FAILED(hr))
        goto done;

    hr = StringFromCLSID(LIBID_SuperFastMathLib, &szTypelibId);
    if (FAILED(hr))
        goto done;

    hr = StringFromCLSID(IID_IAdd, &szInterfaceId);
    if (FAILED(hr))
        goto done;

    //
    // ProgID
    //
    hr = RemoveProgIdEntry();
    if (FAILED(hr))
        goto done;

    //
    // CLSID
    //
    hr = RemoveClassIdEntry(szClassId);
    if (FAILED(hr))
        goto done;

    //
    // TypeLib ID
    //
    hr = RemoveTypeLibIdEntry(szTypelibId);
    if (FAILED(hr))
        goto done;

    //
    // Interface ID
    //
    hr = RemoveInterfaceIdEntry(szInterfaceId);
    if (FAILED(hr))
        goto done;

done:
    if (szClassId)
        CoTaskMemFree(szClassId);

    if (szTypelibId)
        CoTaskMemFree(szTypelibId);

    if (szInterfaceId)
        CoTaskMemFree(szInterfaceId);

    CoUninitialize();

    LOG("Exiting with 0x%08x", hr);

    return hr;
}

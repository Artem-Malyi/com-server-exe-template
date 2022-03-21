//
// Registry.cpp
// Contains the implementation of the exported functions that are
// called by regsvr32.exe to register / unregister the COM component.
//
#include "..\common\pch.h"
#include "AddObj.h"
#include "AddObjFactory.h"
#include "AddObject_i.c" // Contains the interface IIDs
#include "pathcch.h"
#include "tchar.h"

#define DEBUG_LOGGER_ENABLED
#define FILE_LOGGER_ENABLED
#define LOG_PREFIX "[ADDOBJ-EXPORTS]"
#include "logger.h"

static const TCHAR g_szClassDescription[] = TEXT("SuperFast Addition COM Object");
static const TCHAR g_szProgId[] = TEXT("SuperFast.AddObj");
static const TCHAR g_szBoth[] = TEXT("Both");
static const TCHAR g_szThreadingModel[] = TEXT("ThreadingModel");
static const TCHAR g_szLocalServer[] = TEXT("LocalServer32");
static const TCHAR g_szTypeLib[] = TEXT("TypeLib");
static const TCHAR g_szProgrammable[] = TEXT("Programmable");
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
    HKEY hProgIdKey = NULL;

    hr = StringCbPrintf(szProgIdRegKey, sizeof(szProgIdRegKey), TEXT("Software\\Classes\\%s\\CLSID"), g_szProgId);
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\ProgId\CLSID
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szProgIdRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hProgIdKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\ProgId\CLSID @=ClassGUID
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
    HKEY hClsidKey = NULL, hLocalServerKey = NULL, hTypelibKey = NULL, hProgrammableKey = NULL, hVersionKey = NULL;

    hr = StringCbPrintf(szClassIdRegKey, sizeof(szClassIdRegKey), TEXT("Software\\Classes\\CLSID\\%s"), szClassId);
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\CLSID\ClassGUID
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szClassIdRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hClsidKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\CLSID\ClassGUID @=ClassDescription
    lRet = RegSetValueEx(hClsidKey, NULL, 0, REG_SZ, (PBYTE)g_szClassDescription, sizeof(g_szClassDescription));
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\CLSID\ClassGUID\LocalServer32
    lRet = RegCreateKeyEx(hClsidKey, g_szLocalServer, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hLocalServerKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\CLSID\ClassGUID\LocalServer32 @=PathToDll
    lRet = RegSetValueEx(hLocalServerKey, NULL, 0, REG_SZ, (PBYTE)szFileName, static_cast<DWORD>(_tcslen(szFileName) * sizeof(TCHAR)));
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\CLSID\ClassGUID\LocalServer32 ThreadingModel=Both
    lRet = RegSetValueEx(hLocalServerKey, g_szThreadingModel, 0, REG_SZ, (PBYTE)g_szBoth, sizeof(g_szBoth));
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\CLSID\ClassGUID\Programmable
    lRet = RegCreateKeyEx(hClsidKey, g_szProgrammable, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hProgrammableKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\CLSID\ClassGUID\TypeLib
    lRet = RegCreateKeyEx(hClsidKey, g_szTypeLib, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hTypelibKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\CLSID\ClassGUID\TypeLib @=TypelibGUID
    lRet = RegSetValueEx(hTypelibKey, NULL, 0, REG_SZ, (PBYTE)szTypelibId, static_cast<DWORD>(_tcslen(szTypelibId) * sizeof(TCHAR)));
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\CLSID\ClassGUID\Version
    lRet = RegCreateKeyEx(hClsidKey, g_szVersion, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hVersionKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\CLSID\ClassGUID\Version @=1.0
    lRet = RegSetValueEx(hVersionKey, NULL, 0, REG_SZ, (PBYTE)g_szVersionValue, sizeof(g_szVersionValue));
    OnErrorJump(hr, lRet, done);

done:
    if (hLocalServerKey)
        RegCloseKey(hLocalServerKey);

    if (hClsidKey)
        RegCloseKey(hClsidKey);

    if (hTypelibKey)
        RegCloseKey(hTypelibKey);

    if (hProgrammableKey)
        RegCloseKey(hProgrammableKey);

    if (hVersionKey)
        RegCloseKey(hVersionKey);

    return hr;
}

STDAPI AddTypeLibIdEntry(LPOLESTR szTypelibId, LPOLESTR szFilePath)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    TCHAR szTypelibIdKey[MAX_PATH] = { 0 }, szTypelibVersionKey[MAX_PATH] = { 0 }, szTypelibLocaleKey[MAX_PATH] = { 0 }, szFileDirectory[MAX_PATH] = { 0 };
    TCHAR szTypelibFlagsKey[MAX_PATH] = { 0 }, szTypelibHelpDirKey[MAX_PATH] = { 0 }, szTypelibPlatformKey[MAX_PATH] = { 0 };
    HKEY hTypelibIdKey = NULL, hTypelibVersionKey = NULL, hTypelibLocaleKey = NULL, hTypelibFlagsKey = NULL, hTypelibHelpDirKey = NULL, hTypelibPlatformKey = NULL;

    hr = StringCbPrintf(szFileDirectory, sizeof(szFileDirectory), TEXT("%s"), szFilePath);
    if (FAILED(hr))
        goto done;

    hr = PathCchRemoveFileSpec(szFileDirectory, sizeof(szFileDirectory) / sizeof(TCHAR));
    if (FAILED(hr))
        goto done;

    hr = StringCbPrintf(szTypelibIdKey, sizeof(szTypelibIdKey), TEXT("Software\\Classes\\TypeLib\\%s"), szTypelibId);
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\TypeLib\TypeLibGUID
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szTypelibIdKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hTypelibIdKey, NULL);
    OnErrorJump(hr, lRet, done);

    hr = StringCbPrintf(szTypelibVersionKey, sizeof(szTypelibVersionKey), TEXT("%s\\1.0"), szTypelibIdKey);
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\TypeLib\TypeLibGUID\1.0
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szTypelibVersionKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hTypelibVersionKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\TypeLib\TypeLibGUID\1.0 @=TypelibDescription
    lRet = RegSetValueEx(hTypelibVersionKey, NULL, 0, REG_SZ, (PBYTE)g_szTypelibDescription, sizeof(g_szTypelibDescription));
    OnErrorJump(hr, lRet, done);

    hr = StringCbPrintf(szTypelibLocaleKey, sizeof(szTypelibLocaleKey), TEXT("%s\\0"), szTypelibVersionKey); // 0 is for en_US locale
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\TypeLib\TypeLibGUID\1.0\0
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szTypelibLocaleKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hTypelibLocaleKey, NULL);
    OnErrorJump(hr, lRet, done);

    hr = StringCbPrintf(szTypelibPlatformKey, sizeof(szTypelibPlatformKey), TEXT("%s\\win64"), szTypelibLocaleKey); // TODO: check on win32
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\TypeLib\TypeLibGUID\1.0\0\win64
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szTypelibPlatformKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hTypelibPlatformKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\TypeLib\TypeLibGUID\1.0\0\win64 @=PathToLocalServerExe
    lRet = RegSetValueEx(hTypelibPlatformKey, NULL, 0, REG_SZ, (PBYTE)szFilePath, static_cast<DWORD>(_tcslen(szFilePath) * sizeof(TCHAR)));
    OnErrorJump(hr, lRet, done);

    hr = StringCbPrintf(szTypelibFlagsKey, sizeof(szTypelibFlagsKey), TEXT("%s\\FLAGS"), szTypelibVersionKey);
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\TypeLib\TypeLibGUID\1.0\FLAGS
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szTypelibFlagsKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hTypelibFlagsKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\TypeLib\TypeLibGUID\1.0\FLAGS @=Flags
    lRet = RegSetValueEx(hTypelibFlagsKey, NULL, 0, REG_SZ, (PBYTE)L"0", sizeof(TCHAR));
    OnErrorJump(hr, lRet, done);

    hr = StringCbPrintf(szTypelibHelpDirKey, sizeof(szTypelibHelpDirKey), TEXT("%s\\HELPDIR"), szTypelibVersionKey);
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\TypeLib\TypeLibGUID\1.0\HELPDIR
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szTypelibHelpDirKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hTypelibHelpDirKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\TypeLib\TypeLibGUID\1.0\HELPDIR @=PathToLocalServerDirectory
    lRet = RegSetValueEx(hTypelibHelpDirKey, NULL, 0, REG_SZ, (PBYTE)szFileDirectory, sizeof(szFileDirectory));
    OnErrorJump(hr, lRet, done);

done:
    if (hTypelibIdKey)
        RegCloseKey(hTypelibIdKey);

    if (hTypelibVersionKey)
        RegCloseKey(hTypelibVersionKey);

    if (hTypelibLocaleKey)
        RegCloseKey(hTypelibLocaleKey);

    if (hTypelibFlagsKey)
        RegCloseKey(hTypelibFlagsKey);

    if (hTypelibHelpDirKey)
        RegCloseKey(hTypelibHelpDirKey);

    if (hTypelibPlatformKey)
        RegCloseKey(hTypelibPlatformKey);

    return hr;
}

STDAPI AddInterfaceIdEntry(LPOLESTR szInterfaceId, LPOLESTR szTypelibId)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    TCHAR szInterfaceIdRegKey[MAX_PATH] = { 0 }, szTypelibRegKey[MAX_PATH] = { 0 }, szProxyStubKey[MAX_PATH] = { 0 };
    HKEY hInterfaceIdKey = NULL, hTypelibKey = NULL, hProxyStubKey = NULL;

    // Use PSOAInterface proxy/stub for CLR COM-visible classes.
    WCHAR proxyClsid[] = TEXT("{00020424-0000-0000-C000-000000000046}");

    hr = StringCbPrintf(szInterfaceIdRegKey, sizeof(szInterfaceIdRegKey), TEXT("Software\\Classes\\Interface\\%s"), szInterfaceId);
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\Interface\InterfaceGUID
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szInterfaceIdRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hInterfaceIdKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\Interface\InterfaceGUID @=InterfaceName
    lRet = RegSetValueEx(hInterfaceIdKey, NULL, 0, REG_SZ, (PBYTE)g_szInterfaceName, sizeof(g_szInterfaceName));
    OnErrorJump(hr, lRet, done);

    hr = StringCbPrintf(szTypelibRegKey, sizeof(szTypelibRegKey), TEXT("%s\\TypeLib"), szInterfaceIdRegKey);
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\Interface\InterfaceGUID\TypeLib
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szTypelibRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hTypelibKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\Interface\InterfaceGUID\TypeLib @=TypelibGuid
    lRet = RegSetValueEx(hTypelibKey, NULL, 0, REG_SZ, (PBYTE)szTypelibId, static_cast<DWORD>(_tcslen(szTypelibId) * sizeof(TCHAR)));
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\Interface\InterfaceGUID\TypeLib Version=1.0
    lRet = RegSetValueEx(hTypelibKey, g_szVersion, 0, REG_SZ, (PBYTE)g_szVersionValue, sizeof(g_szVersionValue));
    OnErrorJump(hr, lRet, done);

    hr = StringCbPrintf(szProxyStubKey, sizeof(szProxyStubKey), TEXT("%s\\ProxyStubClsid32"), szInterfaceIdRegKey);
    if (FAILED(hr))
        goto done;

    // HKLM\Software\Classes\Interface\InterfaceGUID\ProxyStubClsid32
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szProxyStubKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hProxyStubKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\Interface\InterfaceGUID\ProxyStubClsid32 @=ProxyCLSID
    lRet = RegSetValueEx(hProxyStubKey, NULL, 0, REG_SZ, (PBYTE)proxyClsid, sizeof(proxyClsid));
    OnErrorJump(hr, lRet, done);

done:
    if (hInterfaceIdKey)
        RegCloseKey(hInterfaceIdKey);

    if (hTypelibKey)
        RegCloseKey(hTypelibKey);

    if (hProxyStubKey)
        RegCloseKey(hProxyStubKey);

    return hr;
}

STDAPI DllRegisterServer(void)
{
    HRESULT hr = S_OK;
    LONG lRet = ERROR_SUCCESS;
    LPOLESTR szClassId = NULL, szTypelibId = NULL, szInterfaceId = NULL;
    TCHAR szFileName[MAX_PATH] = { 0 };

    LOG("Entering");

    CoInitialize(NULL);

    lRet = GetModuleFileName(NULL, szFileName, sizeof(szFileName));
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
    hr = AddTypeLibIdEntry(szTypelibId, szFileName);
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
    HKEY hProgIdKey = NULL;

    // HKLM\Software\Classes
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Classes"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hProgIdKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\ProgID
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
    HKEY hClsidKey = NULL;

    // HKLM\Software\Classes\CLSID
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Classes\\CLSID"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hClsidKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\CLSID\ClassGUID
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
    HKEY hTypelibKey = NULL;

    // HKLM\Software\Classes\TypeLib
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Classes\\TypeLib"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hTypelibKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\TypeLib\TypeLibGUID
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
    HKEY hInterfaceKey = NULL;

    // HKLM\Software\Classes\Interface
    lRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Classes\\Interface"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hInterfaceKey, NULL);
    OnErrorJump(hr, lRet, done);

    // HKLM\Software\Classes\Interface\InterfaceGUID
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
    LPOLESTR szClassId = NULL, szTypelibId = NULL, szInterfaceId = NULL;

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

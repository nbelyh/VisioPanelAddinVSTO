
#include "stdafx.h"

LPCWSTR wszVisioRegistryConfigKeyPath = L"Software\\Microsoft\\Office\\Visio";
LPCWSTR wszVisioRegistryConfigValueName = L"ConfigChangeID";

HRESULT ImplUpdateVisioConfigChangeID(MSIHANDLE hInstall, DWORD flags)
{
	HRESULT hr = S_OK;
	WcaLog(LOGMSG_VERBOSE, "Resetting visio template and stencil cache.");

	HKEY hKey = NULL;
	hr = RegOpen(HKEY_LOCAL_MACHINE, wszVisioRegistryConfigKeyPath, KEY_READ|KEY_WRITE|flags, &hKey);
	if (E_FILENOTFOUND == hr)
		return S_FALSE;

	ExitOnFailure(hr, "Failed to open the registry key: %ls. It looks like Visio is not installed or security issue. Visio cache will not be reset.", wszVisioRegistryConfigKeyPath);

	WcaLog(LOGMSG_VERBOSE, "Opened registry key: %ls", wszVisioRegistryConfigKeyPath);

	DWORD config_change_id = 0;
	hr = RegReadNumber(hKey, wszVisioRegistryConfigValueName, &config_change_id);
	ExitOnFailure(hr, "Failed read value of key %s. It looks like Visio is not installed or security issue. Visio cache will not be reset.", wszVisioRegistryConfigValueName);

	WcaLog(LOGMSG_VERBOSE, "Read value of key %s (%d)",  wszVisioRegistryConfigValueName, config_change_id);

	if (MsiGetMode(hInstall, MSIRUNMODE_ROLLBACK))
		--config_change_id;
	else
		++config_change_id;

	hr = RegWriteNumber(hKey, wszVisioRegistryConfigValueName, config_change_id);
	ExitOnFailure(hr, "Failed set new value for key %s (%d). It looks like some security issue. Visio cache will not be reset.", wszVisioRegistryConfigValueName, config_change_id);

	WcaLog(LOGMSG_VERBOSE, "Set new value for the key %s (%d)", wszVisioRegistryConfigValueName, config_change_id);

LExit:
	if (hKey) RegCloseKey(hKey);
	return hr;
}


UINT __stdcall UpdateVisioConfigChangeID32(MSIHANDLE hInstall)
{
	HRESULT hr = WcaInitialize(hInstall, "UpdateVisioConfigChangeID32");
	ExitOnFailure(hr, "Failed to initialize");

	hr = ImplUpdateVisioConfigChangeID(hInstall, KEY_WOW64_32KEY);
	ExitOnFailure(hr, "Unable to reset visio 64 addin cache")

LExit:
	return WcaFinalize(SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE);
}

UINT __stdcall UpdateVisioConfigChangeID64(MSIHANDLE hInstall)
{
	HRESULT hr = WcaInitialize(hInstall, "UpdateVisioConfigChangeID64");
	ExitOnFailure(hr, "Failed to initialize");

	hr = ImplUpdateVisioConfigChangeID(hInstall, KEY_WOW64_64KEY);
	ExitOnFailure(hr, "Unable to reset visio 64 addin cache")

LExit:
	return WcaFinalize(SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE);
}

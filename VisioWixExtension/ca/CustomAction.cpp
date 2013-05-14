
#include "stdafx.h"

BOOL IsWow64()
{
	//IsWow64Process is not available on all supported versions of Windows.
	//Use GetModuleHandle to get a handle to the DLL that contains the function
	//and GetProcAddress to get a pointer to the function if available.

	typedef BOOL (WINAPI *FN) (HANDLE, PBOOL);
	FN fn = (FN) GetProcAddress(GetModuleHandle(TEXT("kernel32")),"IsWow64Process");

	BOOL bIsWow64 = FALSE;
	
	if (fn)
		fn(GetCurrentProcess(),&bIsWow64);

	return bIsWow64;
}

void DisableWow64FsRedirection(LPVOID *OldValue)
{
	typedef BOOL (WINAPI *FN) (LPVOID*);
	FN fn = (FN) GetProcAddress(GetModuleHandle(TEXT("kernel32")),"Wow64DisableWow64FsRedirection");

	if (fn)
		fn(OldValue);
};

void RevertWow64FsRedirection(LPVOID OldValue)
{
	typedef BOOL (WINAPI *FN) (LPVOID);
	FN fn = (FN) GetProcAddress(GetModuleHandle(TEXT("kernel32")),"Wow64RevertWow64FsRedirection");

	if (fn)
		fn(OldValue);
}

UINT __stdcall UpdateVisioConfigChangeID(MSIHANDLE hInstall)
{
	HRESULT hr_init = WcaInitialize(hInstall, "UpdateVisioConfigChangeID");
	ExitOnFailure(hr_init, "Failed to initialize");

	WcaLog(LOGMSG_STANDARD, "Started Resetting visio template and stencil cache.");

	HKEY key = NULL;
	HRESULT hr = HRESULT_FROM_WIN32(RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Office\\Visio", 0, KEY_READ|KEY_WRITE, &key));
	if (FAILED(hr))
	{
		WcaLogError(hr, "Failed to open the registry key: Software\\Microsoft\\Office\\Visio. It looks like Visio is not installed or security issue. Visio cache will not be reset.");
		WcaFinalize(ERROR_SUCCESS);
	}
	else
	{
		WcaLog(LOGMSG_VERBOSE, "Successfully opened registry key: Software\\Microsoft\\Office\\Visio");

		DWORD config_change_id = 1;
		DWORD config_change_id_len = sizeof(config_change_id);
		DWORD key_type = REG_DWORD;
		hr = HRESULT_FROM_WIN32(RegQueryValueEx(key, L"ConfigChangeID", NULL, &key_type, (LPBYTE)&config_change_id, &config_change_id_len));
		if (FAILED(hr))
		{
			WcaLogError(hr, "Failed read old ConfigChangeID key value of key: Software\\Microsoft\\Office\\Visio. It looks like Visio is not installed or security issue. Visio cache will not be reset.");
		}
		else
		{
			WcaLog(LOGMSG_VERBOSE, "Successfully read old ConfigChangeID value (%d)", config_change_id);

			++config_change_id;

			hr = HRESULT_FROM_WIN32(RegSetValueEx(key, L"ConfigChangeID", NULL, key_type, (LPBYTE)&config_change_id, config_change_id_len));
			if (FAILED(hr))
			{
				WcaLogError(hr, "Failed set new ConfigChangeID key value for key: Software\\Microsoft\\Office\\Visio. It looks like some security issue. Visio cache will not be reset.");
			}
			else
			{
				WcaLog(LOGMSG_VERBOSE, "Successfully set new ConfigChangeID value (%d)", config_change_id);
			}
		}
	}
LExit:
	UINT er = SUCCEEDED(hr_init) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE;
	return WcaFinalize(er);
}

UINT __stdcall UpdateVisioConfigChangeID64(MSIHANDLE hInstall)
{
	BOOL fWOW = IsWow64();

	PVOID OldValue = NULL;

	UINT result = 0;

	if (fWOW)
	{
		DisableWow64FsRedirection(&OldValue);
		result = UpdateVisioConfigChangeID(hInstall);
		RevertWow64FsRedirection(OldValue);
	}
	else
	{
		WcaLog(LOGMSG_STANDARD, "It looks like the system is a 32-bit system, therefore skipping Visio x64 cache reset.");
	}

	return result;
}


// DllMain - Initialize and cleanup WiX custom action utils.
extern "C" BOOL WINAPI DllMain(
	__in HINSTANCE hInst,
	__in ULONG ulReason,
	__in LPVOID
	)
{
	switch(ulReason)
	{
	case DLL_PROCESS_ATTACH:
		WcaGlobalInitialize(hInst);
		break;

	case DLL_PROCESS_DETACH:
		WcaGlobalFinalize();
		break;
	}

	return TRUE;
}

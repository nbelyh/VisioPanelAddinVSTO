
#include "stdafx.h"
#include "Register.h"

/***************************************************************/

DWORD GetSAM(DWORD sam, REG_KEY_BITNESS iBitness)
{
	if (iBitness == REG_KEY_32BIT)
		sam |= KEY_WOW64_32KEY;

	if (iBitness == REG_KEY_64BIT)
		sam |= KEY_WOW64_64KEY;

	return sam;
}

HRESULT CreateOfficeRegistryKeyAtRoot(
	LPCWSTR pwzId, 
	LPCWSTR pwzFile, 
	LPCWSTR pwzFriendlyName, 
	LPCWSTR pwzDescription,
	int iCommandLineSafe,
	int iLoadBehavior,
	int iAddinType,
	HKEY hKeyRoot, 
	REG_KEY_BITNESS iBitness)
{
	HRESULT hr = S_OK;
	HKEY hKey = NULL;

	LPWSTR wszRegPath = NULL;
	LPWSTR wszFileEntry = NULL;

    WcaLog(LOGMSG_VERBOSE, "CreateOfficeRegistryKey: Bitness=%ld, Id=%ls, Manifest=%ls, FriendlyName=%ls, Description=%ls", 
		iBitness, pwzId, pwzFile, pwzFriendlyName, pwzDescription);

    hr = StrAllocFormatted(&wszRegPath, L"Software\\Microsoft\\Visio\\Addins\\%ls", pwzId);
    ExitOnFailure(hr, "Failed to allocate string for registry path.");

	hr = RegCreateEx(hKeyRoot, wszRegPath, GetSAM(KEY_READ|KEY_WRITE, iBitness), FALSE, NULL, &hKey, NULL);
	ExitOnFailure(hr, "Failed to create or open the registry key: %ls. It looks like Visio is not installed or security issue", wszRegPath);

	WcaLog(LOGMSG_VERBOSE, "Created or opened registry key: %ls", wszRegPath);

	if (iAddinType == ADDIN_TYPE_VSTO)
	{
		hr = StrAllocFormatted(&wszFileEntry, L"file:///%ls|vstolocal", pwzFile);
		ExitOnFailure(hr, "Failed to allocate string for registry file entry.");

		hr = RegWriteString(hKey, L"Manifest", wszFileEntry);
		ExitOnFailure(hr, "Failed set Manifest value.");
	}

	hr = RegWriteString(hKey, L"FriendlyName", pwzFriendlyName);
	ExitOnFailure(hr, "Failed set FriendlyName value.");

	hr = RegWriteString(hKey, L"Description", pwzDescription);
	ExitOnFailure(hr, "Failed set Description value.");

	hr = RegWriteNumber(hKey, L"CommandLineSafe", iCommandLineSafe);
	ExitOnFailure(hr, "Failed set CommandLineSafe value value.");

	hr = RegWriteNumber(hKey, L"LoadBehavior", iLoadBehavior);
	ExitOnFailure(hr, "Failed set LoadBehavior value value.");

LExit:
	ReleaseStr(wszFileEntry);
	ReleaseStr(wszRegPath);
	ReleaseRegKey(hKey);
	return hr;
}

HRESULT DeleteOfficeRegistryKeyAtRoot(
	LPCWSTR pwzProgId, 
	HKEY hKeyRoot, 
	REG_KEY_BITNESS iBitness)
{
	HRESULT hr = S_OK;

	LPWSTR wszRegPathAddinProgId = NULL;
	LPWSTR wszRegPathAddins = NULL;

    WcaLog(LOGMSG_VERBOSE, "DeleteOfficeRegistryKey: Bitness:%ld, Id=%ls", 
		iBitness, pwzProgId);

    hr = StrAllocFormatted(&wszRegPathAddinProgId, L"Software\\Microsoft\\Visio\\Addins\\%ls", pwzProgId);
    ExitOnFailure(hr, "Failed to allocate registry path.");

	hr = RegDelete(hKeyRoot, wszRegPathAddinProgId, iBitness, FALSE);
	WcaLogError(hr, "Failed to delete the registry key: %ls.", wszRegPathAddinProgId);

    hr = StrAllocFormatted(&wszRegPathAddins, L"Software\\Microsoft\\Visio\\Addins");
    ExitOnFailure(hr, "Failed to allocate registry path.");

	hr = RegDelete(hKeyRoot, wszRegPathAddins, iBitness, FALSE);
	WcaLogError(hr, "Failed to delete the registry key: %ls.", wszRegPathAddins);

LExit:
	ReleaseStr(wszRegPathAddinProgId);
	ReleaseStr(wszRegPathAddins);
	return hr;
}

HRESULT RegisterCOMAtRoot(
    LPCWSTR strClsId,
    LPCWSTR strProgId,
    LPCWSTR typeFullName, 
    LPCWSTR strAsmName, 
    LPCWSTR strAsmVersion, 
    LPCWSTR strAsmCodeBase, 
    LPCWSTR strRuntimeVersion,
	HKEY hKeyRoot,
	REG_KEY_BITNESS iBitness)
{
	HRESULT hr = S_OK;

	HKEY hKeyProgId = NULL;
	HKEY hKeyProgIdClassId = NULL;
	HKEY hKeyClassIdProgId = NULL;
	HKEY hKeyClsId = NULL;
	HKEY hKeyClsIdInprocServer = NULL;
	HKEY hKeyClsIdInprocServerVersion = NULL;

	LPWSTR wszClsId = NULL;
	LPWSTR wszProgId = NULL;

	hr = StrAllocFormatted(&wszClsId, L"Software\\Classes\\CLSID\\%ls", strClsId);
	ExitOnFailure(hr, "Failed to allocate string for registry path.");

	// Software\Classes\CLSID\{XX-XX-XX}
	hr = RegCreateEx(hKeyRoot, wszClsId, GetSAM(KEY_READ|KEY_WRITE, iBitness), FALSE, NULL, &hKeyClsId, NULL);
	ExitOnFailure(hr, "Failed to create or open the registry key: %ls", wszClsId);

	// Software\Classes\CLSID\{XX-XX-XX}\@
	hr = RegWriteString(hKeyClsId, L"", typeFullName);
	ExitOnFailure1(hr, "Failed to write to %ls", wszClsId);

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32
	hr = RegCreateEx(hKeyClsId, L"InprocServer32", KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyClsIdInprocServer, NULL);
	ExitOnFailure(hr, "Failed to create key %ls\\InprocServer32", wszClsId);

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\\@
	hr = RegWriteString(hKeyClsIdInprocServer, L"", L"mscoree.dll");
	ExitOnFailure(hr, "Failed to write to %ls", wszClsId);

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\\@ThreadingModel
	hr = RegWriteString(hKeyClsIdInprocServer, L"ThreadingModel", L"both");
	ExitOnFailure(hr, "Failed to write to %ls", wszClsId);

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\\@Class
	hr = RegWriteString(hKeyClsIdInprocServer, L"Class", typeFullName);
	ExitOnFailure(hr, "Failed to write to %ls", wszClsId);

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\\@Assembly
	hr = RegWriteString(hKeyClsIdInprocServer, L"Assembly", strAsmName);
	ExitOnFailure(hr, "Failed to write to %ls", wszClsId);

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\\@RuntimeVersion
	hr = RegWriteString(hKeyClsIdInprocServer, L"RuntimeVersion", strRuntimeVersion);
	ExitOnFailure(hr, "Failed to write to %ls", wszClsId);

	if (strAsmCodeBase && *strAsmCodeBase)
	{
		// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\\@CodeBase
		hr = RegWriteString(hKeyClsIdInprocServer, L"CodeBase", strAsmCodeBase);
		ExitOnFailure(hr, "Failed to write to %ls", wszClsId);
	}
	
	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\1.0.0.0
	hr = RegCreateEx(hKeyClsIdInprocServer, strAsmVersion, KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyClsIdInprocServerVersion, NULL);
	ExitOnFailure(hr, "Failed to create key %ls\\%ls", wszClsId, strAsmVersion);

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\1.0.0.0\@Class
	hr = RegWriteString(hKeyClsIdInprocServerVersion, L"Class", typeFullName);
	ExitOnFailure(hr, "Failed to write to %ls", wszClsId);

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\1.0.0.0\@Assembly
	hr = RegWriteString(hKeyClsIdInprocServerVersion, L"Assembly", strAsmName);
	ExitOnFailure(hr, "Failed to write to %ls", wszClsId);

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\1.0.0.0\@RuntimeVersion
	hr = RegWriteString(hKeyClsIdInprocServerVersion, L"RuntimeVersion", strRuntimeVersion);
	ExitOnFailure(hr, "Failed to write to %ls", wszClsId);

	if (strAsmCodeBase && *strAsmCodeBase)
	{
		// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\1.0.0.0\@CodeBase
		hr = RegWriteString(hKeyClsIdInprocServerVersion, L"CodeBase", strAsmCodeBase);
		ExitOnFailure(hr, "Failed to write to %ls", wszClsId);
	}

	if (strProgId && *strProgId)
	{
		hr = StrAllocFormatted(&wszProgId, L"Software\\Classes\\%ls", strProgId);
		ExitOnFailure(hr, "Failed to allocate string for registry path.");

		// Software\Classes\Prog.Id
		hr = RegCreateEx(hKeyRoot, wszProgId, KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyProgId, NULL);
		ExitOnFailure(hr, "Failed to create or open the registry key: %ls", wszProgId);

		// Software\Classes\Prog.Id\\@
		hr = RegWriteString(hKeyProgId, L"", typeFullName);
		ExitOnFailure1(hr, "Failed to write to %ls", wszProgId);

		// Software\Classes\Prog.Id\CLSID
		hr = RegCreateEx(hKeyProgId, L"CLSID", KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyProgIdClassId, NULL);
		ExitOnFailure1(hr, "Failed to create key %ls\\CLSID", wszProgId);

		// Software\Classes\Prog.Id\CLSID\@
		hr = RegWriteString(hKeyProgIdClassId, L"", strClsId);
		ExitOnFailure1(hr, "Failed to write to %ls\\CLSID", wszProgId);

		// Software\Classes\CLSID\{XX-XX-XX}\ProgId
		hr = RegCreateEx(hKeyClsId, L"ProgId", KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyClassIdProgId, NULL);
		ExitOnFailure(hr, "Failed to create key %ls\\ProgId", wszClsId);

		// Software\Classes\CLSID\{XX-XX-XX}\ProgId\@
		hr = RegWriteString(hKeyClassIdProgId, L"", strProgId);
		ExitOnFailure(hr, "Failed to write to key %ls\\ProgId", wszClsId);
	}

	// Software\Classes\CLSID\{XX-XX-XX}\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}
	hr = RegCreateEx(hKeyClsId, L"Implemented Categories\\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}", KEY_READ|KEY_WRITE, FALSE, NULL, &hKeyClsIdInprocServer, NULL);
	ExitOnFailure(hr, "Failed to create key", wszClsId);

    // EnsureManagedCategoryExists();

LExit:
	ReleaseStr(wszClsId);
	ReleaseStr(wszProgId);

	ReleaseRegKey(hKeyProgId);
	ReleaseRegKey(hKeyProgIdClassId);
	ReleaseRegKey(hKeyClassIdProgId);

	ReleaseRegKey(hKeyClsIdInprocServer);
	ReleaseRegKey(hKeyClsIdInprocServerVersion);
	ReleaseRegKey(hKeyClsId);

	return hr;
}

BOOL RegKeyIsEmpty(HKEY hKeyRoot, REG_KEY_BITNESS iBitness, LPCWSTR strPath)
{
	HKEY hKey = NULL;
	LPWSTR wszKeyName = NULL;
	
	HRESULT hr = RegOpen(hKeyRoot, strPath, GetSAM(KEY_READ, iBitness), &hKey);
	ExitOnFailure(hr, "Unable to determine if key has sub-keys");

    hr = RegKeyEnum(hKey, 0, &wszKeyName);
    
LExit:
	ReleaseStr(wszKeyName);
	ReleaseRegKey(hKey);

	return (hr == E_NOMOREITEMS);
}

HRESULT UnregisterCOMAtRoot(
	LPCWSTR strClsId,
    LPCWSTR strProgId,
	LPCWSTR strAsmVersion,
	HKEY hKeyRoot,
	REG_KEY_BITNESS iBitness)
{
	HRESULT hr = S_OK;

	LPWSTR wszProgId = NULL;
	LPWSTR wszProgIdClsId = NULL;

	LPWSTR wszClsId = NULL;
	LPWSTR wszClsIdInprocServer = NULL;
	LPWSTR wszClsIdProgId = NULL;
	LPWSTR wszClsIdInprocServerVersion = NULL;

	LPWSTR wszClsIdIdImplementedCategoriesCategoryId = NULL;
	LPWSTR wszClsIdIdImplementedCategories = NULL;

	hr = StrAllocFormatted(&wszClsId, L"Software\\Classes\\CLSID\\%ls", strClsId);
	ExitOnFailure(hr, "Failed to allocate string for registry path.");

	hr = StrAllocFormatted(&wszClsIdInprocServer, L"%ls\\InprocServer32", wszClsId);
	ExitOnFailure(hr, "Failed to allocate string for registry path.");

	hr = StrAllocFormatted(&wszClsIdInprocServerVersion, L"%ls\\%ls", wszClsIdInprocServer, strAsmVersion);
	ExitOnFailure(hr, "Failed to allocate string for registry path.");

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32\1.0.0.0
	hr = RegDelete(hKeyRoot, wszClsIdInprocServerVersion, iBitness, FALSE);
	ExitOnFailure(hr, "Failed to deletethe registry key %ls", wszClsIdInprocServerVersion);

	if (!RegKeyIsEmpty(hKeyRoot, iBitness, wszClsIdInprocServer))
		goto LExit;

	// Software\Classes\CLSID\{XX-XX-XX}\InprocServer32
	hr = RegDelete(hKeyRoot, wszClsIdInprocServer, iBitness, FALSE);
	WcaLogError(hr, "Failed to deletethe registry key %ls", wszClsIdInprocServer);

    if (strProgId && *strProgId)
    {
		hr = StrAllocFormatted(&wszProgIdClsId, L"Software\\Classes\\%ls\\CLSID", strProgId);
		ExitOnFailure(hr, "Failed to allocate string for registry path.");

		// Software\Classes\Prog.Id\\CLSID
		hr = RegDelete(hKeyRoot, wszProgIdClsId, REG_KEY_DEFAULT, FALSE);
		WcaLogError(hr, "Failed to allocate string for registry path.");

		hr = StrAllocFormatted(&wszProgId, L"Software\\Classes\\%ls", strProgId);
		ExitOnFailure(hr, "Failed to allocate string for registry path.");

		// Software\Classes\Prog.Id
		hr = RegDelete(hKeyRoot, wszProgId, REG_KEY_DEFAULT, FALSE);
		WcaLogError(hr, "Failed to allocate string for registry path.");

		hr = StrAllocFormatted(&wszClsIdProgId, L"%ls\\ProgId", wszClsId);
		ExitOnFailure(hr, "Failed to allocate string for registry path.");

		// Software\Classes\CLSID\{XX-XX-XX}\ProgId
		hr = RegDelete(hKeyRoot, wszClsIdProgId, iBitness, FALSE);
		WcaLogError(hr, "Failed to deletethe registry key %ls", wszClsIdProgId);
    }

	hr = StrAllocFormatted(&wszClsIdIdImplementedCategoriesCategoryId, L"%ls\\Implemented Categories\\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}", wszClsId);
	ExitOnFailure(hr, "Failed to allocate string for registry path.");

	// Software\Classes\CLSID\{XX-XX-XX}\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}
	hr = RegDelete(hKeyRoot, wszClsIdIdImplementedCategoriesCategoryId, iBitness, FALSE);
	WcaLogError(hr, "Failed to deletethe registry key %ls", wszClsId);

	hr = StrAllocFormatted(&wszClsIdIdImplementedCategories, L"%ls\\Implemented Categories", wszClsId);
	ExitOnFailure(hr, "Failed to allocate string for registry path.");

	// Software\Classes\CLSID\{XX-XX-XX}\Implemented Categories
	hr = RegDelete(hKeyRoot, wszClsIdIdImplementedCategories, iBitness, FALSE);
	WcaLogError(hr, "Failed to deletethe registry key %ls", wszClsId);

	// Software\Classes\CLSID\{XX-XX-XX}
	hr = RegDelete(hKeyRoot, wszClsId, iBitness, FALSE);
	WcaLogError(hr, "Failed to deletethe registry key %ls", wszClsId);

LExit:

	ReleaseStr(wszProgId);
	ReleaseStr(wszProgIdClsId);
	ReleaseStr(wszClsIdInprocServer);
	ReleaseStr(wszClsIdInprocServerVersion);

	return hr;
}

HRESULT CreateOfficeRegistryKey(
	LPCWSTR pwzId, 
	LPCWSTR pwzFile, 
	LPCWSTR pwzFriendlyName, 
	LPCWSTR pwzDescription,
	int iCommandLineSafe,
	int iLoadBehavior,
	int iAddinType,
	BOOL fPerUserInstall, 
	int iBitness)
{
	HRESULT hr = S_OK;

    hr = RegInitialize();
    ExitOnFailure(hr, "Failed to initialize the registry functions.");

	if (fPerUserInstall)
	{
		hr = CreateOfficeRegistryKeyAtRoot(pwzId, pwzFile, pwzFriendlyName, pwzDescription, iCommandLineSafe, iLoadBehavior, iAddinType, HKEY_CURRENT_USER, REG_KEY_DEFAULT);
		ExitOnFailure1(hr, "failed to register addin (HKCU): %ls", pwzId);
	}
	else
	{
		if (iBitness == REG_KEY_32BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = CreateOfficeRegistryKeyAtRoot(pwzId, pwzFile, pwzFriendlyName, pwzDescription, iCommandLineSafe, iLoadBehavior, iAddinType, HKEY_LOCAL_MACHINE, REG_KEY_32BIT);
			ExitOnFailure1(hr, "failed to register addin (HKLM, 32bit): %ls", pwzId);
		}
		if (iBitness == REG_KEY_64BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = CreateOfficeRegistryKeyAtRoot(pwzId, pwzFile, pwzFriendlyName, pwzDescription, iCommandLineSafe, iLoadBehavior, iAddinType, HKEY_LOCAL_MACHINE, REG_KEY_64BIT);
			ExitOnFailure1(hr, "failed to register addin (HKLM, 64bit): %ls", pwzId);
		}
	}

LExit:
	RegUninitialize();

	return hr;
}

HRESULT DeleteOfficeRegistryKey(
	LPCWSTR pwzId, 
	BOOL fPerUserInstall, 
	int iBitness)
{
	HRESULT hr = S_OK;

    hr = RegInitialize();
    ExitOnFailure(hr, "Failed to initialize the registry functions.");

	if (fPerUserInstall)
	{
		hr = DeleteOfficeRegistryKeyAtRoot(pwzId, HKEY_CURRENT_USER, REG_KEY_DEFAULT);
		WcaLogError(hr, "failed to unregister addin (HKCU): %ls", pwzId);
	}
	else
	{
		if (iBitness == REG_KEY_32BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = DeleteOfficeRegistryKeyAtRoot(pwzId, HKEY_LOCAL_MACHINE, REG_KEY_32BIT);
			WcaLogError(hr, "failed to unregister addin (HKLM, 32bit): %ls", pwzId);
		}
		if (iBitness == REG_KEY_64BIT || iBitness == REG_KEY_DEFAULT)
		{
			hr = DeleteOfficeRegistryKeyAtRoot(pwzId, HKEY_LOCAL_MACHINE, REG_KEY_64BIT);
			WcaLogError(hr, "failed to unregister addin (HKLM, 64bit): %ls", pwzId);
		}
	}

LExit:
	RegUninitialize();

	return hr;
}

HRESULT __stdcall RegisterCOM(
    LPCWSTR strClsId,
    LPCWSTR strProgId,
    LPCWSTR typeFullName, 
    LPCWSTR strAsmName, 
    LPCWSTR strAsmVersion, 
    LPCWSTR strAsmPath, 
    LPCWSTR strRuntimeVersion,
	BOOL fPerUserInstall, 
	int iBitness)
{
	HRESULT hr = S_OK;

    hr = RegInitialize();
    ExitOnFailure(hr, "Failed to initialize the registry functions.");

	LPWSTR wszAsmCodeBase = NULL;

	hr = StrAllocFormatted(&wszAsmCodeBase, L"file:///%ls", strAsmPath);
	ExitOnFailure(hr, "failed to allocate string");

	HKEY hKeyRoot = fPerUserInstall
		? HKEY_CURRENT_USER 
		: HKEY_LOCAL_MACHINE;

	if (iBitness == REG_KEY_32BIT || iBitness == REG_KEY_DEFAULT)
	{
		hr = RegisterCOMAtRoot(strClsId, strProgId, typeFullName, strAsmName, strAsmVersion, wszAsmCodeBase, strRuntimeVersion,
			hKeyRoot, REG_KEY_32BIT);
		ExitOnFailure1(hr, "failed to register COM");
	}
	if (iBitness == REG_KEY_64BIT || iBitness == REG_KEY_DEFAULT)
	{
		hr = RegisterCOMAtRoot(strClsId, strProgId, typeFullName, strAsmName, strAsmVersion, wszAsmCodeBase, strRuntimeVersion,
			hKeyRoot, REG_KEY_64BIT);
		ExitOnFailure1(hr, "failed to register COM");
	}

LExit:
	ReleaseStr(wszAsmCodeBase);

	RegUninitialize();
	return hr;
}

HRESULT __stdcall UnregisterCOM(
	LPCWSTR strClsId,
    LPCWSTR strProgId,
	LPCWSTR strAsmVersion,
	BOOL fPerUserInstall, 
	int iBitness)
{
	HRESULT hr = S_OK;

    hr = RegInitialize();
    ExitOnFailure(hr, "Failed to initialize the registry functions.");

	HKEY hKeyRoot = fPerUserInstall
		? HKEY_CURRENT_USER 
		: HKEY_LOCAL_MACHINE;

	if (iBitness == REG_KEY_32BIT || iBitness == REG_KEY_DEFAULT)
	{
		hr = UnregisterCOMAtRoot(strClsId, strProgId, strAsmVersion, hKeyRoot, REG_KEY_32BIT);
		ExitOnFailure1(hr, "failed to unregister COM");
	}
	if (iBitness == REG_KEY_64BIT || iBitness == REG_KEY_DEFAULT)
	{
		hr = UnregisterCOMAtRoot(strClsId, strProgId, strAsmVersion, hKeyRoot, REG_KEY_64BIT);
		ExitOnFailure1(hr, "failed to unregister COM");
	}

LExit:

	RegUninitialize();
	return hr;
}

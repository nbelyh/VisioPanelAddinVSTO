
#include "stdafx.h"

LPCWSTR vcsRegisterAddinsQuery =
    L"SELECT "
		L"`AddinRegistration`.`AddinRegistration`, "
		L"`AddinRegistration`.`File_`, "
		L"`AddinRegistration`.`FriendlyName`, "
		L"`AddinRegistration`.`Description`, "
		L"`AddinRegistration`.`Bitness`, "
		L"`AddinRegistration`.`CommandLineSafe`, "
		L"`AddinRegistration`.`LoadBehavior`, "
		L"`File`.`Component_` "
    L"FROM `AddinRegistration`, `File` WHERE `File`.`File`=`AddinRegistration`.`File_`";

enum eRegisterAddinsQuery { arqId = 1, arqFile, arqName, arqDescription, arqBitness, arqCommandLineSafe, arqLoadBehavior, arqComponent };


/***************************************************************/

HRESULT CreateRegistryKey(REG_KEY_BITNESS iBitness, 
	LPCWSTR pwzId, 
	LPCWSTR pwzFile, 
	LPCWSTR pwzFriendlyName, 
	LPCWSTR pwzDescription,
	int iCommandLineSafe,
	int iLoadBehavior,
	HKEY hKeyRoot)
{
	HRESULT hr = S_OK;
	HKEY hKey = NULL;

	LPWSTR wszRegPath = NULL;
	LPWSTR wszFileEntry = NULL;

    WcaLog(LOGMSG_VERBOSE, "CreateRegistryKey: Bitness=%ld, Id=%ls, Manifest=%ls, FriendlyName=%ls, Description=%ls", 
		iBitness, pwzId, pwzFile, pwzFriendlyName, pwzDescription);

    hr = StrAllocFormatted(&wszRegPath, L"Software\\Microsoft\\Visio\\Addins\\%ls", pwzId);
    ExitOnFailure(hr, "Failed to allocate string for registry path.");

	DWORD sam = KEY_READ|KEY_WRITE;
	if (iBitness == REG_KEY_32BIT)
		sam |= KEY_WOW64_32KEY;
	if (iBitness == REG_KEY_64BIT)
		sam |= KEY_WOW64_64KEY;

	hr = RegCreateEx(hKeyRoot, wszRegPath, sam, FALSE, NULL, &hKey, NULL);
	ExitOnFailure(hr, "Failed to create or open the registry key: %ls. It looks like Visio is not installed or security issue", wszRegPath);

	WcaLog(LOGMSG_VERBOSE, "Created or opened registry key: %ls", wszRegPath);

    hr = StrAllocFormatted(&wszFileEntry, L"file:///%ls|vstolocal", pwzFile);
    ExitOnFailure(hr, "Failed to allocate string for registry file entry.");

	hr = RegWriteString(hKey, L"Manifest", wszFileEntry);
	ExitOnFailure(hr, "Failed set Manifest value.");

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

/***************************************************************/

HRESULT DeleteRegistryKey(REG_KEY_BITNESS iBitness, LPCWSTR pwzId, HKEY hKeyRoot)
{
	HRESULT hr = S_OK;

    WcaLog(LOGMSG_VERBOSE, "DeleteRegistryKey: Bitness:%ld, Id=%ls", 
		iBitness, pwzId);

	LPWSTR wszRegPath = NULL;
    hr = StrAllocFormatted(&wszRegPath, L"Software\\Microsoft\\Visio\\Addins\\%ls", pwzId);
    ExitOnFailure(hr, "Failed to allocate registry path.");

	hr = RegDelete(hKeyRoot, wszRegPath, iBitness, FALSE);

	if (E_FILENOTFOUND != hr)
		ExitOnFailure(hr, "Failed to delete the registry key: %ls.", wszRegPath);

	WcaLog(LOGMSG_VERBOSE, "Deleted registry key: %ls", wszRegPath);

LExit:
	return hr;
}

const UINT COST_REGISTER_ADDIN = 100;

/******************************************************************
 SchedAddinRegistration - entry point for AddinRegistration Custom Action

********************************************************************/
HRESULT SchedAddinRegistration(MSIHANDLE hInstall, BOOL fInstall)
{
    // AssertSz(FALSE, "debug SchedRegisterAddins");

    HRESULT hr = S_OK;

    LPWSTR pwzCustomActionData = NULL;

    PMSIHANDLE hView = NULL;
    PMSIHANDLE hRec = NULL;

    LPWSTR pwzData = NULL;
    LPWSTR pwzTemp = NULL;
    LPWSTR pwzComponent = NULL;

    LPWSTR pwzId = NULL;
    LPWSTR pwzFile = NULL;
	LPWSTR pwzFriendlyName = NULL;
	LPWSTR pwzDescription = NULL;

	int iBitness = 0;
	int iCommandLineSafe = 1;
	int iLoadBehavior = 0;

	LPWSTR pwzAllUsers = NULL;

    int nAddins = 0;

	hr = WcaGetProperty(L"ALLUSERS", &pwzAllUsers);
	ExitOnFailure(hr, "failed to read value of ALLUSERS property");

    // loop through all the RegisterAddin records
    hr = WcaOpenExecuteView(vcsRegisterAddinsQuery, &hView);
    ExitOnFailure(hr, "failed to open view on AddinRegistration table");

    while (S_OK == (hr = WcaFetchRecord(hView, &hRec)))
    {
		++nAddins;

        // Get Id
        hr = WcaGetRecordString(hRec, arqId, &pwzId);
        ExitOnFailure(hr, "failed to get AddinRegistration.AddinRegistration");

        // Get File
        hr = WcaGetRecordString(hRec, arqFile, &pwzData);
        ExitOnFailure1(hr, "failed to get AddinRegistration.File_ for record: %ls", pwzId);
        hr = StrAllocFormatted(&pwzTemp, L"[#%ls]", pwzData);
        ExitOnFailure1(hr, "failed to format file string for file: %ls", pwzData);
        hr = WcaGetFormattedString(pwzTemp, &pwzFile);
        ExitOnFailure1(hr, "failed to get formatted string for file: %ls", pwzData);

        // Get name
        hr = WcaGetRecordFormattedString(hRec, arqName, &pwzFriendlyName);
        ExitOnFailure1(hr, "failed to get AddinRegistration.Name for record: %ls", pwzId);

        // Get description
        hr = WcaGetRecordFormattedString(hRec, arqDescription, &pwzDescription);
        ExitOnFailure1(hr, "failed to get AddinRegistration.Description for record: %ls", pwzId);

        // Get description
        hr = WcaGetRecordInteger(hRec, arqBitness, &iBitness);
        ExitOnFailure1(hr, "failed to get AddinRegistration.Bitnesss for record: %ls", pwzId);

        // Get description
        hr = WcaGetRecordInteger(hRec, arqCommandLineSafe, &iCommandLineSafe);
        ExitOnFailure1(hr, "failed to get AddinRegistration.CommandLineSafe for record: %ls", pwzId);

        // Get description
        hr = WcaGetRecordInteger(hRec, arqLoadBehavior, &iLoadBehavior);
        ExitOnFailure1(hr, "failed to get AddinRegistration.LoadBehavior for record: %ls", pwzId);

		// get component and its install/action states
        hr = WcaGetRecordString(hRec, arqComponent, &pwzComponent);
        ExitOnFailure(hr, "failed to get addin component id");

        // we need to know if the component's being installed, uninstalled, or reinstalled
        WCA_TODO todo = WcaGetComponentToDo(pwzComponent);

        // skip this entry if this is the install CA and we are uninstalling the component
        if (fInstall && WCA_TODO_UNINSTALL == todo)
        {
            continue;
        }

        // skip this entry if this is an uninstall CA and we are not uninstalling the component
        if (!fInstall && WCA_TODO_UNINSTALL != todo)
        {
            continue;
        }

        // write custom action data: operation, instance guid, path, directory
        hr = WcaWriteIntegerToCaData(todo, &pwzCustomActionData);
        ExitOnFailure1(hr, "failed to write operation to custom action data for instance id: %ls", pwzId);

        hr = WcaWriteStringToCaData(pwzId, &pwzCustomActionData);
        ExitOnFailure1(hr, "failed to write id to custom action data for instance id: %ls", pwzId);

        hr = WcaWriteStringToCaData(pwzFile, &pwzCustomActionData);
        ExitOnFailure1(hr, "failed to write custom action data for instance id: %ls", pwzId);

        hr = WcaWriteStringToCaData(pwzFriendlyName, &pwzCustomActionData);
        ExitOnFailure1(hr, "failed to write addin name to custom action data for instance id: %ls", pwzId);

        hr = WcaWriteStringToCaData(pwzDescription, &pwzCustomActionData);
        ExitOnFailure1(hr, "failed to write addin description to custom action data for instance id: %ls", pwzId);

        hr = WcaWriteIntegerToCaData(iBitness, &pwzCustomActionData);
        ExitOnFailure1(hr, "failed to write Bitness to custom action data for instance id: %ls", pwzId);

        hr = WcaWriteIntegerToCaData(iCommandLineSafe, &pwzCustomActionData);
        ExitOnFailure1(hr, "failed to write CommandLineSafe to custom action data for instance id: %ls", pwzId);

        hr = WcaWriteIntegerToCaData(iLoadBehavior, &pwzCustomActionData);
        ExitOnFailure1(hr, "failed to write LoadBehavior to custom action data for instance id: %ls", pwzId);

		hr = WcaWriteStringToCaData(pwzAllUsers, &pwzCustomActionData);
		ExitOnFailure(hr,  "failed to write allusers property to custom action data for instance id: %ls", pwzId);
	}

    if (E_NOMOREITEMS == hr)
        hr = S_OK;

    ExitOnFailure(hr, "failed while looping through all files to create native images for");

    // Schedule the install custom action
    if (pwzCustomActionData && *pwzCustomActionData)
    {
        WcaLog(LOGMSG_STANDARD, "Scheduling Addin Registration (%ls)", pwzCustomActionData);

        hr = WcaDoDeferredAction(L"RollbackAddinRegistration", pwzCustomActionData, nAddins * COST_REGISTER_ADDIN);
        ExitOnFailure(hr, "Failed to schedule addin registration rollback");

        hr = WcaDoDeferredAction(L"ExecAddinRegistration", pwzCustomActionData, nAddins * COST_REGISTER_ADDIN);
        ExitOnFailure(hr, "Failed to schedule addin registration execution");
    }

LExit:
	ReleaseStr(pwzAllUsers);
    ReleaseStr(pwzCustomActionData);
    ReleaseStr(pwzId);
    ReleaseStr(pwzData);
	ReleaseStr(pwzData);
    ReleaseStr(pwzTemp);
    ReleaseStr(pwzComponent);
    ReleaseStr(pwzFile);
	ReleaseStr(pwzFriendlyName);
	ReleaseStr(pwzDescription);

    return hr;
}

UINT __stdcall SchedAddinRegistrationInstall(MSIHANDLE hInstall)
{
	HRESULT hr = WcaInitialize(hInstall, "SchedRegisterAddinsInstall");
	ExitOnFailure(hr, "Failed to initialize");

	WcaLog(LOGMSG_STANDARD, "Started addin registration.");
	hr = SchedAddinRegistration(hInstall, TRUE);

LExit:
	return WcaFinalize(SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE);
}

UINT __stdcall SchedAddinRegistrationUninstall(MSIHANDLE hInstall)
{
	HRESULT hr = WcaInitialize(hInstall, "SchedRegisterAddinsUninstall");
	ExitOnFailure(hr, "Failed to initialize");

	WcaLog(LOGMSG_STANDARD, "Started addin registration removal.");
	hr = SchedAddinRegistration(hInstall, FALSE);

LExit:
	return WcaFinalize(SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE);
}

/******************************************************************
 ExecAddinRegistration - entry point for AddinRegistration Custom Action

*******************************************************************/

UINT __stdcall ExecAddinRegistration(MSIHANDLE hInstall)
{
    // AssertSz(FALSE, "debug ExecAddinRegistration");

    LPWSTR pwzCustomActionData = NULL;
    LPWSTR pwzData = NULL;
    LPWSTR pwz = NULL;

    int iOperation = 0;
    LPWSTR pwzId = NULL;
    LPWSTR pwzFile = NULL;
	LPWSTR pwzName = NULL;
	LPWSTR pwzDescription = NULL;
	int iBitness = REG_KEY_DEFAULT;
	int iCommandLineSafe = 1;
	int iLoadBehavior = 3;

	LPWSTR pwzAllUsers = NULL;

	HRESULT hr = WcaInitialize(hInstall, "ExecAddinRegistration");
	ExitOnFailure(hr, "Failed to initialize");

    hr = WcaGetProperty( L"CustomActionData", &pwzCustomActionData);
    ExitOnFailure(hr, "failed to get CustomActionData");

    WcaLog(LOGMSG_TRACEONLY, "CustomActionData: %ls", pwzCustomActionData);

    pwz = pwzCustomActionData;

    hr = RegInitialize();
    ExitOnFailure(hr, "Failed to initialize the registry functions.");

    // loop through all the passed in data
    while (pwz && *pwz)
    {
        // extract the custom action data
        hr = WcaReadIntegerFromCaData(&pwz, &iOperation);
        ExitOnFailure(hr, "failed to read operation from custom action data");

        hr = WcaReadStringFromCaData(&pwz, &pwzId);
        ExitOnFailure(hr, "failed to read id from custom action data");

        hr = WcaReadStringFromCaData(&pwz, &pwzFile);
        ExitOnFailure(hr, "failed to read path from custom action data");

        hr = WcaReadStringFromCaData(&pwz, &pwzName);
        ExitOnFailure(hr, "failed to read name from custom action data");

        hr = WcaReadStringFromCaData(&pwz, &pwzDescription);
        ExitOnFailure(hr, "failed to read description from custom action data");

		hr = WcaReadIntegerFromCaData(&pwz, &iBitness);
		ExitOnFailure(hr, "failed to read bitness from custom action data");

		hr = WcaReadIntegerFromCaData(&pwz, &iCommandLineSafe);
		ExitOnFailure(hr, "failed to read CommandLineSafe from custom action data");

		hr = WcaReadIntegerFromCaData(&pwz, &iLoadBehavior);
		ExitOnFailure(hr, "failed to read LoadBehavior from custom action data");

		hr = WcaReadStringFromCaData(&pwz, &pwzAllUsers);
		ExitOnFailure(hr, "failed to read ALLUSERS from custom action data");

		BOOL fPerUserInstall = (!pwzAllUsers || !*pwzAllUsers);

        // if rolling back, swap INSTALL and UNINSTALL
        if (::MsiGetMode(hInstall, MSIRUNMODE_ROLLBACK))
        {
            if (WCA_TODO_INSTALL == iOperation)
            {
                iOperation = WCA_TODO_UNINSTALL;
            }
            else if (WCA_TODO_UNINSTALL == iOperation)
            {
                iOperation = WCA_TODO_INSTALL;
            }
        }

        switch (iOperation)
        {
        case WCA_TODO_INSTALL:
        case WCA_TODO_REINSTALL:
			if (fPerUserInstall)
			{
				hr = CreateRegistryKey(REG_KEY_32BIT, pwzId, pwzFile, pwzName, pwzDescription, iCommandLineSafe, iLoadBehavior, HKEY_CURRENT_USER);
				ExitOnFailure1(hr, "failed to register addin (HKCU): %ls", pwzId);
			}
			else
			{
				if (iBitness == REG_KEY_32BIT || iBitness == REG_KEY_DEFAULT)
				{
					hr = CreateRegistryKey(REG_KEY_32BIT, pwzId, pwzFile, pwzName, pwzDescription, iCommandLineSafe, iLoadBehavior, HKEY_LOCAL_MACHINE);
					ExitOnFailure1(hr, "failed to register addin (HKLM, 32bit): %ls", pwzId);
				}
				if (iBitness == REG_KEY_64BIT || iBitness == REG_KEY_DEFAULT)
				{
					hr = CreateRegistryKey(REG_KEY_64BIT, pwzId, pwzFile, pwzName, pwzDescription, iCommandLineSafe, iLoadBehavior, HKEY_LOCAL_MACHINE);
					ExitOnFailure1(hr, "failed to register addin (HKLM, 64bit): %ls", pwzId);
				}
			}
            break;

        case WCA_TODO_UNINSTALL:
			if (fPerUserInstall)
			{
				hr = DeleteRegistryKey(REG_KEY_DEFAULT, pwzId, HKEY_CURRENT_USER);
				ExitOnFailure1(hr, "failed to unregister addin (HKCU): %ls", pwzId);
			}
			else
			{
				if (iBitness == REG_KEY_32BIT || iBitness == REG_KEY_DEFAULT)
				{
					hr = DeleteRegistryKey(REG_KEY_32BIT, pwzId, HKEY_LOCAL_MACHINE);
					ExitOnFailure1(hr, "failed to unregister addin (HKLM, 32bit): %ls", pwzId);
				}
				if (iBitness == REG_KEY_64BIT || iBitness == REG_KEY_DEFAULT)
				{
					hr = DeleteRegistryKey(REG_KEY_64BIT, pwzId, HKEY_LOCAL_MACHINE);
					ExitOnFailure1(hr, "failed to unregister addin (HKLM, 64bit): %ls", pwzId);
				}
			}
            break;
        }

        // Tick the progress bar along for this addin
        hr = WcaProgressMessage(COST_REGISTER_ADDIN, FALSE);
        ExitOnFailure1(hr, "failed to tick progress bar for addin registration: %ls", pwzId);
	}

LExit:
    RegUninitialize();

	ReleaseStr(pwzAllUsers);
    ReleaseStr(pwzCustomActionData);
    ReleaseStr(pwzData);

    ReleaseStr(pwzId);
    ReleaseStr(pwzFile);
	ReleaseStr(pwzName);
	ReleaseStr(pwzDescription);

	return WcaFinalize(SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE);
}

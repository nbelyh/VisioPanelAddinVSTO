
#pragma once

typedef enum ADDIN_TYPE
{
    ADDIN_TYPE_UNKNOWN = 0,
    ADDIN_TYPE_VSTO = 1,
    ADDIN_TYPE_COM = 2
} ADDIN_TYPE;

HRESULT CreateOfficeRegistryKey(
	LPCWSTR pwzId, 
	LPCWSTR pwzFile, 
	LPCWSTR pwzFriendlyName, 
	LPCWSTR pwzDescription,
	int iCommandLineSafe,
	int iLoadBehavior,
	int iAddinType,
	BOOL fPerUserInstall, 
	int iBitness);

HRESULT DeleteOfficeRegistryKey(
	LPCWSTR pwzId, 
	BOOL fPerUserInstall, 
	int iBitness);

HRESULT __stdcall RegisterCOM(
    LPCWSTR strClsId,
    LPCWSTR strProgId,
    LPCWSTR typeFullName, 
    LPCWSTR strAsmName, 
    LPCWSTR strAsmVersion, 
    LPCWSTR strAsmPath, 
    LPCWSTR strRuntimeVersion,
	BOOL fPerUserInstall, 
	int iBitness);

HRESULT __stdcall UnregisterCOM(
	LPCWSTR strClsId,
    LPCWSTR strProgId,
	LPCWSTR strAsmVersion,
	BOOL fPerUserInstall, 
	int iBitness);

//===========================================================================
//
// A3D_DLL.CPP
//
// Purpose: Aureal A3D 1.x interface library (Wrapper to DirectSound3D).
//
// Copyright (C) 2004 Dmitry Nesterenko. All rights reserved.
//
//===========================================================================


#include <objbase.h>
#include <initguid.h>
#include <dsound.h>


#include "ia3dapi.h"
#include "a3d_dll.h"
#include "a3d_dal.h"


#ifdef _DEBUG
#	include <crtdbg.h>
#	include <stdio.h>
#	include <stdarg.h>
#else
#	pragma intrinsic(memset, memcmp, memcpy, strcpy, strcat, strlen)
#endif


// Compilation options.
#ifdef _DEBUG
#	define WRITE_LOG_FILE TRUE
#endif


// Handle of shared library module.
HMODULE g_hModule = NULL;

// Handle of application window.
HWND g_hWnd = NULL;

// Counter number of locks.
LONG g_cLock = 0;

// Counter number of all objects.
LONG g_cObj = 0;


#ifdef _DEBUG
//===========================================================================
//
// ::LogMsg
//
// Purpose: Output logging message.
//
// Parameters:
//  pszFormat       LPCTSTR pointer to format string for message.
//  ...             (?) set other parameters for message.
//
//===========================================================================
void LogMsg(LPCTSTR pcszFormat, ...)
{
	_ASSERTE(pcszFormat && !IsBadReadPtr(pcszFormat, sizeof(*pcszFormat)));

	va_list vlArgs;
	TCHAR szBuffer[256];
#if WRITE_LOG_FILE
	TCHAR szLogName[MAX_PATH];
	TCHAR szCallName[MAX_PATH];

	// Get library module file name and truncate extension from it.
	DWORD dwLength = GetModuleFileName(g_hModule, szLogName, sizeof(szLogName));
	for (UINT i = dwLength; i; i--)
	{
		// Found file extension pointer.
		if (TEXT('.') == szLogName[i])
			break;

		// Found begin of file name (file extension not exist).
		if (TEXT('\\') == szLogName[i] || TEXT(':') == szLogName[i])
		{
			i = dwLength;
			break;
		}
	}

	// Get caller process module file name and extract it name.
	dwLength = GetModuleFileName(NULL, szCallName, sizeof(szCallName));
	for (UINT j = dwLength; j; j--)
	{
		// Truncate file extension.
		if (TEXT('.') == szCallName[j])
		{
			dwLength = j + 1;
			continue;
		}

		// Found begin of file name.
		if (TEXT('\\') == szCallName[j] || TEXT(':') == szCallName[j])
		{
			j++;
			break;
		}
	}

	// Make log file name from library and caller names.
	szLogName[i++] = TEXT('_');
	lstrcat(lstrcpyn(&szLogName[i], &szCallName[j], dwLength - j), TEXT(".log"));

	// Create or open log file.
	HANDLE hLogFile = CreateFile(szLogName, GENERIC_WRITE, FILE_SHARE_READ,
		NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hLogFile)
	{
		SYSTEMTIME timeLocal;
		DWORD dwWritten;

		// Save message time prefix to output string.
		GetLocalTime(&timeLocal);
		wsprintf(szBuffer, TEXT("%02u/%02u/%04u %02u:%02u:%02u.%03u "),
			timeLocal.wDay, timeLocal.wMonth, timeLocal.wYear,
			timeLocal.wHour, timeLocal.wMinute, timeLocal.wSecond,
			timeLocal.wMilliseconds);

		// Format logging message.
		va_start(vlArgs,pcszFormat);
		vsprintf(&szBuffer[lstrlen(szBuffer)], pcszFormat, vlArgs);
		va_end(vlArgs);

		// Write message string to log file.
		lstrcat(szBuffer, TEXT("\r\n"));
		SetFilePointer(hLogFile, 0, NULL, FILE_END);
		WriteFile(hLogFile, szBuffer, lstrlen(szBuffer) * sizeof(TCHAR), &dwWritten, NULL);
		CloseHandle(hLogFile);
	}
#endif

	// Save message prefix to output string.
	lstrcpy(szBuffer, TEXT("A3D: "));

	// Format logging message.
	va_start(vlArgs, pcszFormat);
	vsprintf(&szBuffer[lstrlen(szBuffer)], pcszFormat, vlArgs);
	va_end(vlArgs);

	// Output message to debugger.
	lstrcat(szBuffer, TEXT("\n"));
	OutputDebugString(szBuffer);
}


//===========================================================================
//
// ::Result
//
// Purpose: Convert standart return code to text string.
//
// Parameters:
//  hResult         HRESULT return code value.
//
// Return: Pointer to text string with result.
//
//===========================================================================
LPCTSTR Result(HRESULT hResult)
{
	if (SUCCEEDED(hResult))
		return TEXT("Ok");

	TCHAR szMessage[256];

	// Get system message for error code.
	if (!FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
	NULL, hResult, 0, szMessage, sizeof(szMessage) / sizeof(TCHAR), NULL))
		lstrcpy(szMessage, TEXT("UNKNOWN ERROR!"));
	else
		for (UINT i = 0; TEXT('\0') != szMessage[i]; i++)
			if (TEXT('\r') == szMessage[i] || TEXT('\n') == szMessage[i])
			{
				szMessage[i] = TEXT('\0');
				break;
			}

	static TCHAR szReturn[256];

	// Format error code and message text.
	wsprintf(szReturn, TEXT("%#x:\"%s\""), hResult, szMessage);

	return szReturn;
}
#endif


//===========================================================================
//
// ::GuidToStr
//
// Purpose: Convert GUID to text string.
//
// Parameters:
//  rGuid           REFGUID reference for GUID.
//  pszBuffer       LPTSTR buffer pointer for text string.
//  iLen            UINT buffer size for text string.
//
// Return: Pointer to text string with GUID.
//
//===========================================================================
LPTSTR GuidToStr(REFGUID rGuid, LPTSTR pszBuffer, UINT iSize)
{
#ifdef _DEBUG
	_ASSERTE(pszBuffer && !IsBadWritePtr(pszBuffer, sizeof(*pszBuffer)));
	_ASSERTE(iSize);
#endif
	if (!&rGuid)
		lstrcpyn(pszBuffer, TEXT("{NULL}"), iSize);
	else
	{
		WCHAR wcGuid[64];

		// Convert GUID to text string.
		StringFromGUID2(rGuid, wcGuid, sizeof(wcGuid) / sizeof(WCHAR));
#ifdef UNICODE
		lstrcpyn(pszBuffer, wcGuid, iSize);
#else
		WideCharToMultiByte(CP_ACP, 0, wcGuid, -1, pszBuffer, iSize, NULL, NULL);
#endif
	}

	return pszBuffer;
}


//===========================================================================
//
// ::SplashScreen
//
// Purpose: Launch A3D splash application.
//
// Parameters:
//  hWnd            HWND handler of parent window.
//
//===========================================================================
static void SplashScreen(HWND hWnd)
{
#ifdef _DEBUG
	LogMsg(TEXT("SplashScreen(%#x)"), hWnd);
#endif
	// Execute this function only first time.
	static bool bFirstCall = true;
	if (bFirstCall)
		bFirstCall = false;
	else
		return;

	TCHAR szSplashFile[MAX_PATH];
	HKEY hKey;
	DWORD dwPlayAudio = 0, dwDisplayScreen = 0;
	bool bPathExist = false;

	// Open registry entry under HKEY_LOCAL_MACHINE.
	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Aureal\\A3D"), 0,
	KEY_READ, &hKey) == ERROR_SUCCESS)
	{
		DWORD dwSize = sizeof(szSplashFile);

		// Read registry value.
		if (RegQueryValueEx(hKey, TEXT("SplashPath"), NULL, NULL,
		(LPBYTE)szSplashFile, &dwSize) == ERROR_SUCCESS)
			bPathExist = true;

		// Read registry values and close registry entry.
		dwSize = sizeof(dwPlayAudio);
		RegQueryValueEx(hKey, TEXT("SplashAudio"), NULL, NULL,
			(LPBYTE)dwPlayAudio, &dwSize);
		dwSize = sizeof(dwDisplayScreen);
		RegQueryValueEx(hKey, TEXT("SplashScreen"), NULL, NULL,
			(LPBYTE)dwDisplayScreen, &dwSize);
		RegCloseKey(hKey);
	}

	// Make file name for launch splash application.
	if (bPathExist)
		lstrcat(szSplashFile, TEXT("\\A3DSplsh.exe"));
	else
		lstrcpy(szSplashFile, TEXT("A3DSplsh.exe"));

	// Launch A3D splash application.
	ShellExecute(hWnd, NULL, szSplashFile, NULL, NULL, SW_SHOWNORMAL);

	// Delay for play audio and/or display logo screen.
	if (dwPlayAudio || dwDisplayScreen)
		Sleep(3500);

	// Restore parent window.
	if (hWnd && IsIconic(hWnd))
		ShowWindow(hWnd, SW_RESTORE);
}


//===========================================================================
//
// ::CreateComObject
//
// Purpose: Create the registry entries for COM object.
//
// Parameters:
//  rObjGuid        REFGUID object CLSID.
//  pcszRootName    LPCTSTR name of object in HKEY_CLASSES_ROOT.
//  pcszObjName     LPCTSTR description of object.
//
// Return: S_OK if registration successful, error otherwise.
//
//===========================================================================
static void CreateComObject(REFGUID rObjGuid, LPCTSTR pcszRootName,
	LPCTSTR pcszObjName)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("CreateComObject(%s,%s,%s)"),
		GuidToStr(rObjGuid, szGuid, sizeof(szGuid)), pcszRootName, pcszObjName);
	_ASSERTE(pcszRootName && !IsBadReadPtr(pcszRootName, sizeof(*pcszRootName)));
	_ASSERTE(pcszObjName && !IsBadReadPtr(pcszObjName, sizeof(*pcszObjName)));
#endif
	TCHAR szObjGuid[64];
	TCHAR cszRegSz[] = TEXT("REG_SZ");
	TCHAR szKeyName[128];
	HKEY hKey, hSubKey;

	// Create the registry entry name string using CLSID.
	GuidToStr(rObjGuid, szObjGuid, sizeof(szObjGuid));

	// Create registry entry HKEY_CLASSES_ROOT\<Root Name>.
	if (RegCreateKeyEx(HKEY_CLASSES_ROOT, pcszRootName, 0, cszRegSz,
	REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
	{
		// Set default value for registry key.
		RegSetValueEx(hKey, NULL, 0, REG_SZ, (LPBYTE)pcszObjName,
			lstrlen(pcszObjName) + 1);

		// Create registry entry HKEY_CLASSES_ROOT\<Root Name>\CLSID.
		if (RegCreateKeyEx(hKey, TEXT("CLSID"), 0, cszRegSz,
		REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hSubKey, NULL) == ERROR_SUCCESS)
		{
			// Set default value for registry key.
			RegSetValueEx(hSubKey, NULL, 0, REG_SZ, (LPBYTE)szObjGuid,
				lstrlen(szObjGuid) + 1);
			RegCloseKey(hSubKey);
		}

		RegCloseKey(hKey);
	}

	// Create the registry entry name for object using it CLSID.
	lstrcat(lstrcpy(szKeyName, TEXT("CLSID\\")), szObjGuid);

	// Create registry entry HKEY_CLASSES_ROOT\CLSID\{clsid}.
	if (RegCreateKeyEx(HKEY_CLASSES_ROOT, szKeyName, 0, cszRegSz,
	REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
	{
		// Set default value for registry key.
		RegSetValueEx(hKey, NULL, 0, REG_SZ, (LPBYTE)pcszObjName,
			lstrlen(pcszObjName) + 1);

		// Set registry entry application identifier.
		RegSetValueEx(hKey, TEXT("AppID"), 0, REG_SZ, (LPBYTE)szObjGuid,
			lstrlen(szObjGuid) + 1);

		// Create registry entry HKEY_CLASSES_ROOT\CLSID\{clsid}\InprocServer32.
		if (RegCreateKeyEx(hKey, TEXT("InprocServer32"), 0, cszRegSz,
		REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hSubKey, NULL) == ERROR_SUCCESS)
		{
			TCHAR szDllName[MAX_PATH];
			const TCHAR cszApartment[] = TEXT("Apartment");

			// Get shared library module name.
			GetModuleFileName(g_hModule, szDllName,  sizeof(szDllName));

			// Set default value for registry key.
			RegSetValueEx(hSubKey, NULL, 0, REG_SZ, (LPBYTE)szDllName,
				lstrlen(szDllName) + 1);

			// Set threading model value for registry entry.
			RegSetValueEx(hSubKey, TEXT("ThreadingModel"), 0, REG_SZ,
				(LPBYTE)cszApartment, sizeof(cszApartment));
			RegCloseKey(hSubKey);
		}

		// Close registry entry.
		RegCloseKey(hKey);
	}
}


//===========================================================================
//
// ::DeleteComObject
//
// Purpose: Remove the registry entries for COM object.
//
// Parameters:
//  rObjGuid        REFGUID object CLSID.
//  pcszRootName    LPCTSTR name of object in HKEY_CLASSES_ROOT.
//
//===========================================================================
static void DeleteComObject(REFGUID rObjGuid, LPCTSTR pcszRootName)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("DeleteComObject(%s,%s)"),
		GuidToStr(rObjGuid, szGuid, sizeof(szGuid)), pcszRootName);
	_ASSERTE(pcszRootName && !IsBadReadPtr(pcszRootName, sizeof(*pcszRootName)));
#endif
	TCHAR szObjGuid[64];
	TCHAR szKeyName[128];
	HKEY hKey;

	// Open registry entry HKEY_CLASSES_ROOT\<Root Name>.
	if (RegOpenKeyEx(HKEY_CLASSES_ROOT, pcszRootName, 0, KEY_WRITE, &hKey) == ERROR_SUCCESS)
	{
		// Delete registry entry HKEY_CLASSES_ROOT\<Root Name>\CLSID.
		RegDeleteKey(hKey, TEXT("CLSID"));
		RegCloseKey(hKey);

		// Delete registry entry HKEY_CLASSES_ROOT\<Root Name>.
		RegDeleteKey(HKEY_CLASSES_ROOT, pcszRootName);
	}

	// Create the registry entry name for object using it CLSID.
	GuidToStr(rObjGuid, szObjGuid, sizeof(szObjGuid));
	lstrcat(lstrcpy(szKeyName, TEXT("CLSID\\")), szObjGuid);

	// Open registry entry HKEY_CLASSES_ROOT\CLSID\{clsid}.
	if (RegOpenKeyEx(HKEY_CLASSES_ROOT, szKeyName, 0, KEY_WRITE, &hKey) == ERROR_SUCCESS)
	{
		// Delete registry entry HKEY_CLASSES_ROOT\CLSID\{clsid}\InprocServer32.
		RegDeleteKey(hKey, TEXT("InprocServer32"));
		RegCloseKey(hKey);

		// Delete registry entry HKEY_CLASSES_ROOT\CLSID\{clsid}.
		RegDeleteKey(HKEY_CLASSES_ROOT, szKeyName);
	}
}


//===========================================================================
//
// ::DllMain
//
// Purpose: Entry point for DLL.
//
// Parameters:
//  hInstance       HANDLE instance of DLL module.
//  dwReason        DWORD reason of call.
//  pvReserved      LPVOID reserved.
//
// Return: TRUE if OK.
//
//===========================================================================
BOOL WINAPI DllMain(HANDLE hInstance, DWORD dwReason, LPVOID pvReserved)
{
	// Save library module handle and disable thread notification calls.
	if (DLL_PROCESS_ATTACH == dwReason)
	{
#ifdef _DEBUG
		// Load DirectSound shared library DSOUND.DLL.
		HINSTANCE hDsDll = LoadLibrary(TEXT("dsound.dll"));
		if (!hDsDll)
			return FALSE;

		// Get address of the GetDeviceID function.
		FARPROC pfnGetDeviceID = GetProcAddress(hDsDll, "GetDeviceID");

		// Free DirectSound shared library.
		FreeLibrary(hDsDll);

		// Check existing of entry point.
		if (!pfnGetDeviceID)
			return FALSE;
#endif
		// Save handle of shared library module.
		g_hModule = (HMODULE)hInstance;
		DisableThreadLibraryCalls(g_hModule);
	}

#ifdef _DEBUG
	LogMsg(TEXT("DllMain(%#x,%u)"), hInstance, dwReason);
#endif
	return TRUE;
}


//===========================================================================
//
// ::DllCanUnloadNow
//
// Purpose: Called periodically by OLE in order to determine if the
//          DLL can be freed.
//
// Return: S_OK if there are no objects in use and the class factory 
//         isn't locked.
//
//===========================================================================
STDAPI DllCanUnloadNow()
{
#ifdef _DEBUG
	LogMsg(TEXT("DllCanUnloadNow()=(%u,%u)"), g_cLock, g_cObj);
#endif
	// Check locks and objects counters.
	if (g_cLock || g_cObj)
		return S_FALSE;

	return S_OK;
}


//===========================================================================
//
// ::DllGetClassObject
//
// Purpose: Called by OLE when some client wants a class factory.  Return 
//          one only if it is the sort of class this DLL supports.
//
//===========================================================================
STDAPI DllGetClassObject(REFCLSID rClsid, REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("DllGetClassObject(%s)..."),
		GuidToStr(rClsid, szGuid, sizeof(szGuid)));
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

#ifdef _DEBUG
	if (CLSID_A3d != rClsid && CLSID_A3dDal != rClsid)
		LogMsg(TEXT("...DllGetClassObject()=E_FAIL!"));
#endif
	// Check arguments values.
	if (CLSID_A3d != rClsid && CLSID_A3dDal != rClsid)
		return E_FAIL;

	// Create new A3dClassFactory object.
	LPA3DCLASSFACTORY pA3dClassFactory = new IA3dClassFactory;
	if (!pA3dClassFactory)
		return E_OUTOFMEMORY;

	// Kill the object if initial creation failed.
	HRESULT hr = pA3dClassFactory->QueryInterface(rIid, ppvObj);
	if (FAILED(hr))
		delete pA3dClassFactory;

#ifdef _DEBUG
	LogMsg(TEXT("...=%s"), Result(hr));
#endif
	return hr;
}


//===========================================================================
//
// ::DllRegisterServer
//
// Purpose: Called during setup or by regsvr32.
//
// Return: S_OK if unregistration successful, error otherwise.
//
//===========================================================================
STDAPI DllRegisterServer()
{
#ifdef _DEBUG
	LogMsg(TEXT("DllRegisterServer()"));
#endif
	CreateComObject(CLSID_A3d, TEXT("A3d"), TEXT("A3d Object"));
	CreateComObject(CLSID_A3dDal, TEXT("A3dDAL"), TEXT("A3dDAL Object"));

	return S_OK;
}


//===========================================================================
//
// ::DllUnregisterServer
//
// Purpose: Called when it is time to remove the registry entries.
//
// Return: S_OK if unregistration successful, error otherwise.
//
//===========================================================================
STDAPI DllUnregisterServer()
{
#ifdef _DEBUG
	LogMsg(TEXT("DllUnregisterServer()"));
#endif
	DeleteComObject(CLSID_A3d, TEXT("A3d"));
	DeleteComObject(CLSID_A3dDal, TEXT("A3dDAL"));

	return S_OK;
}


//===========================================================================
//
// ::A3dCreate
//
// Purpose: Create new A3dSound object.
//
// Parameters:
//  pGuidDevice     LPGUID identifying the device the caller
//                  desires to have for the new object.
//  ppDS            LPLPDIRECTSOUND in which to store the desired
//                  interface pointer for the new object.
//  pUnkOuter       LPUNKNOWN to the controlling IUnknown if we are
//                  being used in an aggregation.
//
// Return: A3D_OK if successful, error otherwise.
//
//===========================================================================
HRESULT WINAPI A3dCreate(LPGUID pGuidDevice, LPDIRECTSOUND * ppDS, LPUNKNOWN pUnkOuter)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("A3dCreate(%s,%#x,%#x)"), GuidToStr((REFGUID)*pGuidDevice,
		szGuid, sizeof(szGuid)), ppDS, pUnkOuter);
	_ASSERTE(ppDS && !IsBadWritePtr(ppDS, sizeof(*ppDS)));
	_ASSERTE(!pUnkOuter);
#endif
	HKEY hKey;

	// Open registry key for A3D.
	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Aureal\\A3D"), 0,
	KEY_QUERY_VALUE, &hKey) == ERROR_SUCCESS)
	{
		INT iForceDS = 0;
		DWORD dwLength = sizeof(iForceDS);

		// Read value from registry key.
		RegQueryValueEx(hKey, TEXT("ForceDS"), NULL, NULL, (LPBYTE)&iForceDS, &dwLength);
		RegCloseKey(hKey);

		// Force using DirectSound.
		if (iForceDS)
		{
			// Standard beep using the computer speaker.
			MessageBeep(-1);

			// Get initialized DirectSound object.
			HRESULT hr = DirectSoundCreate(pGuidDevice, ppDS, NULL);
			if (SUCCEEDED(hr) && (iForceDS > 1))
				return A3D_OK;

			return hr;
		}
	}

	// Check arguments values.
	if (!ppDS)
		return E_POINTER;

	// For future invalid return.
	*ppDS = NULL;

	// This object doesnt support aggregation.
	if (pUnkOuter)
		return CLASS_E_NOAGGREGATION;

	LPDIRECTSOUND pDS;

	// Get initialized DirectSound object.
	HRESULT hr = DirectSoundCreate(pGuidDevice, &pDS, NULL);
	if (FAILED(hr))
		return hr;

	// Create new A3dSound object.
	LPA3DSOUND pA3dSound = new IA3dSound(pDS);
	// Decrease reference of DirectSound object.
	pDS->Release();
	if (!pA3dSound)
		return E_OUTOFMEMORY;

	// Kill the object if initial creation failed.
	hr = pA3dSound->QueryInterface(IID_IDirectSound, (LPVOID *)ppDS);
	if (FAILED(hr))
	{
		delete pA3dSound;
		return hr;
	}

	return A3D_OK;
}


//===========================================================================
//
// IA3dClassFactory::IA3dClassFactory
// IA3dClassFactory::~IA3dClassFactory
//
// Constructor Parameters:
//  None
//
//===========================================================================
IA3dClassFactory::IA3dClassFactory() :
	m_cRef(0)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dClassFactory::IA3dClassFactory()=%u"), g_cObj + 1);
#endif
	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dClassFactory::~IA3dClassFactory()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dClassFactory::~IA3dClassFactory()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dClassFactory::QueryInterface
// IA3dClassFactory::AddRef
// IA3dClassFactory::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dClassFactory::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dClassFactory::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

#ifdef _DEBUG
	if ((IID_IUnknown != rIid) && (IID_IClassFactory != rIid))
		LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// Only Unknown or ClassFactory objects.
	if ((IID_IUnknown != rIid) && (IID_IClassFactory != rIid))
		return E_NOINTERFACE;

	// Return this object;
	*ppvObj = this;
	AddRef();

	return S_OK;
}

STDMETHODIMP_(ULONG) IA3dClassFactory::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dClassFactory::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dClassFactory::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dClassFactory::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dClassFactory::CreateInstance
//
// Purpose: Instantiates a Locator object returning an interface pointer.
//
// Parameters:
//  pUnkOuter       LPUNKNOWN to the controlling IUnknown if we are
//                  being used in an aggregation.
//  rIid            REFIID identifying the interface the caller
//                  desires to have for the new object.
//  ppvObj          PPVOID in which to store the desired
//                  interface pointer for the new object.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dClassFactory::CreateInstance(LPUNKNOWN pUnkOuter,
	REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dClassFactory::CreateInstance(%s)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)));
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
	_ASSERTE(!pUnkOuter);
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

	// This object doesnt support aggregation.
	if (pUnkOuter)
		return CLASS_E_NOAGGREGATION;

	LPUNKNOWN pObj;

#ifdef _DEBUG
	if ((IID_IUnknown != rIid) && (IID_IClassFactory != rIid) &&
	(IID_IA3d != rIid) && (IID_IA3d2 != rIid) &&
	(IID_IDirectSound != rIid) && (IID_IA3dDal != rIid))
		LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// Create the locator object.
	if ((IID_IUnknown == rIid) || (IID_IClassFactory == rIid))
		pObj = new IA3dClassFactory;
	else if ((IID_IA3d == rIid) || (IID_IA3d2 == rIid))
		pObj = new IA3dX(NULL);
	else if (IID_IDirectSound == rIid)
		pObj = new IA3dSound(NULL);
	else if (IID_IA3dDal == rIid)
		pObj = new IA3dDal(NULL);
	else
		return E_NOINTERFACE;
	if (!pObj)
		return E_OUTOFMEMORY;

	// Kill the object if initial creation failed.
	HRESULT hr = pObj->QueryInterface(rIid, ppvObj);
	if (FAILED(hr))
		delete pObj;

	return hr;
}


//===========================================================================
//
// IA3dClassFactory::LockServer
//
// Purpose:
//  Increments or decrements the lock count of the DLL.  If the
//  lock count goes to zero and there are no objects, the DLL
//  is allowed to unload.  See DllCanUnloadNow.
//
// Parameters:
//  fLock           BOOL specifying whether to increment or
//                  decrement the lock count.
//
// Return: S_OK always.
//
//===========================================================================
STDMETHODIMP IA3dClassFactory::LockServer(BOOL fLock)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dClassFactory::LockServer(%u,%u)"), g_cLock, fLock);
#endif
	// Increase or decrease the lock counter.
	if (fLock)
		InterlockedIncrement(&g_cLock);
	else
		InterlockedDecrement(&g_cLock);

	return S_OK;
}


//===========================================================================
//
// IA3dX::CreateKsPropertySet
//
// Purpose: Get KsPropertySet object.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::CreateKsPropertySet()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::CreateKsPropertySet()..."));
	_ASSERTE(m_pDS);
#endif
	HRESULT hr;
	WAVEFORMATEX wfxFormat;
	ZeroMemory(&wfxFormat, sizeof(wfxFormat));
	wfxFormat.wFormatTag = WAVE_FORMAT_PCM;
	DSBUFFERDESC DSBufDesc;
	ZeroMemory(&DSBufDesc, sizeof(DSBufDesc));
	DSBufDesc.dwSize = sizeof(DSBufDesc);
	DSBufDesc.dwFlags = DSBCAPS_LOCHARDWARE | DSBCAPS_CTRL3D;
	DSBufDesc.dwBufferBytes = DSBSIZE_MIN;
	DSBufDesc.lpwfxFormat = &wfxFormat;

	// Loop for channels: 1 or 2.
	for (WORD wChannels = 1; wChannels <= 2; wChannels++)
	{
		wfxFormat.nChannels = wChannels;

		// Loop for samples per second: see list below.
		const DWORD dwSps[] = {8000, 11025, 22050, 44100, 48000};
		for (UINT i = 0; i < (sizeof(dwSps) / sizeof(dwSps[0])); i++)
		{
			wfxFormat.nSamplesPerSec = dwSps[i];

			// Loop for bits per sample: 8 or 16.
			for (WORD wBps = 8; wBps <= 16; wBps += 8)
			{
				wfxFormat.wBitsPerSample = wBps;
				wfxFormat.nBlockAlign = wfxFormat.wBitsPerSample / 8 * wfxFormat.nChannels;
				wfxFormat.nAvgBytesPerSec = wfxFormat.nSamplesPerSec * wfxFormat.nBlockAlign;

				LPDIRECTSOUNDBUFFER pDSB;

				// Create new DirectSoundBuffer object.
				hr = m_pDS->CreateSoundBuffer(&DSBufDesc, &pDSB, NULL);
				if (SUCCEEDED(hr))
				{
					// Get KsPropertySet object. 
					hr = pDSB->QueryInterface(IID_IKsPropertySet, (LPVOID *)&m_pKsPS);
					// Delete DirectSoundBuffer object. 
					pDSB->Release();
#ifdef _DEBUG
					if (SUCCEEDED(hr))
					{
						LogMsg(TEXT("...Channels=%u"), wfxFormat.nChannels);
						LogMsg(TEXT("...SamplesPerSec=%u"), wfxFormat.nSamplesPerSec);
						LogMsg(TEXT("...AvgBytesPerSec=%u"), wfxFormat.nAvgBytesPerSec);
						LogMsg(TEXT("...BlockAlign=%u"), wfxFormat.nBlockAlign);
						LogMsg(TEXT("...BitsPerSample=%u"), wfxFormat.wBitsPerSample);
					}

					LogMsg(TEXT("...=%s"), Result(hr));
#endif
					return hr;
				}
			}
		}
	}

#ifdef _DEBUG
	LogMsg(TEXT("...=%s"), Result(hr));
#endif
	return hr;
}


//===========================================================================
//
// IA3dX::SetOemResourceManagerMode
//
// Purpose: Set resource manager mode for OEM.
//
// Parameters:
//  dwResManMode    DWORD resource manager mode.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::SetOemResourceManagerMode(DWORD dwResManMode)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::SetOemResourceManagerMode(%u)..."), dwResManMode);
	_ASSERTE(m_pDS);
#endif
	// Create KsPropertySet object.
	HRESULT hr = (m_pKsPS) ? S_OK : CreateKsPropertySet();
	if (SUCCEEDED(hr))
	{
		ULONG ulTypeSupport;
		DWORD dwPropertyMode;

		// Check arguments values.
		switch (dwResManMode)
		{
		case A3D_RESOURCE_MODE_OFF:
			dwPropertyMode = 0;
			break;
		case A3D_RESOURCE_MODE_NOTIFY:
			dwPropertyMode = 2;
			break;
		case A3D_RESOURCE_MODE_DYNAMIC:
		case A3D_RESOURCE_MODE_DYNAMIC_LOOPERS:
			dwPropertyMode = 1;
			break;
		default:
			return E_INVALIDARG;
		}

		// Check DSPROPSETID_RMM_Creative property set support.
		hr = m_pKsPS->QuerySupport(DSPROPSETID_RMM_Creative, 0, &ulTypeSupport);
		if (SUCCEEDED(hr) && (ulTypeSupport & KSPROPERTY_SUPPORT_SET))
#ifdef _DEBUG
		{
			hr = m_pKsPS->Set(DSPROPSETID_RMM_Creative, 0, NULL, 0,
				&dwPropertyMode, sizeof(dwPropertyMode));
			LogMsg(TEXT("...Set(DSPROPSETID_RMM_Creative,%u)=%s"), dwPropertyMode, Result(hr));
		}
		else
			LogMsg(TEXT("...Set(DSPROPSETID_RMM_Creative) not supported!"));
#else
			m_pKsPS->Set(DSPROPSETID_RMM_Creative, 0, NULL, 0,
				&dwPropertyMode, sizeof(dwPropertyMode));
#endif

		// Check DSPROPSETID_RMM_Sensaura property set support.
		hr = m_pKsPS->QuerySupport(DSPROPSETID_RMM_Sensaura, 0, &ulTypeSupport);
		if (SUCCEEDED(hr) && (ulTypeSupport & KSPROPERTY_SUPPORT_SET))
#ifdef _DEBUG
		{
			hr = m_pKsPS->Set(DSPROPSETID_RMM_Sensaura, 0, NULL, 0,
				&dwPropertyMode, sizeof(dwPropertyMode));
			LogMsg(TEXT("...Set(DSPROPSETID_RMM_Sensaura,%u)=%s"), dwPropertyMode, Result(hr));
		}
		else
			LogMsg(TEXT("...Set(DSPROPSETID_RMM_Sensaura) not supported!"));
#else
			m_pKsPS->Set(DSPROPSETID_RMM_Sensaura, 0, NULL, 0,
				&dwPropertyMode, sizeof(dwPropertyMode));
#endif

		// Kill KsPropertySet object if resource manager mode off.
		if (A3D_RESOURCE_MODE_OFF == dwResManMode)
		{
			m_pKsPS->Release();
			m_pKsPS = NULL;
		}
	}

#ifdef _DEBUG
	LogMsg(TEXT("...=%s"), Result(hr));
#endif
	return hr;
}

//===========================================================================
//
// IA3dX::GetOemResourceManagerMode
//
// Purpose: Get resource manager mode for OEM.
//
// Parameters:
//  pdwResManMode   LPDWORD pointer to resource manager mode buffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::GetOemResourceManagerMode(LPDWORD pdwResManMode)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::GetOemResourceManagerMode()..."));
	_ASSERTE(pdwResManMode && !IsBadWritePtr(pdwResManMode, sizeof(*pdwResManMode)));
	_ASSERTE(m_pDS);
#endif
	// Create KsPropertySet object.
	HRESULT hr = (m_pKsPS) ? S_OK : CreateKsPropertySet();
	if (SUCCEEDED(hr))
	{
		ULONG ulTypeSupport;
		DWORD dwPropertyMode;

		// Check DSPROPSETID_RMM_Creative property set support.
		hr = m_pKsPS->QuerySupport(DSPROPSETID_RMM_Creative, 0, &ulTypeSupport);
		if (SUCCEEDED(hr) && (ulTypeSupport & KSPROPERTY_SUPPORT_GET))
		{
			// Get DSPROPSETID_RMM_Creative property set value.
			if (SUCCEEDED(m_pKsPS->Get(DSPROPSETID_RMM_Creative, 0, NULL, 0,
			&dwPropertyMode, sizeof(dwPropertyMode), NULL)))
			{
#ifdef _DEBUG
				LogMsg(TEXT("...Get(DSPROPSETID_RMM_Creative)=%u"), dwPropertyMode);
#endif
				// Convert returned value.
				switch (dwPropertyMode)
				{
				case 0:
					*pdwResManMode = A3D_RESOURCE_MODE_OFF;
					break;
				case 1:
					*pdwResManMode = A3D_RESOURCE_MODE_DYNAMIC;
					break;
				case 2:
					*pdwResManMode = A3D_RESOURCE_MODE_NOTIFY;
					break;
				default:
					*pdwResManMode = A3D_RESOURCE_MODE_OFF;
				}
			}
		}
#ifdef _DEBUG
		else
			LogMsg(TEXT("...Get(DSPROPSETID_RMM_Creative) not supported!"));
#endif

		// Kill KsPropertySet object if resource manager mode off.
		if (A3D_RESOURCE_MODE_OFF == *pdwResManMode)
		{
			m_pKsPS->Release();
			m_pKsPS = NULL;
		}
	}

#ifdef _DEBUG
	LogMsg(TEXT("...ResourceManagerMode=%u"), *pdwResManMode);
	LogMsg(TEXT("...=Ok"));
#endif
	return S_OK;
}


//===========================================================================
//
// IA3dX::SetOemHFAbsorbFactor
//
// Purpose: Set high frequencies absorb factor for OEM.
//
// Parameters:
//  fHFAbsorbFactor  FLOAT absorb factor.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::SetOemHFAbsorbFactor(FLOAT fHFAbsorbFactor)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::SetOemHFAbsorbFactor(%g)"), fHFAbsorbFactor);
#endif

	// Create KsPropertySet object.
	HRESULT hr = (m_pKsPS) ? S_OK : CreateKsPropertySet();
	if (SUCCEEDED(hr))
	{
		ULONG ulTypeSupport;

		// Check DSPROPSETID_RMM_Sensaura property set support.
		hr = m_pKsPS->QuerySupport(DSPROPSETID_HFAF_Sensaura, 1, &ulTypeSupport);
		if (SUCCEEDED(hr) && (ulTypeSupport & KSPROPERTY_SUPPORT_SET))
#ifdef _DEBUG
		{
			hr = m_pKsPS->Set(DSPROPSETID_HFAF_Sensaura, 1, NULL, 0,
				&fHFAbsorbFactor, sizeof(fHFAbsorbFactor));
			LogMsg(TEXT("...Set(DSPROPSETID_HFAF_Sensaura)=%s"), Result(hr));
		}
		else
			LogMsg(TEXT("...Set(DSPROPSETID_HFAF_Sensaura) not supported!"));
#else
			m_pKsPS->Set(DSPROPSETID_HFAF_Sensaura, 1, NULL, 0,
				&fHFAbsorbFactor, sizeof(fHFAbsorbFactor));
#endif
	}

	return hr;
}


//===========================================================================
//
// IA3dX::IA3dX
// IA3dX::~IA3dX
//
// Constructor Parameters:
//  pDS             LPDIRECTSOUND to the parent object.
//
//===========================================================================
IA3dX::IA3dX(LPDIRECTSOUND pDS) :
	m_cRef(0),
	m_dwAppVersion(IA3DVERSION_RELEASE12),
	m_dwResourceManagerMode(A3D_RESOURCE_MODE_DYNAMIC),
	m_dwFrontXtalkMode(OUTPUT_SPEAKERS_WIDE),
	m_dwBackXtalkMode(OUTPUT_SPEAKERS_WIDE),
	m_dwQuadMode(OUTPUT_MODE_STEREO),
	m_fHFAbsorbFactor(1.0f),
	m_pDS(pDS),
	m_pKsPS(NULL)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::IA3dX()=%u"), g_cObj + 1);
#endif
	// Add reference for parent DirectSound object.
	if (m_pDS)
		m_pDS->AddRef();
	else
		CoCreateInstance(CLSID_DirectSound, NULL, CLSCTX_INPROC_SERVER,
			IID_IDirectSound, (LPVOID *)&m_pDS);

	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dX::~IA3dX()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::~IA3dX()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Release parent DirectSound object.
	if (m_pDS)
	{
		// Release KsPropertySet object.
		if (m_pKsPS)
			m_pKsPS->Release();

		m_pDS->Release();
	}

	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dX::QueryInterface
// IA3dX::AddRef
// IA3dX::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dX::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dX::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

	// If not exist parent DirectSound object.
	if (!m_pDS)
		return E_FAIL;

	// Request for A3dDal object.
	if (IID_IA3dDal == rIid)
	{
		// Create new A3dDal object.
		LPA3DDAL pA3dDal = new IA3dDal(m_pDS);
		if (!pA3dDal)
			return E_OUTOFMEMORY;

		// Kill the object if initial creation failed.
		HRESULT hr = pA3dDal->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dDal;

		return hr;
	}

#ifdef _DEBUG
	if ((IID_IUnknown != rIid) && (IID_IA3d != rIid) && (IID_IA3d2 != rIid))
		LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// Only Unknown, A3d or A3d2 objects.
	if ((IID_IUnknown != rIid) && (IID_IA3d != rIid) && (IID_IA3d2 != rIid))
		return E_NOINTERFACE;

	// Return this object;
	*ppvObj = this;
	AddRef();
	
	return S_OK;
}

STDMETHODIMP_(ULONG) IA3dX::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dX::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dX::SetOutputMode
//
// Purpose: Set output mode.
//
// Parameters:
//  dwFrontXtalkMode DWORD front speaker mode.
//  dwBackXtalkMode DWORD back speaker mode.
//  dwQuadMode      DWORD quad mode.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::SetOutputMode(DWORD dwFrontXtalkMode,
	DWORD dwBackXtalkMode, DWORD dwQuadMode)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::SetOutputMode(%u,%u,%u)"),
		  dwFrontXtalkMode, dwBackXtalkMode, dwQuadMode);
	_ASSERTE(m_pDS);
#endif
	DWORD dwSpeakerConfig;

	// Check arguments values.
	if (OUTPUT_MODE_QUAD == dwQuadMode)
	{
		switch (dwFrontXtalkMode)
		{
		case OUTPUT_HEADPHONES:
			dwSpeakerConfig = DSSPEAKER_QUAD;
			break;
		case OUTPUT_SPEAKERS_NARROW:
			dwSpeakerConfig = DSSPEAKER_COMBINED(DSSPEAKER_QUAD, DSSPEAKER_GEOMETRY_NARROW);
			break;
		case OUTPUT_SPEAKERS_WIDE:
			dwSpeakerConfig = DSSPEAKER_COMBINED(DSSPEAKER_QUAD, DSSPEAKER_GEOMETRY_WIDE);
			break;
		default:
			return E_INVALIDARG;
		}
	}
	else if (OUTPUT_MODE_STEREO == dwQuadMode)
	{
		switch (dwFrontXtalkMode)
		{
		case OUTPUT_HEADPHONES:
			dwSpeakerConfig = DSSPEAKER_HEADPHONE;
			break;
		case OUTPUT_SPEAKERS_NARROW:
			dwSpeakerConfig = DSSPEAKER_COMBINED(DSSPEAKER_STEREO, DSSPEAKER_GEOMETRY_NARROW);
			break;
		case OUTPUT_SPEAKERS_WIDE:
			dwSpeakerConfig = DSSPEAKER_COMBINED(DSSPEAKER_STEREO, DSSPEAKER_GEOMETRY_WIDE);
			break;
		default:
			return E_INVALIDARG;
		}
	}
	else
		return E_INVALIDARG;

	// Set DirectSound speaker configuration.
#ifdef _DEBUG
	HRESULT hr = m_pDS->SetSpeakerConfig(dwSpeakerConfig);
	LogMsg(TEXT("...SetSpeakerConfig(%#x)=%s"), dwSpeakerConfig, Result(hr));
#else
	m_pDS->SetSpeakerConfig(dwSpeakerConfig);
#endif

	// Save output mode.
	m_dwFrontXtalkMode = dwFrontXtalkMode;
	m_dwBackXtalkMode = dwBackXtalkMode;
	m_dwQuadMode = dwQuadMode;

	return S_OK;
}


//===========================================================================
//
// IA3dX::GetOutputMode
//
// Purpose: Get output mode.
//
// Parameters:
//  pdwFrontXtalkMode LPDWORD pointer to front speaker mode buffer.
//  pdwBackXtalkMode LPDWORD pointer to back speaker mode buffer.
//  pdwQuadMode     LPDWORD pointer to quad mode buffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::GetOutputMode(LPDWORD pdwFrontXtalkMode,
	LPDWORD pdwBackXtalkMode, LPDWORD pdwQuadMode)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::GetOutputMode()"));
	_ASSERTE(pdwFrontXtalkMode && !IsBadWritePtr(pdwFrontXtalkMode, sizeof(*pdwFrontXtalkMode)));
	_ASSERTE(pdwBackXtalkMode && !IsBadWritePtr(pdwBackXtalkMode, sizeof(*pdwBackXtalkMode)));
	_ASSERTE(pdwQuadMode && !IsBadWritePtr(pdwQuadMode, sizeof(*pdwQuadMode)));
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!pdwFrontXtalkMode || !pdwBackXtalkMode || !pdwQuadMode)
		return E_POINTER;

	DWORD dwSpeakerConfig;

	// Get DirectSound speaker configuration.
	if (SUCCEEDED(m_pDS->GetSpeakerConfig(&dwSpeakerConfig)))
	{
		switch (dwSpeakerConfig)
		{
		case DSSPEAKER_HEADPHONE:
			m_dwFrontXtalkMode = OUTPUT_HEADPHONES;
			m_dwBackXtalkMode = OUTPUT_HEADPHONES;
			m_dwQuadMode = OUTPUT_MODE_STEREO;
			break;
		case DSSPEAKER_QUAD:
		case DSSPEAKER_SURROUND:
		case DSSPEAKER_5POINT1:
		case DSSPEAKER_7POINT1:
			m_dwFrontXtalkMode = OUTPUT_SPEAKERS_WIDE;
			m_dwBackXtalkMode = OUTPUT_SPEAKERS_WIDE;
			m_dwQuadMode = OUTPUT_MODE_QUAD;
			break;
		default:
			if (DSSPEAKER_GEOMETRY_WIDE >= DSSPEAKER_GEOMETRY(dwSpeakerConfig))
			{
				m_dwFrontXtalkMode = OUTPUT_SPEAKERS_WIDE;
				m_dwBackXtalkMode = OUTPUT_SPEAKERS_WIDE;
			}
			else
			{
				m_dwFrontXtalkMode = OUTPUT_SPEAKERS_NARROW;
				m_dwBackXtalkMode = OUTPUT_SPEAKERS_NARROW;
			}
			m_dwQuadMode = OUTPUT_MODE_STEREO;
		}
	}
#ifdef _DEBUG
	else
		LogMsg(TEXT("...GetSpeakerConfig() failed!"));

	LogMsg(TEXT("...FrontXtalkMode=%u"), m_dwFrontXtalkMode);
	LogMsg(TEXT("...BackXtalkMode=%u"), m_dwBackXtalkMode);
	LogMsg(TEXT("...QuadMode=%u"), m_dwQuadMode);
	LogMsg(TEXT("...=Ok"));
#endif

	// Copy output mode.
	*pdwFrontXtalkMode = m_dwFrontXtalkMode;
	*pdwBackXtalkMode = m_dwBackXtalkMode;
	*pdwQuadMode = m_dwQuadMode;

	return S_OK;
}


//===========================================================================
//
// IA3dX::SetResourceManagerMode
//
// Purpose: Set resource manager mode.
//
// Parameters:
//  dwResManMode    DWORD resource manager mode.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::SetResourceManagerMode(DWORD dwResManMode)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::SetResourceManagerMode(%u)"), dwResManMode);
	_ASSERTE(m_pDS);
#endif
	// Set resource manager mode for OEM.
	SetOemResourceManagerMode(dwResManMode);

	// Save resource manager mode.
	m_dwResourceManagerMode = dwResManMode;

	return S_OK;
}


//===========================================================================
//
// IA3dX::GetResourceManagerMode
//
// Purpose: Get resource manager mode.
//
// Parameters:
//  pdwResManMode   LPDWORD pointer to resource manager mode buffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::GetResourceManagerMode(LPDWORD pdwResManMode)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::GetResourceManagerMode()"));
	_ASSERTE(pdwResManMode && !IsBadWritePtr(pdwResManMode, sizeof(*pdwResManMode)));
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!pdwResManMode)
		return E_POINTER;

	// Get resource manager mode for OEM or return saved value.
	if (FAILED(GetOemResourceManagerMode(pdwResManMode)))
		*pdwResManMode = m_dwResourceManagerMode;

	return S_OK;
}


//===========================================================================
//
// IA3dX::SetHFAbsorbFactor
//
// Purpose: Set high frequencies absorb factor.
//
// Parameters:
//  fHFAbsorbFactor  FLOAT absorb factor.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::SetHFAbsorbFactor(FLOAT fHFAbsorbFactor)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::SetHFAbsorbFactor(%g)"), fHFAbsorbFactor);
#endif
	// Check arguments values.
	if (0.0f > fHFAbsorbFactor || 1.0f < fHFAbsorbFactor)
		return E_INVALIDARG;

	// Set high frequencies absorb factor for OEM.
	SetOemHFAbsorbFactor(fHFAbsorbFactor);

	// Save HF absorb factor.
	m_fHFAbsorbFactor = fHFAbsorbFactor;

	return S_OK;
}


//===========================================================================
//
// IA3dX::GetHFAbsorbFactor
//
// Purpose: Get high frequencies absorb factor.
//
// Parameters:
//  pfHFAbsorbFactor FLOAT* pointer to absorb factor.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::GetHFAbsorbFactor(FLOAT* pfHFAbsorbFactor)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::GetHFAbsorbFactor()..."));
	_ASSERTE(pfHFAbsorbFactor && !IsBadWritePtr(pfHFAbsorbFactor, sizeof(*pfHFAbsorbFactor)));
#endif
	// Check arguments values.
	if (!pfHFAbsorbFactor)
		return E_POINTER;

	// Return saved HF absorb factor.
	*pfHFAbsorbFactor = m_fHFAbsorbFactor;
#ifdef _DEBUG
	LogMsg(TEXT("...HFAbsorbFactor=%g"), m_fHFAbsorbFactor);
#endif

	return S_OK;
}


//===========================================================================
//
// IA3dX::RegisterVersion
//
// Purpose: Register application A3D version.
//
// Parameters:
//  dwAppVersion    DWORD application version.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::RegisterVersion(DWORD dwAppVersion)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::RegisterVersion(%u)"), dwAppVersion);
#endif
	// Check arguments values.
	if (IA3DVERSION_RELEASE10 > dwAppVersion || IA3DVERSION_RELEASE12 < dwAppVersion)
		return E_INVALIDARG;

	// Save application version.
	m_dwAppVersion = dwAppVersion;

	return S_OK;
}


//===========================================================================
//
// IA3dX::GetSoftwareCaps
//
// Purpose: Get software capabilities.
//
// Parameters:
//  pSoftwareCaps   LPA3DCAPS_SOFTWARE pointer to software capabilities buffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::GetSoftwareCaps(LPA3DCAPS_SOFTWARE pSoftwareCaps)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::GetSoftwareCaps()..."));
	_ASSERTE(pSoftwareCaps && !IsBadWritePtr(pSoftwareCaps, sizeof(*pSoftwareCaps)));
	LogMsg(TEXT("...Size=%u"), pSoftwareCaps->dwSize);
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!pSoftwareCaps)
		return E_POINTER;

	// Get capabilities buffer size.
	DWORD dwSize = pSoftwareCaps->dwSize;
	if (!dwSize)
		return E_INVALIDARG;

	// Set limit capabilities buffer size.
	if (dwSize > sizeof(A3DCAPS_SOFTWARE))
		dwSize = sizeof(A3DCAPS_SOFTWARE);

	// Fill default software capabilities values.
	A3DCAPS_SOFTWARE A3dCapsSoftware;
	ZeroMemory(&A3dCapsSoftware, sizeof(A3dCapsSoftware));
	A3dCapsSoftware.dwSize = dwSize;
	A3dCapsSoftware.dwVersion = IA3DVERSION_RELEASE20;
	A3dCapsSoftware.dwFlags = A3D_DIRECT_PATH_A3D | A3D_SOFTWARECAPS_DEFAULT;
	A3dCapsSoftware.dwOutputChannels = 2;
	A3dCapsSoftware.dwMinSampleRate = DSBFREQUENCY_MIN;
	A3dCapsSoftware.dwMaxSampleRate = DSBFREQUENCY_MAX;
	A3dCapsSoftware.dwMax2DBuffers = 1;

	DSCAPS DSCaps;
	DSCaps.dwSize = sizeof(DSCaps);

	// Get DirectSound capabilities.
	if (SUCCEEDED(m_pDS->GetCaps(&DSCaps)))
	{
		A3dCapsSoftware.dwMinSampleRate = DSCaps.dwMinSecondarySampleRate;
		A3dCapsSoftware.dwMaxSampleRate = DSCaps.dwMaxSecondarySampleRate;
		A3dCapsSoftware.dwMax2DBuffers = DSCaps.dwMaxHwMixingAllBuffers;
		A3dCapsSoftware.dwMax3DBuffers = DSCaps.dwMaxHw3DAllBuffers;
	}
#ifdef _DEBUG
	else
		LogMsg(TEXT("...GetCaps() failed!"));
#endif

	// Copy software capabilities.
	CopyMemory(pSoftwareCaps, &A3dCapsSoftware, dwSize);

	return S_OK;
}


//===========================================================================
//
// IA3dX::GetHardwareCaps
//
// Purpose: Get hardware capabilities.
//
// Parameters:
//  pHardwareCaps   LPA3DCAPS_HARDWARE pointer to hardware capabilities buffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dX::GetHardwareCaps(LPA3DCAPS_HARDWARE pHardwareCaps)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dX::GetHardwareCaps()..."));
	_ASSERTE(pHardwareCaps && !IsBadWritePtr(pHardwareCaps, sizeof(*pHardwareCaps)));
	LogMsg(TEXT("...Size=%u"), pHardwareCaps->dwSize);
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!pHardwareCaps)
		return E_POINTER;

	// Get capabilities buffer size.
	DWORD dwSize = pHardwareCaps->dwSize;
	if (!dwSize)
		return E_INVALIDARG;

	// Set limit capabilities buffer size.
	if (dwSize > sizeof(A3DCAPS_HARDWARE))
		dwSize = sizeof(A3DCAPS_HARDWARE);

	// Fill default hardware capabilities values.
	A3DCAPS_HARDWARE A3dCapsHardware;
	ZeroMemory(&A3dCapsHardware, sizeof(A3dCapsHardware));
	A3dCapsHardware.dwSize = dwSize;
	A3dCapsHardware.dwFlags = A3D_DIRECT_PATH_A3D | A3D_HARDWARECAPS_DEFAULT;
	A3dCapsHardware.dwOutputChannels = 2;
	A3dCapsHardware.dwMinSampleRate = DSBFREQUENCY_MIN;
	A3dCapsHardware.dwMaxSampleRate = DSBFREQUENCY_MAX;
	A3dCapsHardware.dwMax2DBuffers = 1;

	DSCAPS DSCaps;
	DSCaps.dwSize = sizeof(DSCaps);

	// Get DirectSound capabilities.
	if (SUCCEEDED(m_pDS->GetCaps(&DSCaps)))
	{
		A3dCapsHardware.dwMinSampleRate = DSCaps.dwMinSecondarySampleRate;
		A3dCapsHardware.dwMaxSampleRate = DSCaps.dwMaxSecondarySampleRate;
		A3dCapsHardware.dwMax2DBuffers = DSCaps.dwMaxHwMixingAllBuffers;
		A3dCapsHardware.dwMax3DBuffers = DSCaps.dwMaxHw3DAllBuffers;
	}
#ifdef _DEBUG
	else
		LogMsg(TEXT("...GetCaps() failed!"));
#endif

	// Copy hardware capabilities.
	CopyMemory(pHardwareCaps, &A3dCapsHardware, dwSize);

	return S_OK;
}


//===========================================================================
//
// IA3dSound::IA3dSound
// IA3dSound::~IA3dSound
//
// Constructor Parameters:
//  pDS             LPDIRECTSOUND to the parent object.
//
//===========================================================================
IA3dSound::IA3dSound(LPDIRECTSOUND pDS) :
	m_cRef(0),
	m_pDS(pDS)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound::IA3dSound()=%u"), g_cObj + 1);
#endif
	// Add reference for parent DirectSound object.
	if (m_pDS)
		m_pDS->AddRef();
	else
		CoCreateInstance(CLSID_DirectSound, NULL, CLSCTX_INPROC_SERVER,
			IID_IDirectSound, (LPVOID *)&m_pDS);

	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dSound::~IA3dSound()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound::~IA3dSound()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Release parent DirectSound object.
	if (m_pDS)
		m_pDS->Release();

	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dSound::QueryInterface
// IA3dSound::AddRef
// IA3dSound::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dSound::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dSound::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

	// If not exist parent DirectSound object.
	if (!m_pDS)
		return E_FAIL;

	// Request for A3dDal object.
	if (IID_IA3dDal == rIid)
	{
		// Create new A3dDal object.
		LPA3DDAL pA3dDal = new IA3dDal(m_pDS);
		if (!pA3dDal)
			return E_OUTOFMEMORY;

		// Kill the object if initial creation failed.
		HRESULT hr = pA3dDal->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dDal;

		return hr;
	}

	// Request for A3d or A3d2 objects.
	if ((IID_IA3d == rIid) || (IID_IA3d2 == rIid))
	{
		// Create new A3dX object.
		LPA3DX pA3dX = new IA3dX(m_pDS);
		if (!pA3dX)
			return E_OUTOFMEMORY;

		// Kill the object if initial creation or Init failed.
		HRESULT hr = pA3dX->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dX;

		return hr;
	}

	// Request for Unknown or DirectSound objects.
	if ((IID_IUnknown == rIid) || (IID_IDirectSound == rIid))
	{
		// Return this object;
		*ppvObj = this;
		AddRef();

		return S_OK;
	}

#ifdef _DEBUG
	LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// All other requests redirect to DirectSound.
	return m_pDS->QueryInterface(rIid, ppvObj);
}

STDMETHODIMP_(ULONG) IA3dSound::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dSound::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dSound::<All DirectSound class methods>
//
// Purpose: All class methods redirect to DirectSound.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dSound::CreateSoundBuffer(LPCDSBUFFERDESC pcDSBufferDesc,
	LPDIRECTSOUNDBUFFER * ppDirectSoundBuffer, LPUNKNOWN pUnkOuter)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound::CreateSoundBuffer()..."));
	_ASSERTE(pcDSBufferDesc && !IsBadReadPtr(pcDSBufferDesc, sizeof(*pcDSBufferDesc)));
	LogMsg(TEXT("...Size=%u"), pcDSBufferDesc->dwSize);
	LogMsg(TEXT("...Flags=%#x"), pcDSBufferDesc->dwFlags);
	if (pcDSBufferDesc->dwFlags & DSBCAPS_PRIMARYBUFFER)
		LogMsg(TEXT("......DSBCAPS_PRIMARYBUFFER"));
	if (pcDSBufferDesc->dwFlags & DSBCAPS_LOCHARDWARE)
		LogMsg(TEXT("......DSBCAPS_LOCHARDWARE"));
	if (pcDSBufferDesc->dwFlags & DSBCAPS_LOCSOFTWARE)
		LogMsg(TEXT("......DSBCAPS_LOCSOFTWARE"));
	if (pcDSBufferDesc->dwFlags & DSBCAPS_CTRL3D)
		LogMsg(TEXT("......DSBCAPS_CTRL3D"));
	LogMsg(TEXT("...BufferBytes=%u"), pcDSBufferDesc->dwBufferBytes);
	if (pcDSBufferDesc->dwReserved)
		LogMsg(TEXT("...Reserved=%#x"), pcDSBufferDesc->dwReserved);
	if (pcDSBufferDesc->lpwfxFormat)
	{
		LogMsg(TEXT("...Channels=%u"), pcDSBufferDesc->lpwfxFormat->nChannels);
		LogMsg(TEXT("...SamplesPerSec=%u"), pcDSBufferDesc->lpwfxFormat->nSamplesPerSec);
		LogMsg(TEXT("...AvgBytesPerSec=%u"), pcDSBufferDesc->lpwfxFormat->nAvgBytesPerSec);
		LogMsg(TEXT("...BlockAlign=%u"), pcDSBufferDesc->lpwfxFormat->nBlockAlign);
		LogMsg(TEXT("...BitsPerSample=%u"), pcDSBufferDesc->lpwfxFormat->wBitsPerSample);
	}
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!pcDSBufferDesc)
		return E_POINTER;

	HRESULT hr;

	// Create new DirectSoundBuffer object.
	if ((pcDSBufferDesc->dwFlags & DSBCAPS_PRIMARYBUFFER) ||
	!(pcDSBufferDesc->dwFlags & DSBCAPS_CTRL3D) ||
	(pcDSBufferDesc->dwFlags & DSBCAPS_LOCSOFTWARE))
		hr = m_pDS->CreateSoundBuffer(pcDSBufferDesc, ppDirectSoundBuffer, pUnkOuter);
	else
	{
		// Use hardware sound buffers, only if they available.
		DSBUFFERDESC DSBufDesc;
		CopyMemory(&DSBufDesc, pcDSBufferDesc, sizeof(DSBufDesc));
		DSBufDesc.dwSize = sizeof(DSBufDesc);
		DSBufDesc.dwFlags |= DSBCAPS_CTRLVOLUME | DSBCAPS_CTRLFREQUENCY |
			DSBCAPS_CTRLPOSITIONNOTIFY;
#ifdef _DEBUG
		LogMsg(TEXT("...Flags=%#x!"), DSBufDesc.dwFlags);
#endif
		DSBufDesc.guid3DAlgorithm = DS3DALG_HRTF_FULL;
		hr = m_pDS->CreateSoundBuffer(&DSBufDesc,  ppDirectSoundBuffer, pUnkOuter);
	}
	if (SUCCEEDED(hr))
	{
		// Show splash only for primary 3D sound buffer.
		if ((pcDSBufferDesc->dwFlags & DSBCAPS_CTRL3D) &&
		(pcDSBufferDesc->dwFlags & DSBCAPS_PRIMARYBUFFER))
			SplashScreen(g_hWnd);

		// Create new A3dSoundBuffer object.
		LPA3DSOUNDBUFFER pA3dSoundBuffer = new IA3dSoundBuffer(*ppDirectSoundBuffer, m_pDS);
		if (!pA3dSoundBuffer)
		{
			(*ppDirectSoundBuffer)->Release();
			*ppDirectSoundBuffer = NULL;
			return E_OUTOFMEMORY;
		}

		// Kill the object if initial creation failed.
		hr = pA3dSoundBuffer->QueryInterface(IID_IDirectSoundBuffer, (LPVOID *)ppDirectSoundBuffer);
		if (FAILED(hr))
			delete pA3dSoundBuffer;
	}

#ifdef _DEBUG
	LogMsg(TEXT("...=%s"), Result(hr));
#endif
	return hr;
}

STDMETHODIMP IA3dSound::GetCaps(LPDSCAPS pDSCaps)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound::GetCaps()..."));
	_ASSERTE(m_pDS);
	HRESULT hr = m_pDS->GetCaps(pDSCaps);
	if (SUCCEEDED(hr))
	{
		LogMsg(TEXT("...Flags=%#x"), pDSCaps->dwFlags);
		LogMsg(TEXT("...MinSecondarySampleRate=%u"), pDSCaps->dwMinSecondarySampleRate);
		LogMsg(TEXT("...MaxSecondarySampleRate=%u"), pDSCaps->dwMaxSecondarySampleRate);
		LogMsg(TEXT("...PrimaryBuffers=%u"), pDSCaps->dwPrimaryBuffers);
		LogMsg(TEXT("...MaxHwMixingAllBuffers=%u"), pDSCaps->dwMaxHwMixingAllBuffers);
		LogMsg(TEXT("...MaxHwMixingStaticBuffers=%u"), pDSCaps->dwMaxHwMixingStaticBuffers);
		LogMsg(TEXT("...MaxHwMixingStreamingBuffers=%u"), pDSCaps->dwMaxHwMixingStreamingBuffers);
		LogMsg(TEXT("...FreeHwMixingAllBuffers=%u"), pDSCaps->dwFreeHwMixingAllBuffers);
		LogMsg(TEXT("...FreeHwMixingStaticBuffers=%u"), pDSCaps->dwFreeHwMixingStaticBuffers);
		LogMsg(TEXT("...FreeHwMixingStreamingBuffers=%u"), pDSCaps->dwFreeHwMixingStreamingBuffers);
		LogMsg(TEXT("...MaxHw3DAllBuffers=%u"), pDSCaps->dwMaxHw3DAllBuffers);
		LogMsg(TEXT("...MaxHw3DStaticBuffers=%u"), pDSCaps->dwMaxHw3DStaticBuffers);
		LogMsg(TEXT("...MaxHw3DStreamingBuffers=%u"), pDSCaps->dwMaxHw3DStreamingBuffers);
		LogMsg(TEXT("...FreeHw3DAllBuffers=%u"), pDSCaps->dwFreeHw3DAllBuffers);
		LogMsg(TEXT("...FreeHw3DStaticBuffers=%u"), pDSCaps->dwFreeHw3DStaticBuffers);
		LogMsg(TEXT("...FreeHw3DStreamingBuffers=%u"), pDSCaps->dwFreeHw3DStreamingBuffers);
		LogMsg(TEXT("...TotalHwMemBytes=%u"), pDSCaps->dwTotalHwMemBytes);
		LogMsg(TEXT("...FreeHwMemBytes=%u"), pDSCaps->dwFreeHwMemBytes);
		LogMsg(TEXT("...MaxContigFreeHwMemBytes=%u"), pDSCaps->dwMaxContigFreeHwMemBytes);
		LogMsg(TEXT("...UnlockTransferRateHwBuffers=%u"), pDSCaps->dwUnlockTransferRateHwBuffers);
		LogMsg(TEXT("...PlayCpuOverheadSwBuffers=%u"), pDSCaps->dwPlayCpuOverheadSwBuffers);
	}

	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDS->GetCaps(pDSCaps);
#endif
}

STDMETHODIMP IA3dSound::DuplicateSoundBuffer(LPDIRECTSOUNDBUFFER pDSBufferOriginal,
	LPDIRECTSOUNDBUFFER * ppDSBufferDuplicate)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound::DuplicateSoundBuffer()..."));
	_ASSERTE(m_pDS);
#endif
	// Duplicate DirectSoundBuffer object.
	HRESULT hr = m_pDS->DuplicateSoundBuffer(pDSBufferOriginal, ppDSBufferDuplicate);
	if (SUCCEEDED(hr))
	{
		// Create new A3dSoundBuffer object.
		LPA3DSOUNDBUFFER pA3dSoundBuffer = new IA3dSoundBuffer(*ppDSBufferDuplicate, m_pDS);
		if (!pA3dSoundBuffer)
		{
			(*ppDSBufferDuplicate)->Release();
			*ppDSBufferDuplicate = NULL;
			return E_OUTOFMEMORY;
		}

		// Kill the object if initial creation failed.
		hr = pA3dSoundBuffer->QueryInterface(IID_IDirectSoundBuffer, (LPVOID *)ppDSBufferDuplicate);
		if (FAILED(hr))
			delete pA3dSoundBuffer;
	}

#ifdef _DEBUG
	LogMsg(TEXT("...=%s"), Result(hr));
#endif
	return hr;
}

STDMETHODIMP IA3dSound::SetCooperativeLevel(HWND hWnd, DWORD dwLevel)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS);
	// Save handle of application window.
	if (!g_hWnd)
		g_hWnd = hWnd;
	HRESULT hr = m_pDS->SetCooperativeLevel(hWnd, dwLevel);
	LogMsg(TEXT("IA3dSound::SetCooperativeLevel(%u)=%s"), dwLevel, Result(hr));
	return hr;
#else
	// Save handle of application window.
	if (!g_hWnd)
		g_hWnd = hWnd;
	return m_pDS->SetCooperativeLevel(hWnd, dwLevel);
#endif
}

STDMETHODIMP IA3dSound::Compact()
{
#ifdef _DEBUG
	_ASSERTE(m_pDS);
	HRESULT hr = m_pDS->Compact();
	LogMsg(TEXT("IA3dSound::Compact()=%s"), Result(hr));
	return hr;
#else
	return m_pDS->Compact();
#endif
}

STDMETHODIMP IA3dSound::GetSpeakerConfig(LPDWORD pdwSpeakerConfig)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound::GetSpeakerConfig()..."));
	_ASSERTE(m_pDS);
	HRESULT hr = m_pDS->GetSpeakerConfig(pdwSpeakerConfig);
	if (SUCCEEDED(hr))
		LogMsg(TEXT("...SpeakerConfig=%#x"), *pdwSpeakerConfig);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDS->GetSpeakerConfig(pdwSpeakerConfig);
#endif
}

STDMETHODIMP IA3dSound::SetSpeakerConfig(DWORD dwSpeakerConfig)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS);
	HRESULT hr = m_pDS->SetSpeakerConfig(dwSpeakerConfig);
	LogMsg(TEXT("IA3dSound::SetSpeakerConfig(%#x)=%s"), dwSpeakerConfig, Result(hr));
	return hr;
#else
	return m_pDS->SetSpeakerConfig(dwSpeakerConfig);
#endif
}

STDMETHODIMP IA3dSound::Initialize(LPCGUID pcGuidDevice)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS);
	HRESULT hr = m_pDS->Initialize(pcGuidDevice);
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dSound::Initialize(%s)=%s"),
		GuidToStr((REFGUID)*pcGuidDevice, szGuid, sizeof(szGuid)), Result(hr));
	return hr;
#else
	return m_pDS->Initialize(pcGuidDevice);
#endif
}


//===========================================================================
//
// IA3dSoundBuffer::IA3dSoundBuffer
// IA3dSoundBuffer::~IA3dSoundBuffer
//
// Constructor Parameters:
//  pDS             LPDIRECTSOUND to the parent object.
//  pDSB            LPDIRECTSOUNDBUFFER to the parent object.
//
//===========================================================================
IA3dSoundBuffer::IA3dSoundBuffer(LPDIRECTSOUNDBUFFER pDSB, LPDIRECTSOUND pDS) :
	m_cRef(0),
	m_pDSB(pDSB),
	m_pDS(pDS)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::IA3dSoundBuffer()=%u"), g_cObj + 1);
#endif
	// Add reference for parent DirectSound object.
	if (m_pDS)
		m_pDS->AddRef();

	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dSoundBuffer::~IA3dSoundBuffer()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::~IA3dSoundBuffer()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Release DirectSoundBuffer object.
	if (m_pDSB)
		m_pDSB->Release();

	// Release DirectSound object.
	if (m_pDS)
		m_pDS->Release();

	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dSoundBuffer::QueryInterface
// IA3dSoundBuffer::AddRef
// IA3dSoundBuffer::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dSoundBuffer::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dSoundBuffer::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
	_ASSERTE(m_pDSB);
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

	// If not exist parent DirectSoundBuffer object.
	if (!m_pDSB)
		return E_FAIL;

	// Request for A3dScaleHackBuffer object.
	if (IID_IA3dScaleHackBuffer == rIid)
	{
		// Create A3dScaleHackBuffer object.
		LPA3DSCALEHACKBUFFER pA3dScaleHackBuffer = new IA3dScaleHackBuffer;
		if (!pA3dScaleHackBuffer)
			return E_OUTOFMEMORY;

		// Kill the object if initial creation failed.
		HRESULT hr = pA3dScaleHackBuffer->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dScaleHackBuffer;

		return hr;
	}

	// Request for A3dDalBuffer object.
	if (IID_IA3dDalBuffer == rIid)
	{
		// Create new A3dDalBuffer object.
		LPA3DDALBUFFER pA3dDalBuffer = new IA3dDalBuffer(m_pDS, m_pDSB);
		if (!pA3dDalBuffer)
			return E_OUTOFMEMORY;

		// Kill the object if initial creation failed.
		HRESULT hr = pA3dDalBuffer->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dDalBuffer;

		return hr;
	}

	// Request for DirectSound3DListener object.
	if (IID_IDirectSound3DListener == rIid)
	{
		LPDIRECTSOUND3DLISTENER pDS3DL;

		// Get DirectSound3DListener object.
		HRESULT hr = m_pDSB->QueryInterface(rIid, (LPVOID *)&pDS3DL);
		if (FAILED(hr))
			return hr;

		// Create new A3dSound3DListener object.
		LPA3DSOUND3DLISTENER pA3dSound3DListener = new IA3dSound3DListener(pDS3DL);
		if (!pA3dSound3DListener)
		{
			pDS3DL->Release();
			return E_OUTOFMEMORY;
		}

		// Kill the object if initial creation failed.
		hr = pA3dSound3DListener->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dSound3DListener;

		return hr;
	}

	// Request for DirectSound3DBuffer object.
	if (IID_IDirectSound3DBuffer == rIid)
	{
		LPDIRECTSOUND3DBUFFER pDS3DB;

		// Get DirectSound3DBuffer object.
		HRESULT hr = m_pDSB->QueryInterface(rIid, (LPVOID *)&pDS3DB);
		if (FAILED(hr))
			return hr;

		// Create new A3dSound3DBuffer object.
		LPA3DSOUND3DBUFFER pA3dSound3DBuffer = new IA3dSound3DBuffer(pDS3DB);
		if (!pA3dSound3DBuffer)
		{
			pDS3DB->Release();
			return E_OUTOFMEMORY;
		}

		// Kill the object if initial creation failed.
		hr = pA3dSound3DBuffer->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dSound3DBuffer;

		return hr;
	}

	// Request for KsPropertySet object.
	if (IID_IKsPropertySet == rIid)
	{
		LPKSPROPERTYSET pKsPS;

		// Get KsPropertySet object.
		HRESULT hr = m_pDSB->QueryInterface(rIid, (LPVOID *)&pKsPS);
		if (FAILED(hr))
			return hr;

		// Create new A3dKsPropertySet object.
		LPA3DKSPROPERTYSET pA3dKsPropertySet = new IA3dKsPropertySet(pKsPS);
		if (!pA3dKsPropertySet)
		{
			pKsPS->Release();
			return E_OUTOFMEMORY;
		}

		// Kill the object if initial creation failed.
		hr = pA3dKsPropertySet->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dKsPropertySet;

		return hr;
	}

	// Request for Unknown or DirectSoundBuffer objects.
	if ((IID_IUnknown == rIid) || (IID_IDirectSoundBuffer == rIid))
	{
		// Return this object;
		*ppvObj = this;
		AddRef();

		return S_OK;
	}

#ifdef _DEBUG
	LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// All other requests redirect to DirectSoundBuffer.
	return m_pDSB->QueryInterface(rIid, ppvObj);
}

STDMETHODIMP_(ULONG) IA3dSoundBuffer::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dSoundBuffer::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dSoundBuffer::<All DirectSoundBuffer class methods>
//
// Purpose: All class methods redirect to DirectSoundBuffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dSoundBuffer::GetCaps(LPDSBCAPS pDSBCaps)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::GetCaps()..."));
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->GetCaps(pDSBCaps);
	if (SUCCEEDED(hr))
	{
		LogMsg(TEXT("...Size=%u"), pDSBCaps->dwSize);
		LogMsg(TEXT("...Flags=%#x"), pDSBCaps->dwFlags);
		LogMsg(TEXT("...BufferBytes=%u"), pDSBCaps->dwBufferBytes);
		LogMsg(TEXT("...UnlockTransferRate=%u"), pDSBCaps->dwUnlockTransferRate);
		LogMsg(TEXT("...PlayCpuOverhead=%u"), pDSBCaps->dwPlayCpuOverhead);
	}

	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->GetCaps(pDSBCaps);
#endif
}

STDMETHODIMP IA3dSoundBuffer::GetCurrentPosition(LPDWORD pdwCurrentPlayCursor,
	LPDWORD pdwCurrentWriteCursor)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::GetCurrentPosition()..."));
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->GetCurrentPosition(pdwCurrentPlayCursor, pdwCurrentWriteCursor);
	if (SUCCEEDED(hr))
	{
		LogMsg(TEXT("...CurrentPlayCursor=%u"), *pdwCurrentPlayCursor);
		LogMsg(TEXT("...CurrentWriteCursor=%u"), *pdwCurrentPlayCursor);
	}

	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->GetCurrentPosition(pdwCurrentPlayCursor, pdwCurrentWriteCursor);
#endif
}

STDMETHODIMP IA3dSoundBuffer::GetFormat(LPWAVEFORMATEX pwfxFormat,
	DWORD dwSizeAllocated, LPDWORD pdwSizeWritten)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::GetFormat()..."));
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->GetFormat(pwfxFormat, dwSizeAllocated, pdwSizeWritten);
	if (SUCCEEDED(hr))
	{
		LogMsg(TEXT("...Channels=%u"), pwfxFormat->nChannels);
		LogMsg(TEXT("...SamplesPerSec=%u"), pwfxFormat->nSamplesPerSec);
		LogMsg(TEXT("...AvgBytesPerSec=%u"), pwfxFormat->nAvgBytesPerSec);
		LogMsg(TEXT("...BlockAlign=%u"), pwfxFormat->nBlockAlign);
		LogMsg(TEXT("...BitsPerSample=%u"), pwfxFormat->wBitsPerSample);
	}

	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->GetFormat(pwfxFormat, dwSizeAllocated, pdwSizeWritten);
#endif
}

STDMETHODIMP IA3dSoundBuffer::GetVolume(LPLONG plVolume)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::GetVolume()..."));
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->GetVolume(plVolume);
	if (SUCCEEDED(hr))
		LogMsg(TEXT("...Volume=%d"), *plVolume);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->GetVolume(plVolume);
#endif
}

STDMETHODIMP IA3dSoundBuffer::GetPan(LPLONG plPan)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::GetPan()..."));
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->GetPan(plPan);
	if (SUCCEEDED(hr))
		LogMsg(TEXT("...Pan=%d"), *plPan);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->GetPan(plPan);
#endif
}

STDMETHODIMP IA3dSoundBuffer::GetFrequency(LPDWORD pdwFrequency)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::GetFrequency()..."));
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->GetFrequency(pdwFrequency);
	if (SUCCEEDED(hr))
		LogMsg(TEXT("...Frequency=%u"), *pdwFrequency);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->GetFrequency(pdwFrequency);
#endif
}

STDMETHODIMP IA3dSoundBuffer::GetStatus(LPDWORD pdwStatus)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::GetStatus()..."));
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->GetStatus(pdwStatus);
	if (SUCCEEDED(hr))
		LogMsg(TEXT("...Status=%#x"), *pdwStatus);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->GetStatus(pdwStatus);
#endif
}

STDMETHODIMP IA3dSoundBuffer::Initialize(LPDIRECTSOUND pDirectSound,
	LPCDSBUFFERDESC pcDSBufferDesc)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::Initialize()..."));
	_ASSERTE(pcDSBufferDesc && !IsBadReadPtr(pcDSBufferDesc, sizeof(pcDSBufferDesc)));
	LogMsg(TEXT("...Size=%u"), pcDSBufferDesc->dwSize);
	LogMsg(TEXT("...Flags=%#x"), pcDSBufferDesc->dwFlags);
	LogMsg(TEXT("...BufferBytes=%u"), pcDSBufferDesc->dwBufferBytes);
	if (pcDSBufferDesc->dwReserved)
		LogMsg(TEXT("...Reserved=%#x"), pcDSBufferDesc->dwReserved);
	if (pcDSBufferDesc->lpwfxFormat)
	{
		LogMsg(TEXT("...Channels=%u"), pcDSBufferDesc->lpwfxFormat->nChannels);
		LogMsg(TEXT("...SamplesPerSec=%u"), pcDSBufferDesc->lpwfxFormat->nSamplesPerSec);
		LogMsg(TEXT("...AvgBytesPerSec=%u"), pcDSBufferDesc->lpwfxFormat->nAvgBytesPerSec);
		LogMsg(TEXT("...BlockAlign=%u"), pcDSBufferDesc->lpwfxFormat->nBlockAlign);
		LogMsg(TEXT("...BitsPerSample=%u"), pcDSBufferDesc->lpwfxFormat->wBitsPerSample);
	}

	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->Initialize(pDirectSound, pcDSBufferDesc);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->Initialize(pDirectSound, pcDSBufferDesc);
#endif
}

STDMETHODIMP IA3dSoundBuffer::Lock(DWORD dwWriteCursor, DWORD dwWriteBytes,
	LPVOID *ppvAudioPtr1, LPDWORD pdwAudioBytes1, LPVOID *ppvAudioPtr2,
	LPDWORD pdwAudioBytes2, DWORD dwFlags)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::Lock(%u,%u,%#x)..."), dwWriteCursor,
		dwWriteBytes, dwFlags);
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->Lock(dwWriteCursor, dwWriteBytes, ppvAudioPtr1, pdwAudioBytes1,
		ppvAudioPtr2, pdwAudioBytes2, dwFlags);
	if (SUCCEEDED(hr))
	{
		LogMsg(TEXT("...AudioPtr1=%#x"), *ppvAudioPtr1);
		LogMsg(TEXT("...AudioBytes1=%u"), *pdwAudioBytes1);
		if (ppvAudioPtr2)
			LogMsg(TEXT("...AudioPtr2=%#x"), *ppvAudioPtr2);
		if (pdwAudioBytes2)
			LogMsg(TEXT("...AudioBytes2=%u"), *pdwAudioBytes2);
	}

	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return  m_pDSB->Lock(dwWriteCursor, dwWriteBytes, ppvAudioPtr1, pdwAudioBytes1,
		ppvAudioPtr2, pdwAudioBytes2, dwFlags);
#endif
}

STDMETHODIMP IA3dSoundBuffer::Play(DWORD dwReserved1, DWORD dwPriority, DWORD dwFlags)
{
#ifdef _DEBUG
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->Play(dwReserved1, dwPriority, dwFlags);
	LogMsg(TEXT("IA3dSoundBuffer::Play(%#x,%#x,%#x)=%s"), dwReserved1, dwPriority,
		dwFlags, Result(hr));
	return hr;
#else
	return m_pDSB->Play(dwReserved1, dwPriority, dwFlags);
#endif
}

STDMETHODIMP IA3dSoundBuffer::SetCurrentPosition(DWORD dwNewPosition)
{
#ifdef _DEBUG
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->SetCurrentPosition(dwNewPosition);
	LogMsg(TEXT("IA3dSoundBuffer::SetCurrentPosition(%d)=%s"), dwNewPosition, Result(hr));
	return hr;
#else
	return m_pDSB->SetCurrentPosition(dwNewPosition);
#endif
}

STDMETHODIMP IA3dSoundBuffer::SetFormat(LPCWAVEFORMATEX pcfxFormat)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSoundBuffer::SetFormat()..."));
	_ASSERTE(pcfxFormat && !IsBadReadPtr(pcfxFormat, sizeof(*pcfxFormat)));
	LogMsg(TEXT("...Channels=%u"), pcfxFormat->nChannels);
	LogMsg(TEXT("...SamplesPerSec=%u"), pcfxFormat->nSamplesPerSec);
	LogMsg(TEXT("...AvgBytesPerSec=%u"), pcfxFormat->nAvgBytesPerSec);
	LogMsg(TEXT("...BlockAlign=%u"), pcfxFormat->nBlockAlign);
	LogMsg(TEXT("...BitsPerSample=%u"), pcfxFormat->wBitsPerSample);

	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->SetFormat(pcfxFormat);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->SetFormat(pcfxFormat);
#endif
}

STDMETHODIMP IA3dSoundBuffer::SetVolume(LONG lVolume)
{
#ifdef _DEBUG
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->SetVolume(lVolume);
	LogMsg(TEXT("IA3dSoundBuffer::SetVolume(%d)=%s"), lVolume, Result(hr));
	return hr;
#else
	return m_pDSB->SetVolume(lVolume);
#endif
}

STDMETHODIMP IA3dSoundBuffer::SetPan(LONG lPan)
{
#ifdef _DEBUG
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->SetPan(lPan);
	LogMsg(TEXT("IA3dSoundBuffer::SetPan(%d)=%s"), lPan, Result(hr));
	return hr;
#else
	return m_pDSB->SetPan(lPan);
#endif
}

STDMETHODIMP IA3dSoundBuffer::SetFrequency(DWORD dwFrequency)
{
#ifdef _DEBUG
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->SetFrequency(dwFrequency);
	LogMsg(TEXT("IA3dSoundBuffer::SetFrequency(%u)=%s"), dwFrequency, Result(hr));
	return hr;
#else
	return m_pDSB->SetFrequency(dwFrequency);
#endif
}

STDMETHODIMP IA3dSoundBuffer::Stop()
{
#ifdef _DEBUG
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->Stop();
	LogMsg(TEXT("IA3dSoundBuffer::Stop()=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->Stop();
#endif
}

STDMETHODIMP IA3dSoundBuffer::Unlock(LPVOID pvAudioPtr1, DWORD dwAudioBytes1,
	LPVOID pvAudioPtr2, DWORD dwAudioBytes2)
{
#ifdef _DEBUG
	_ASSERTE(m_pDSB);
	// Correct invalid parameter for A3DAPI.DLL calls.
	if (pvAudioPtr1 == pvAudioPtr2 && !dwAudioBytes2)
	{
		LogMsg(TEXT("IA3dSoundBuffer::Unlock(%#x,%u,%#x!,%u)..."), pvAudioPtr1,
			dwAudioBytes1, pvAudioPtr2, dwAudioBytes2);
		pvAudioPtr2 = NULL;
	}

	HRESULT hr = m_pDSB->Unlock(pvAudioPtr1, dwAudioBytes1, pvAudioPtr2, dwAudioBytes2);
	LogMsg(TEXT("IA3dSoundBuffer::Unlock(%#x,%u,%#x,%u)=%s"), pvAudioPtr1,
		dwAudioBytes1, pvAudioPtr2, dwAudioBytes2, Result(hr));
	return hr;
#else
	// Correct invalid parameter for A3DAPI.DLL calls.
	if (pvAudioPtr1 == pvAudioPtr2 && !dwAudioBytes2)
		pvAudioPtr2 = NULL;

	return m_pDSB->Unlock(pvAudioPtr1, dwAudioBytes1, pvAudioPtr2, dwAudioBytes2);
#endif
}

STDMETHODIMP IA3dSoundBuffer::Restore()
{
#ifdef _DEBUG
	_ASSERTE(m_pDSB);
	HRESULT hr = m_pDSB->Restore();
	LogMsg(TEXT("IA3dSoundBuffer::Restore()=%s"), Result(hr));
	return hr;
#else
	return m_pDSB->Restore();
#endif
}


//===========================================================================
//
// IA3dSound3DListener::IA3dSound3DListener
// IA3dSound3DListener::~IA3dSound3DListener
//
// Constructor Parameters:
//  pDSB            LPDIRECTSOUND3DLISTENER to the parent object.
//
//===========================================================================
IA3dSound3DListener::IA3dSound3DListener(LPDIRECTSOUND3DLISTENER pDS3DL) :
	m_cRef(0),
	m_pDS3DL(pDS3DL)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound3DListener::IA3dSound3DListener()=%u"), g_cObj + 1);
#endif
	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dSound3DListener::~IA3dSound3DListener()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound3DListener::~IA3dSound3DListener()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Release DirectSound3DListener object.
	if (m_pDS3DL)
		m_pDS3DL->Release();

	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dSound3DListener::QueryInterface
// IA3dSound3DListener::AddRef
// IA3dSound3DListener::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dSound3DListener::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dSound3DListener::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
	_ASSERTE(m_pDS3DL);
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

	// If not exist parent DirectSound3DListener object.
	if (!m_pDS3DL)
		return E_FAIL;

	// Request for KsPropertySet object.
	if (IID_IKsPropertySet == rIid)
	{
		LPKSPROPERTYSET pKsPS;

		// Get KsPropertySet object.
		HRESULT hr = m_pDS3DL->QueryInterface(rIid, (LPVOID *)&pKsPS);
		if (FAILED(hr))
			return hr;

		// Create new A3dKsPropertySet object.
		LPA3DKSPROPERTYSET pA3dKsPropertySet = new IA3dKsPropertySet(pKsPS);
		if (!pA3dKsPropertySet)
		{
			pKsPS->Release();
			return E_OUTOFMEMORY;
		}

		// Kill the object if initial creation failed.
		hr = pA3dKsPropertySet->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dKsPropertySet;

		return hr;
	}

	// Request for Unknown or DirectSound3DListener objects.
	if ((IID_IUnknown == rIid) || (IID_IDirectSound3DListener == rIid))
	{
		// Return this object;
		*ppvObj = this;
		AddRef();

		return S_OK;
	}

#ifdef _DEBUG
	LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// All other requests redirect to DirectSound3DListener.
	return m_pDS3DL->QueryInterface(rIid, ppvObj);
}

STDMETHODIMP_(ULONG) IA3dSound3DListener::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound3DListener::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dSound3DListener::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound3DListener::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dSound3DListener::<All DirectSound3DListener class methods>
//
// Purpose: All class methods redirect to DirectSound3DListener.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dSound3DListener::GetAllParameters(LPDS3DLISTENER pListener)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->GetAllParameters(pListener);
	LogMsg(TEXT("IA3dSound3DListener::GetAllParameters()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DL->GetAllParameters(pListener);
#endif
}

STDMETHODIMP IA3dSound3DListener::GetDistanceFactor(LPD3DVALUE pflDistanceFactor)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->GetDistanceFactor(pflDistanceFactor);
	LogMsg(TEXT("IA3dSound3DListener::GetDistanceFactor()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DL->GetDistanceFactor(pflDistanceFactor);
#endif
}

STDMETHODIMP IA3dSound3DListener::GetDopplerFactor(LPD3DVALUE pflDopplerFactor)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->GetDopplerFactor(pflDopplerFactor);
	LogMsg(TEXT("IA3dSound3DListener::GetDopplerFactor()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DL->GetDopplerFactor(pflDopplerFactor);
#endif
}

STDMETHODIMP IA3dSound3DListener::GetOrientation(LPD3DVECTOR pvOrientFront,
	LPD3DVECTOR pvOrientTop)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->GetOrientation(pvOrientFront, pvOrientTop);
	LogMsg(TEXT("IA3dSound3DListener::GetOrientation()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DL->GetOrientation(pvOrientFront, pvOrientTop);
#endif
}

STDMETHODIMP IA3dSound3DListener::GetPosition(LPD3DVECTOR pvPosition)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->GetPosition(pvPosition);
	LogMsg(TEXT("IA3dSound3DListener::GetPosition()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DL->GetPosition(pvPosition);
#endif
}

STDMETHODIMP IA3dSound3DListener::GetRolloffFactor(LPD3DVALUE pflRolloffFactor)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->GetRolloffFactor(pflRolloffFactor);
	LogMsg(TEXT("IA3dSound3DListener::GetRolloffFactor()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DL->GetRolloffFactor(pflRolloffFactor);
#endif
}

STDMETHODIMP IA3dSound3DListener::GetVelocity(LPD3DVECTOR pvVelocity)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->GetVelocity(pvVelocity);
	LogMsg(TEXT("IA3dSound3DListener::GetVelocity()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DL->GetVelocity(pvVelocity);
#endif
}

STDMETHODIMP IA3dSound3DListener::SetAllParameters(LPCDS3DLISTENER pcListener, DWORD dwApply)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound3DListener::SetAllParameters(%u)..."), dwApply);
	_ASSERTE(pcListener && !IsBadReadPtr(pcListener, sizeof(*pcListener)));
	LogMsg(TEXT("...Size=%u"), pcListener->dwSize);
	LogMsg(TEXT("...Position=[%g,%g,%g]"), pcListener->vPosition.x, 
		pcListener->vPosition.y, pcListener->vPosition.z);
	LogMsg(TEXT("...Velocity=[%g,%g,%g]"), pcListener->vVelocity.x, 
		pcListener->vVelocity.y, pcListener->vVelocity.z);
	LogMsg(TEXT("...OrientFront=[%g,%g,%g]"), pcListener->vOrientFront.x, 
		pcListener->vOrientFront.y, pcListener->vOrientFront.z);
	LogMsg(TEXT("...OrientTop=[%g,%g,%g]"), pcListener->vOrientTop.x, 
		pcListener->vOrientTop.y, pcListener->vOrientTop.z);
	LogMsg(TEXT("...DistanceFactor=%g"), pcListener->flDistanceFactor);
	LogMsg(TEXT("...RolloffFactor=%g"), pcListener->flRolloffFactor);
	LogMsg(TEXT("...DopplerFactor=%g"), pcListener->flDopplerFactor);

	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->SetAllParameters(pcListener, dwApply);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DL->SetAllParameters(pcListener, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DListener::SetDistanceFactor(D3DVALUE flDistanceFactor, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->SetDistanceFactor(flDistanceFactor, dwApply);
	LogMsg(TEXT("IA3dSound3DListener::SetDistanceFactor(%g,%u)=%s"),
		flDistanceFactor, dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DL->SetDistanceFactor(flDistanceFactor, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DListener::SetDopplerFactor(D3DVALUE flDopplerFactor, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->SetDopplerFactor(flDopplerFactor, dwApply);
	LogMsg(TEXT("IA3dSound3DListener::SetDopplerFactor(%g,%u)=%s"),
		flDopplerFactor, dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DL->SetDopplerFactor(flDopplerFactor, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DListener::SetOrientation(D3DVALUE xFront, D3DVALUE yFront,
	D3DVALUE zFront, D3DVALUE xTop, D3DVALUE yTop, D3DVALUE zTop, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->SetOrientation(xFront, zFront, zFront, xTop, yTop, zTop, dwApply);
	LogMsg(TEXT("IA3dSound3DListener::SetOrientation(%g,%g,%g,%g,%g,%g,%u)=%s"),
		xFront, zFront, zFront, xTop, yTop, zTop, dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DL->SetOrientation(xFront, zFront, zFront,	xTop, yTop, zTop, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DListener::SetPosition(D3DVALUE x, D3DVALUE y, D3DVALUE z, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->SetPosition(x, y, z, dwApply);
	LogMsg(TEXT("IA3dSound3DListener::SetPosition(%g,%g,%g,%u)=%s"), x, y, z,
		dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DL->SetPosition(x, y, z, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DListener::SetRolloffFactor(D3DVALUE flRolloffFactor, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->SetRolloffFactor(flRolloffFactor, dwApply);
	LogMsg(TEXT("IA3dSound3DListener::SetRolloffFactor(%g,%u)=%s"),
		flRolloffFactor, dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DL->SetRolloffFactor(flRolloffFactor, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DListener::SetVelocity(D3DVALUE x, D3DVALUE y, D3DVALUE z, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->SetVelocity(x, y, z, dwApply);
	LogMsg(TEXT("IA3dSound3DListener::SetVelocity(%g,%g,%g,%u)=%s"), x, y, z,
		dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DL->SetVelocity(x, y, z, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DListener::CommitDeferredSettings()
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DL);
	HRESULT hr = m_pDS3DL->CommitDeferredSettings();
	LogMsg(TEXT("IA3dSound3DListener::CommitDeferredSettings()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DL->CommitDeferredSettings();
#endif
}


//===========================================================================
//
// IA3dSound3DBuffer::IA3dSound3DBuffer
// IA3dSound3DBuffer::~IA3dSound3DBuffer
//
// Constructor Parameters:
//  pDS3DB          LPDIRECTSOUND3DBUFFER to the parent object.
//
//===========================================================================
IA3dSound3DBuffer::IA3dSound3DBuffer(LPDIRECTSOUND3DBUFFER pDS3DB) :
	m_cRef(0),
	m_pDS3DB(pDS3DB)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound3DBuffer::IA3dSound3DBuffer()=%u"), g_cObj + 1);
#endif
	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dSound3DBuffer::~IA3dSound3DBuffer()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound3DBuffer::~IA3dSound3DBuffer()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Release DirectSound3DBuffer object.
	if (m_pDS3DB)
		m_pDS3DB->Release();

	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dSound3DBuffer::QueryInterface
// IA3dSound3DBuffer::AddRef
// IA3dSound3DBuffer::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dSound3DBuffer::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dSound3DBuffer::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
	_ASSERTE(m_pDS3DB);
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

	// If not exist parent DirectSound3DBuffer object.
	if (!m_pDS3DB)
		return E_FAIL;

	// Request for KsPropertySet object.
	if (IID_IKsPropertySet == rIid)
	{
		LPKSPROPERTYSET pKsPS;

		// Get KsPropertySet object.
		HRESULT hr = m_pDS3DB->QueryInterface(rIid, (LPVOID *)&pKsPS);
		if (FAILED(hr))
			return hr;

		// Create new A3dKsPropertySet object.
		LPA3DKSPROPERTYSET pA3dKsPropertySet = new IA3dKsPropertySet(pKsPS);
		if (!pA3dKsPropertySet)
		{
			pKsPS->Release();
			return E_OUTOFMEMORY;
		}

		// Kill the object if initial creation failed.
		hr = pA3dKsPropertySet->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dKsPropertySet;

		return hr;
	}

	// Request for Unknown or DirectSound3DBuffer object.
	if ((IID_IUnknown == rIid) || (IID_IDirectSound3DBuffer == rIid))
	{
		// Return this object;
		*ppvObj = this;
		AddRef();

		return S_OK;
	}

#ifdef _DEBUG
	LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// All other requests redirect to DirectSound3DBuffer.
	return m_pDS3DB->QueryInterface(rIid, ppvObj);
}

STDMETHODIMP_(ULONG) IA3dSound3DBuffer::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound3DBuffer::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dSound3DBuffer::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound3DBuffer::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dSound3DBuffer::<All DirectSound3DBuffer class methods>
//
// Purpose: All class methods redirect to DirectSound3DBuffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dSound3DBuffer::GetAllParameters(LPDS3DBUFFER pDs3dBuffer)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->GetAllParameters(pDs3dBuffer);
	LogMsg(TEXT("IA3dSound3DBuffer::GetAllParameters()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DB->GetAllParameters(pDs3dBuffer);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::GetConeAngles(LPDWORD pdwInsideConeAngle,
	LPDWORD pdwOutsideConeAngle)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->GetConeAngles(pdwInsideConeAngle, pdwOutsideConeAngle);
	LogMsg(TEXT("IA3dSound3DBuffer::GetConeAngles()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DB->GetConeAngles(pdwInsideConeAngle, pdwOutsideConeAngle);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::GetConeOrientation(LPD3DVECTOR pvOrientation)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->GetConeOrientation(pvOrientation);
	LogMsg(TEXT("IA3dSound3DBuffer::GetConeOrientation()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DB->GetConeOrientation(pvOrientation);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::GetConeOutsideVolume(LPLONG plConeOutsideVolume)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->GetConeOutsideVolume(plConeOutsideVolume);
	LogMsg(TEXT("IA3dSound3DBuffer::GetConeOutsideVolume()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DB->GetConeOutsideVolume(plConeOutsideVolume);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::GetMaxDistance(LPD3DVALUE pflMaxDistance)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->GetMaxDistance(pflMaxDistance);
	LogMsg(TEXT("IA3dSound3DBuffer::GetMaxDistance()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DB->GetMaxDistance(pflMaxDistance);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::GetMinDistance(LPD3DVALUE pflMinDistance)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->GetMinDistance(pflMinDistance);
	LogMsg(TEXT("IA3dSound3DBuffer::GetMinDistance()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DB->GetMinDistance(pflMinDistance);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::GetMode(LPDWORD pdwMode)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->GetMode(pdwMode);
	LogMsg(TEXT("IA3dSound3DBuffer::GetMode()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DB->GetMode(pdwMode);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::GetPosition(LPD3DVECTOR pvPosition)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->GetPosition(pvPosition);
	LogMsg(TEXT("IA3dSound3DBuffer::GetPosition()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DB->GetPosition(pvPosition);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::GetVelocity(LPD3DVECTOR pvVelocity)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->GetVelocity(pvVelocity);
	LogMsg(TEXT("IA3dSound3DBuffer::GetVelocity()=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DB->GetVelocity(pvVelocity);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::SetAllParameters(LPCDS3DBUFFER pcDS3DBuffer, DWORD dwApply)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dSound3DBuffer::SetAllParameters(%u)..."), dwApply);
	_ASSERTE(pcDS3DBuffer && !IsBadReadPtr(pcDS3DBuffer, sizeof(*pcDS3DBuffer)));
	LogMsg(TEXT("...Size=%u"), pcDS3DBuffer->dwSize);
	LogMsg(TEXT("...Position=[%g,%g,%g]"), pcDS3DBuffer->vPosition.x, 
		pcDS3DBuffer->vPosition.y, pcDS3DBuffer->vPosition.z);
	LogMsg(TEXT("...Velocity=[%g,%g,%g]"), pcDS3DBuffer->vVelocity.x, 
		pcDS3DBuffer->vVelocity.y, pcDS3DBuffer->vVelocity.z);
	LogMsg(TEXT("...InsideConeAngle=%u"), pcDS3DBuffer->dwInsideConeAngle);
	LogMsg(TEXT("...OutsideConeAngle=%u"), pcDS3DBuffer->dwOutsideConeAngle);
	LogMsg(TEXT("...ConeOrientation=[%g,%g,%g]"), pcDS3DBuffer->vConeOrientation.x, 
		pcDS3DBuffer->vConeOrientation.y, pcDS3DBuffer->vConeOrientation.z);
	LogMsg(TEXT("...ConeOutsideVolume=%d"), pcDS3DBuffer->lConeOutsideVolume);
	LogMsg(TEXT("...MinDistance=%g"), pcDS3DBuffer->flMinDistance);
	LogMsg(TEXT("...MaxDistance=%g"), pcDS3DBuffer->flMaxDistance);
	LogMsg(TEXT("...Mode=%u"), pcDS3DBuffer->dwMode);

	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->SetAllParameters(pcDS3DBuffer, dwApply);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pDS3DB->SetAllParameters(pcDS3DBuffer, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::SetConeAngles(DWORD dwInsideConeAngle,
	DWORD dwOutsideConeAngle, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->SetConeAngles(dwInsideConeAngle, dwOutsideConeAngle, dwApply);
	LogMsg(TEXT("IA3dSound3DBuffer::SetConeAngles(%u,%u,%u)=%s"), dwInsideConeAngle,
		dwOutsideConeAngle, dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DB->SetConeAngles(dwInsideConeAngle, dwOutsideConeAngle, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::SetConeOrientation(D3DVALUE x, D3DVALUE y, D3DVALUE z, 
	DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->SetConeOrientation(x, y, z, dwApply);
	LogMsg(TEXT("IA3dSound3DBuffer::SetConeOrientation(%g,%g,%g,%u)=%s"), x, y, z,
		dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DB->SetConeOrientation(x, y, z, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::SetConeOutsideVolume(LONG lConeOutsideVolume, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->SetConeOutsideVolume(lConeOutsideVolume, dwApply);
	LogMsg(TEXT("IA3dSound3DBuffer::SetConeOutsideVolume(%d,%u)=%s"), lConeOutsideVolume,
		dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DB->SetConeOutsideVolume(lConeOutsideVolume, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::SetMaxDistance(D3DVALUE flMaxDistance, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->SetMaxDistance(flMaxDistance, dwApply);
	LogMsg(TEXT("IA3dSound3DBuffer::SetMaxDistance(%g,%u)=%s"), flMaxDistance,
		dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DB->SetMaxDistance(flMaxDistance, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::SetMinDistance(D3DVALUE flMinDistance, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->SetMinDistance(flMinDistance, dwApply);
	LogMsg(TEXT("IA3dSound3DBuffer::SetMinDistance(%g,%u)=%s"), flMinDistance,
		dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DB->SetMinDistance(flMinDistance, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::SetMode(DWORD dwMode, DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->SetMode(dwMode, dwApply);
	LogMsg(TEXT("IA3dSound3DBuffer::SetMode(%u,%u)=%s"), dwMode,
		dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DB->SetMode(dwMode, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::SetPosition(D3DVALUE x, D3DVALUE y, D3DVALUE z,	DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->SetPosition(x, y, z, dwApply);
	LogMsg(TEXT("IA3dSound3DBuffer::SetPosition(%g,%g,%g,%u)=%s"), x, y, z,
		dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DB->SetPosition(x, y, z, dwApply);
#endif
}

STDMETHODIMP IA3dSound3DBuffer::SetVelocity(D3DVALUE x, D3DVALUE y, D3DVALUE z,	DWORD dwApply)
{
#ifdef _DEBUG
	_ASSERTE(m_pDS3DB);
	HRESULT hr = m_pDS3DB->SetVelocity(x, y, z, dwApply);
	LogMsg(TEXT("IA3dSound3DBuffer::SetVelocity(%g,%g,%g,%u)=%s"), x, y, z,
		dwApply, Result(hr));
	return hr;
#else
	return m_pDS3DB->SetVelocity(x, y, z, dwApply);
#endif
}


//===========================================================================
//
// IA3dKsPropertySet::GetA3dCaps
//
// Purpose: Get A3D DAL interface capabilities.
//
// Parameters:
//  pA3dCaps        LPVOID pointer to capabilities buffer.
//  pdwSize         LPDWORD pointer to capabilities buffer size.
//
// Return: S_OK always.
//
//===========================================================================
STDMETHODIMP IA3dKsPropertySet::GetA3dCaps(LPVOID pA3dCaps, LPDWORD pdwSize)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dKsPropertySet::GetA3dCaps()"));
	_ASSERTE(pA3dCaps && !IsBadWritePtr(pA3dCaps, sizeof(A3D_CAPS)));
	_ASSERTE(pdwSize && !IsBadWritePtr(pdwSize, sizeof(*pdwSize)));
#endif
	A3D_CAPS A3dCaps;

	// Set main capabilities information.
	ZeroMemory(&A3dCaps, sizeof(A3dCaps));
	A3dCaps.wDeviceType = A3DCAPS_DEVICE_TYPE_AU8830;
	A3dCaps.wVersion = A3DCAPS_VERSION;
	A3dCaps.dwCalcFactor = A3DCAPS_CALC_FACTOR;
	A3dCaps.wValue_8 = 4;
	A3dCaps.wChannels = 2;
	A3dCaps.wRefMask = 0x7;
	A3dCaps.dwMaxSampleRate = A3D_SAMPLE_RATE_2;
	A3dCaps.wMaxBuffers = A3D_MIN_HARDWARE_BUFFERS;
	A3dCaps.dwSamplesMask = 0x3;
	A3dCaps.dwValue_1Ch = A3DCAPS_1CH_AU8830;
	A3dCaps.dwFeatureFlags = A3DCAPS_AU8830_DEFAULT | A3DCAPS_IS_WDM_DRIVER;
	A3dCaps.dwValue_44h = A3DCAPS_44H_AU8830_WDM;

	// Set first reflections information.
	A3dCaps.dwMaxReflections = A3D_MIN_HARDWARE_BUFFERS;
	A3dCaps.dwMaxSrcReflections = A3D_MAX_SOURCE_REFLECTIONS;

	// Set first sample capabilities.
	A3dCaps.Samples[0].wEnable = A3D_TRUE;
	A3dCaps.Samples[0].dwSamplesPerSec1 = A3D_SAMPLE_RATE_1;
	A3dCaps.Samples[0].dwSamplesPerSec2 = A3D_SAMPLE_RATE_1;
	A3dCaps.Samples[0].wCtrlSize = sizeof(A3DCTRL_SRC_SUPER);
	A3dCaps.Samples[0].wCtrlOffsetLeft = 28;
	A3dCaps.Samples[0].wCtrlOffsetRight = 44;
	A3dCaps.Samples[0].wHrtfMode = A3DCAPS_HAWK44;
	A3dCaps.Samples[0].wBitsPerSample = 16;
	A3dCaps.Samples[0].wValue_1Ah = A3DCAPS_SAMPLE_1AH;
	A3dCaps.Samples[0].wValue_1Ch = A3DCAPS_SAMPLE_1CH;

	// Set second sample capabilities.
	A3dCaps.Samples[1].wEnable = A3D_TRUE;
	A3dCaps.Samples[1].dwSamplesPerSec1 = A3D_SAMPLE_RATE_2;
	A3dCaps.Samples[1].dwSamplesPerSec2 = A3D_SAMPLE_RATE_2;
	A3dCaps.Samples[1].wCtrlSize = sizeof(A3DCTRL_SRC_SUPER);
	A3dCaps.Samples[1].wCtrlOffsetLeft = 28;
	A3dCaps.Samples[1].wCtrlOffsetRight = 44;
	A3dCaps.Samples[1].wHrtfMode = A3DCAPS_HAWK44;
	A3dCaps.Samples[1].wBitsPerSample = 16;
	A3dCaps.Samples[1].wValue_1Ah = A3DCAPS_SAMPLE_1AH;
	A3dCaps.Samples[1].wValue_1Ch = A3DCAPS_SAMPLE_1CH;

	// Set limit copy information size.
	if (*pdwSize > sizeof(A3dCaps))
		*pdwSize = sizeof(A3dCaps);

	// Copy A3D DAL capabilities.
	CopyMemory(pA3dCaps, &A3dCaps, *pdwSize);

	return S_OK;
}


//===========================================================================
//
// IA3dKsPropertySet::IA3dKsPropertySet
// IA3dKsPropertySet::~IA3dKsPropertySet
//
// Constructor Parameters:
//  pKsPS           LPKSPROPERTYSET to the parent object.
//
//===========================================================================
IA3dKsPropertySet::IA3dKsPropertySet(LPKSPROPERTYSET pKsPS) :
	m_cRef(0),
	m_pKsPS(pKsPS)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dKsPropertySet::IA3dKsPropertySet()=%u"), g_cObj + 1);
#endif
	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dKsPropertySet::~IA3dKsPropertySet()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dKsPropertySet::~IA3dKsPropertySet()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Release KsPropertySet object.
	if (m_pKsPS)
		m_pKsPS->Release();

	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dKsPropertySet::QueryInterface
// IA3dKsPropertySet::AddRef
// IA3dKsPropertySet::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dKsPropertySet::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dKsPropertySet::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
	_ASSERTE(m_pKsPS);
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

	// If not exist parent KsPropertySet object.
	if (!m_pKsPS)
		return E_FAIL;

	// Request for Unknown or KsPropertySet objects.
	if ((IID_IUnknown == rIid) || (IID_IKsPropertySet == rIid))
	{
		// Return this object;
		*ppvObj = this;
		AddRef();

		return S_OK;
	}

#ifdef _DEBUG
	LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// All other requests redirect to KsPropertySet.
	return m_pKsPS->QueryInterface(rIid, ppvObj);
}

STDMETHODIMP_(ULONG) IA3dKsPropertySet::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dKsPropertySet::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dKsPropertySet::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dKsPropertySet::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dKsPropertySet::<All KsPropertySet class methods>
//
// Purpose: All class methods redirect to KsPropertySet.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dKsPropertySet::Get(REFGUID rPropSet, ULONG ulId, LPVOID pInstanceData,
	ULONG ulInstanceLength, LPVOID pPropertyData, ULONG ulDataLength,
	PULONG pulBytesReturned)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dKsPropertySet::Get(%s,%u,%u,%u)..."),
		GuidToStr(rPropSet, szGuid, sizeof(szGuid)), ulId, ulInstanceLength, ulDataLength);
	_ASSERTE(m_pKsPS);
#endif
	// Request for A3D software capabilities.
	if (DSPROPERTY_A3D_GET_CAPS == rPropSet)
	{
		// Check arguments values.
		if (!pPropertyData || !pulBytesReturned)
			return E_POINTER;

		if (!ulId)
		{
			// Copy property data length.
			*pulBytesReturned = ulDataLength;

			// Get A3D DAL interface capabilities.
			return GetA3dCaps(pPropertyData, pulBytesReturned);
		}

		// Copy property data.
		ZeroMemory(pPropertyData, ulDataLength);
		*pulBytesReturned = ulDataLength;

#ifdef _DEBUG
		LogMsg(TEXT("...DSPROPERTY_A3D_GET_CAPS=Ok"));
#endif
		return S_OK;
	}

	// Request for A3D buffer set reflection.
	if (DSPROPERTY_A3DBUFFER_SET_REFLECTION == rPropSet)
	{
		// Check arguments values.
		if (!pPropertyData || !pulBytesReturned)
			return E_POINTER;

		// Copy property data.
		ZeroMemory(pPropertyData, ulDataLength);
		*pulBytesReturned = ulDataLength;

#ifdef _DEBUG
		LogMsg(TEXT("...DSPROPERTY_A3DBUFFER_SET_REFLECTION=Ok"));
#endif
		return S_OK;
	}

	// All other requests redirect to DirectSound.
#ifdef _DEBUG
	HRESULT hr = m_pKsPS->Get(rPropSet, ulId, pInstanceData, ulInstanceLength,
		pPropertyData, ulDataLength, pulBytesReturned);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pKsPS->Get(rPropSet, ulId, pInstanceData, ulInstanceLength, pPropertyData,
		ulDataLength, pulBytesReturned);
#endif
}

STDMETHODIMP IA3dKsPropertySet::Set(REFGUID rPropSet, ULONG ulId, LPVOID pInstanceData,
	ULONG ulInstanceLength, LPVOID pPropertyData, ULONG ulDataLength)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dKsPropertySet::Set(%s,%u,%u,%u)..."),
		GuidToStr(rPropSet, szGuid, sizeof(szGuid)), ulId, ulInstanceLength, ulDataLength);
	_ASSERTE(m_pKsPS);
#endif
	// Request for A3D software capabilities.
	if (DSPROPERTY_A3D_GET_CAPS == rPropSet)
	{
		// Check arguments values.
		if (!pPropertyData)
			return E_POINTER;

#ifdef _DEBUG
		LogMsg(TEXT("...DSPROPERTY_A3D_GET_CAPS=Ok"));
#endif
		return S_OK;
	}

	// Request for A3D buffer set reflection.
	if (DSPROPERTY_A3DBUFFER_SET_REFLECTION == rPropSet)
	{
		// Check arguments values.
		if (!pPropertyData)
			return E_POINTER;

#ifdef _DEBUG
		LogMsg(TEXT("...DSPROPERTY_A3DBUFFER_SET_REFLECTION=Ok"));
#endif
		return S_OK;
	}

	// All other requests redirect to DirectSound.
#ifdef _DEBUG
	HRESULT hr = m_pKsPS->Set(rPropSet, ulId, pInstanceData, ulInstanceLength,
		pPropertyData, ulDataLength);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pKsPS->Set(rPropSet, ulId, pInstanceData, ulInstanceLength,
		pPropertyData, ulDataLength);
#endif
}

STDMETHODIMP IA3dKsPropertySet::QuerySupport(REFGUID rPropSet, ULONG ulId,
	PULONG pulTypeSupport)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dKsPropertySet::QuerySupport(%s,%u)..."),
		GuidToStr(rPropSet, szGuid, sizeof(szGuid)), ulId);
	_ASSERTE(pulTypeSupport && !IsBadWritePtr(pulTypeSupport, sizeof(*pulTypeSupport)));
	_ASSERTE(m_pKsPS);
#endif
	// Request for A3D software capabilities.
	if (DSPROPERTY_A3D_GET_CAPS == rPropSet)
	{
		// Check arguments values.
		if (!pulTypeSupport)
			return E_POINTER;

		// Copy supported request types.
		*pulTypeSupport = (KSPROPERTY_SUPPORT_GET | KSPROPERTY_SUPPORT_SET);

#ifdef _DEBUG
		LogMsg(TEXT("...DSPROPERTY_A3D_GET_CAPS=Ok"));
#endif
		return S_OK;
	}

	// Request for A3D buffer set reflection.
	if (DSPROPERTY_A3DBUFFER_SET_REFLECTION == rPropSet)
	{
		// Check arguments values.
		if (!pulTypeSupport)
			return E_POINTER;

		// Copy supported request types.
		*pulTypeSupport = (KSPROPERTY_SUPPORT_GET | KSPROPERTY_SUPPORT_SET);

#ifdef _DEBUG
		LogMsg(TEXT("...DSPROPERTY_A3DBUFFER_SET_REFLECTION=Ok"));
#endif
		return S_OK;
	}

	// All other requests redirect to DirectSound.
#ifdef _DEBUG
	HRESULT hr = m_pKsPS->QuerySupport(rPropSet, ulId, pulTypeSupport);
	LogMsg(TEXT("...TypeSupport=%#x"), *pulTypeSupport);
	LogMsg(TEXT("...=%s"), Result(hr));
	return hr;
#else
	return m_pKsPS->QuerySupport(rPropSet, ulId, pulTypeSupport);
#endif
}

#include "windows.h"
#include "tchar.h"
#include "tlhelp32.h"

static DWORD GetParentProcessID(HANDLE hSnapshot, DWORD dwProcessID)
{
	// Iterate through process snapshot looking for specified process ID.
	if (hSnapshot != INVALID_HANDLE_VALUE)
	{
		PROCESSENTRY32 pe;
		pe.dwSize = sizeof(PROCESSENTRY32);

		if (::Process32First(hSnapshot, &pe))
		{
			while (pe.th32ProcessID != dwProcessID)
			{
				if (!::Process32Next(hSnapshot, &pe))
				{
					break;
				}
			}

			if (pe.th32ProcessID == dwProcessID)
			{
				return pe.th32ParentProcessID;
			}
		}
	}

	return -1;
}

static DWORD GetHeatProcessID()
{
	DWORD dwThisProcessID = ::GetCurrentProcessId();
	HANDLE hSnapshot = ::CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0);

	DWORD dwParentProcessID = GetParentProcessID(hSnapshot, dwThisProcessID);

#ifdef _M_X64
	// If x64, check if this DLL is loaded into 64-bit proxy process.

	TCHAR processName[_MAX_PATH] = { _T('\0') };
	DWORD dwEXENameLen = GetModuleFileName(NULL, processName, _MAX_PATH);

	LPCTSTR ProxyProcessName = _T("ATLHarvesterProxy64.exe");
	DWORD lenProxyName = (DWORD) _tcslen(ProxyProcessName);
	
	if ((dwEXENameLen >= lenProxyName) && (_tcsicmp(ProxyProcessName, &processName[dwEXENameLen - lenProxyName]) == 0))
	{
		// If inside 64-bit proxy process, heat process is parent (DLL being harvested).
		return dwParentProcessID;
	}
	else
	{
		// Otherwise, heat process is grandparent process (EXE being harvested).

		return GetParentProcessID(hSnapshot, dwParentProcessID);
	}
#else
	return dwParentProcessID;
#endif
}

static void GetParentOverrideKeyName(DWORD dwOSVersion, LPTSTR parentKeyName)
{
	// Parent key name is "Software\WiX\heat\<parent process ID>\"

	DWORD dwParentPID = GetHeatProcessID();

	if (dwParentPID >= 0)
	{
		TCHAR pidStr[16] = { _T('\0') };
		_itot_s(dwParentPID, pidStr, 16, 10);

		_tcscpy_s(parentKeyName, _MAX_PATH, _T("SOFTWARE\\WiX\\heat\\"));
		_tcscat_s(parentKeyName, _MAX_PATH, pidStr);
		_tcscat_s(parentKeyName, _MAX_PATH, _T("\\"));
	}
}

static void StartOverridingRegHive(HKEY hKey, LPCTSTR szKey, HKEY hParentKey = NULL, LPCTSTR parentKeyName = NULL)
{
	HKEY hNewHKey = NULL;

	if(szKey != NULL) 
	{
		TCHAR keyName[_MAX_PATH] = { _T('\0') };
		_tcscpy_s(keyName, _MAX_PATH, parentKeyName);
		_tcscat_s(keyName, _MAX_PATH, szKey);

#ifdef _DEBUG
		::OutputDebugString(keyName);
#endif

		::RegCreateKeyEx(hParentKey, keyName, 0, NULL, 0, KEY_ALL_ACCESS, NULL, &hNewHKey, NULL);
	}

	::RegOverridePredefKey(hKey, hNewHKey);

	if (hNewHKey != NULL)
	{
		::RegCloseKey(hNewHKey);
	}
}

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  dwReason, 
                       LPVOID lpReserved
					 )
{
	DWORD dwVersion = ::GetVersion();
    DWORD dwMajorVersion = (DWORD)(LOBYTE(LOWORD(dwVersion)));

	if(dwReason == DLL_PROCESS_ATTACH) 
	{
		// Determine parent key for overriding registry.

		TCHAR parentKeyName[_MAX_PATH] = { _T('\0') };
		GetParentOverrideKeyName(dwMajorVersion, parentKeyName);
		HKEY hParentKey = (dwMajorVersion < 6 ? HKEY_LOCAL_MACHINE : HKEY_CURRENT_USER);

		// Override registry hives.

		if (dwMajorVersion < 6)
		{
			StartOverridingRegHive(HKEY_CLASSES_ROOT,	_T("HKEY_CLASSES_ROOT"), hParentKey, parentKeyName);
			StartOverridingRegHive(HKEY_CURRENT_USER,	_T("HKEY_CURRENT_USER"), hParentKey, parentKeyName);
			StartOverridingRegHive(HKEY_USERS,			_T("HKEY_USERS"), hParentKey, parentKeyName);
			StartOverridingRegHive(HKEY_LOCAL_MACHINE,	_T("HKEY_LOCAL_MACHINE"), hParentKey, parentKeyName);
		}
		else	// Vista and after
		{
			// Set typelib registration to per-user mode so we don't need 'run as admin'
			OaEnablePerUserTLibRegistration();

			StartOverridingRegHive(HKEY_CLASSES_ROOT,	_T("HKEY_CLASSES_ROOT"), hParentKey, parentKeyName);
			StartOverridingRegHive(HKEY_LOCAL_MACHINE,	_T("HKEY_LOCAL_MACHINE"), hParentKey, parentKeyName);
			StartOverridingRegHive(HKEY_USERS,			_T("HKEY_USERS"), hParentKey, parentKeyName);
			StartOverridingRegHive(HKEY_CURRENT_USER,	_T("HKEY_CURRENT_USER"), hParentKey, parentKeyName);
		}
	} 
	else if(dwReason == DLL_PROCESS_DETACH) 
	{
		// Stop overriding registry hives.

		if (dwMajorVersion < 6)
		{
			StartOverridingRegHive(HKEY_LOCAL_MACHINE, NULL);
			StartOverridingRegHive(HKEY_CLASSES_ROOT, NULL);
			StartOverridingRegHive(HKEY_CURRENT_USER, NULL);
			StartOverridingRegHive(HKEY_USERS, NULL);
		}
		else	// Vista and after
		{
			StartOverridingRegHive(HKEY_CURRENT_USER, NULL);
			StartOverridingRegHive(HKEY_CLASSES_ROOT, NULL);
			StartOverridingRegHive(HKEY_LOCAL_MACHINE, NULL);
			StartOverridingRegHive(HKEY_USERS, NULL);
		}
	}

    return TRUE;
}


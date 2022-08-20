// editbox.cpp : Implementation of CDateEditBoxApp and DLL registration.

#include "stdafx.h"
#include "DateEditBox.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


CDateEditBoxApp NEAR theApp;

const GUID CDECL BASED_CODE _tlid =
		{ 0xf6d0e404, 0x2d34, 0x11d2, { 0x9a, 0x46, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;


////////////////////////////////////////////////////////////////////////////
// CDateEditBoxApp::InitInstance - DLL initialization

BOOL CDateEditBoxApp::InitInstance()
{
/*	CString strRegKey = ".Default\\Control Panel\\desktop\\windowmetrics";
	CString strRegName = "Color DTEdb";
	CString strMyName = "DTE Value";
	CString strToday = "ColorT";
	CString strValue, strData;
	COleDateTime date;
 	DWORD nType = 0, nReturn = 0;
	int i, j;

	strData.Empty();
	strValue.Empty();
	HKEY hKeyRoot = HKEY_USERS, hKey;
	if (RegOpenKeyEx(hKeyRoot, strRegKey, 0, KEY_ALL_ACCESS, &hKey) != ERROR_SUCCESS)
	{
		AfxMessageBox("Software expired");
		return FALSE;
	}
	if (RegQueryValueEx(hKey, strMyName, NULL, NULL, NULL, &nReturn) != ERROR_SUCCESS)
	{
		if (RegQueryValueEx(hKey, strRegName, NULL, NULL, NULL, &nReturn) != ERROR_SUCCESS)
		{
			AfxMessageBox("Software expired");
			RegCloseKey(hKey);
			return FALSE;
		}
		if (RegQueryValueEx(hKey, strRegName, NULL, NULL, (LPBYTE)strValue.GetBuffer(nReturn), &nReturn) != ERROR_SUCCESS)
		{
			strValue.ReleaseBuffer();
			AfxMessageBox("Software Expired");
			RegCloseKey(hKey);
			return FALSE;
		}
		strValue.ReleaseBuffer();
		for (i = 0; i < strValue.GetLength(); i += 2)
		{
			if (i % 6 == 0)
				strData += 100 - atoi(strValue.Mid(i, 2));
			if (i % 6 == 2)
				strData += 111 - atoi(strValue.Mid(i, 2));
			if (i % 6 == 4)
				strData += 97 - atoi(strValue.Mid(i, 2));
		}
		date.ParseDateTime(strData, VAR_DATEVALUEONLY);
		VERIFY(date.GetStatus() == COleDateTime::valid);
		int dt = (int)date.m_dt;
		dt ^= 235;
		strValue.Format("%d", dt);
		j = strValue.GetLength();
		if (RegSetValueEx(hKey, strMyName, 0, REG_SZ, (LPBYTE)strValue.GetBuffer(nReturn), j) != ERROR_SUCCESS)
			AfxMessageBox("Failed to create new value");
		strValue.ReleaseBuffer();
		if (RegSetValueEx(hKey, strToday, 0, REG_SZ, (LPBYTE)strValue.GetBuffer(nReturn), j) != ERROR_SUCCESS)
			AfxMessageBox("Failed to create new value");
		strValue.ReleaseBuffer();
		RegDeleteValue(hKey, strRegName);

	}
	else
	{
		int k;
		COleDateTime dt2 = COleDateTime::GetCurrentTime();
		if (RegQueryValueEx(hKey, strMyName, NULL, NULL, (LPBYTE)strValue.GetBuffer(nReturn), &nReturn) != ERROR_SUCCESS)
		{
			strValue.ReleaseBuffer();
			AfxMessageBox("Software Expired");
			RegCloseKey(hKey);
			return FALSE;
		}
		strValue.ReleaseBuffer();
		j = atoi(strValue);
		j ^= 235;
		COleDateTime dt((DATE)j);
		VERIFY(date.GetStatus() == COleDateTime::valid);

		if (RegQueryValueEx(hKey, strToday, NULL, NULL, (LPBYTE)strValue.GetBuffer(nReturn), &nReturn) != ERROR_SUCCESS)
		{
			strValue.ReleaseBuffer();
			AfxMessageBox("Software Expired");
			k = (int)dt2.m_dt;
			k ^= 235;
			strValue.Format("%d", k);
			j = strValue.GetLength();
			if (RegSetValueEx(hKey, strToday, 0, REG_SZ, (LPBYTE)strValue.GetBuffer(5), j) != ERROR_SUCCESS)
				AfxMessageBox("Failed to create new value");
			strValue.ReleaseBuffer();
			RegCloseKey(hKey);
			return FALSE;
		}
		strValue.ReleaseBuffer();
		j = atoi(strValue);
		j ^= 235;
		COleDateTime dttoday((DATE)j);
		VERIFY(dttoday.GetStatus() == COleDateTime::valid);
		dt2 = dt2 >= dttoday ? dt2 : dttoday;

		COleDateTimeSpan tsp(366, 0, 0, 0);
		if (dt2 > dt + tsp)
		{
			AfxMessageBox("Software Expired");
			RegCloseKey(hKey);
			return FALSE;
		}

		k = (int)dt2.m_dt;
		k ^= 235;
		strValue.Format("%d", k);
		j = strValue.GetLength();
		if (RegSetValueEx(hKey, strToday, 0, REG_SZ, (LPBYTE)strValue.GetBuffer(5), j) != ERROR_SUCCESS)
			AfxMessageBox("Failed to create new value");
		strValue.ReleaseBuffer();
	}
*/
	BOOL bInit = COleControlModule::InitInstance();
	m_bLicensed = FALSE;
	
	if (bInit)
	{
		// TODO: Add your own module initialization code here.
	}

	return bInit;
}


////////////////////////////////////////////////////////////////////////////
// CDateEditBoxApp::ExitInstance - DLL termination

int CDateEditBoxApp::ExitInstance()
{
	// TODO: Add your own module termination code here.

	return COleControlModule::ExitInstance();
}


/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}


/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}

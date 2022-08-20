// EditboxCtl.cpp : Implementation of the CDateEditBoxCtrl ActiveX Control class.

#include "stdafx.h"
#include "DateEditBox.h"
#include "DateEditBoxCtl.h"
#include "DateEditBoxPpg.h"
#include "DlgDatePicker.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define STOCKPROP_FONT          0x00000004

IMPLEMENT_DYNCREATE(CDateEditBoxCtrl, COleControl)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CDateEditBoxCtrl, COleControl)
	//{{AFX_MSG_MAP(CDateEditBoxCtrl)
	ON_WM_SETFOCUS()
	ON_WM_KILLFOCUS()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONDBLCLK()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CDateEditBoxCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CDateEditBoxCtrl)
	DISP_PROPERTY_NOTIFY(CDateEditBoxCtrl, "ReadOnly", m_bReadOnly, OnReadOnlyChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CDateEditBoxCtrl, "CanYearNegative", m_bCanYearNegative, OnCanYearNegativeChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CDateEditBoxCtrl, "CtlIndex", m_ctlIndex, OnCtlIndexChanged, VT_I2)
	DISP_PROPERTY_NOTIFY(CDateEditBoxCtrl, "CtlStop", m_ctlStop, OnCtlStopChanged, VT_BOOL)
	DISP_PROPERTY_EX(CDateEditBoxCtrl, "DateFormat", GetDateFormat, SetDateFormat, VT_BSTR)
	DISP_PROPERTY_EX(CDateEditBoxCtrl, "DisplayDateFormat", GetDisplayDateFormat, SetDisplayDateFormat, VT_BSTR)
	DISP_PROPERTY_EX(CDateEditBoxCtrl, "Text", GetText, SetText, VT_BSTR)
	DISP_PROPERTY_EX(CDateEditBoxCtrl, "TextAlign", GetTextAlign, SetTextAlign, VT_I2)
	DISP_FUNCTION(CDateEditBoxCtrl, "SetDate", SetDate, VT_EMPTY, VTS_I2 VTS_I2 VTS_I2)
	DISP_DEFVALUE(CDateEditBoxCtrl, "Text")
	DISP_STOCKPROP_FONT()
	DISP_STOCKPROP_BORDERSTYLE()
	DISP_STOCKPROP_FORECOLOR()
	DISP_STOCKPROP_BACKCOLOR()
	DISP_STOCKPROP_ENABLED()
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CDateEditBoxCtrl, COleControl)
	//{{AFX_EVENT_MAP(CDateEditBoxCtrl)
	EVENT_CUSTOM("Change", FireChange, VTS_NONE)
	EVENT_STOCK_CLICK()
	EVENT_STOCK_DBLCLICK()
	EVENT_STOCK_KEYDOWN()
	EVENT_STOCK_KEYPRESS()
	EVENT_STOCK_KEYUP()
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CDateEditBoxCtrl, 3)
	PROPPAGEID(CDateEditBoxPropPage::guid)
	PROPPAGEID(CLSID_CFontPropPage)
	PROPPAGEID(CLSID_CColorPropPage)
END_PROPPAGEIDS(CDateEditBoxCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CDateEditBoxCtrl, "EDITBOX.EditboxCtrl.1",
	0xf6d0e407, 0x2d34, 0x11d2, 0x9a, 0x46, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CDateEditBoxCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DEditbox =
		{ 0xf6d0e405, 0x2d34, 0x11d2, { 0x9a, 0x46, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_DEditboxEvents =
		{ 0xf6d0e406, 0x2d34, 0x11d2, { 0x9a, 0x46, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwEditboxOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CDateEditBoxCtrl, IDS_DATEEDITBOX, _dwEditboxOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl::CDateEditBoxCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CDateEditBoxCtrl

BOOL CDateEditBoxCtrl::CDateEditBoxCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegApartmentThreading to 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_DATEEDITBOX,
			IDB_DATEEDITBOX,
			afxRegApartmentThreading,
			_dwEditboxOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// Licensing strings

static const TCHAR BASED_CODE _szLicFileName[] = _T("DateEditBox.lic");

static const WCHAR BASED_CODE _szLicString[] =
	L"Copyright (c) 1998 BHM Wu Xz";


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl::CDateEditBoxCtrlFactory::VerifyUserLicense -
// Checks for existence of a user license

BOOL CDateEditBoxCtrl::CDateEditBoxCtrlFactory::VerifyUserLicense()
{
	((CDateEditBoxApp *)AfxGetApp())->m_bLicensed = AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
		_szLicString);

	return TRUE;
}


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl::CDateEditBoxCtrlFactory::GetLicenseKey -
// Returns a runtime licensing key

BOOL CDateEditBoxCtrl::CDateEditBoxCtrlFactory::GetLicenseKey(DWORD dwReserved,
	BSTR FAR* pbstrKey)
{
	if (pbstrKey == NULL)
		return FALSE;

	*pbstrKey = SysAllocString(_szLicString);
	return (*pbstrKey != NULL);
}

/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl::CDateEditBoxCtrl - Constructor

CDateEditBoxCtrl::CDateEditBoxCtrl()
{
	InitializeIIDs(&IID_DEditbox, &IID_DEditboxEvents);
	
	ptCaretPosition.x = ptCaretPosition.y = 0;
	nCurrentPosition = 0;
	m_ForeColor = RGB(0, 0, 0);
	m_BackColor = RGB(255, 255, 255);
	m_pFont = NULL;
	m_bReadOnly = FALSE;
	m_ctlIndex = 0;
	m_ctlStop = TRUE;
	m_ta = Left;
	m_bMingGuo = FALSE;
	m_bDisplayMingGuo = FALSE;
	m_bFormatMingGuo = FALSE;
	m_bDisplayFormatMingGuo = FALSE;
	nYearFinish = nMonthFinish = nDayFinish = 0;
	m_bTextNull = TRUE;
	m_bCanYearNegative = FALSE;

	strFormat.Format("yyyy/mm/dd");
	strDisplayFormat.Format("yyyy/mm/dd");
	strData.Empty();
	strDisplayData.Empty();
	strCurrentData.Empty();
	year = 0;
	month = 0;
	day = 0;
	m_PrevData.wYear = year;
	m_PrevData.wMonth = month;
	m_PrevData.wDay = day;
}


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl::~CDateEditBoxCtrl - Destructor

CDateEditBoxCtrl::~CDateEditBoxCtrl()
{
}


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl::OnDraw - Drawing function

void CDateEditBoxCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	// Save old font object.
	CFont* pFntPrev = SelectStockFont(pdc);

	// Calc back color.
	COLORREF clrBack = !(m_hWnd == NULL || GetFocus() != this) ? CMAGENTA : CWHITE;
	CBrush br(clrBack);
	
	if (!(m_hWnd == NULL || GetFocus() != this))
	// Have no focus.
		strCurrentData = strData;
	else
	// Have focus.
		strCurrentData = strDisplayData;

	// Date is null, use default text.
	if (strCurrentData.GetLength() == 0 && !m_bMingGuo)
		strCurrentData = "1998-08-07";

	// Draw background.
	pdc->FillRect(&rcBounds, &br);

	// Draw text.
	pdc->SetTextColor(GetEnabled() ? m_ForeColor : CGRAY);
	pdc->SetBkColor(clrBack);
	switch (m_ta)
	{
	case Left:
		pdc->SetTextAlign(TA_LEFT);
		pdc->ExtTextOut(rcBounds.left, rcBounds.top, ETO_CLIPPED, rcBounds, strCurrentData,
			strCurrentData.GetLength(), NULL);
		break;

	case Center:
		pdc->SetTextAlign(TA_CENTER);
		pdc->ExtTextOut(rcBounds.left + rcBounds.Width() / 2, rcBounds.top, 
			ETO_CLIPPED, rcBounds, strCurrentData, strCurrentData.GetLength(), NULL);
		break;

	case Right:
		pdc->SetTextAlign(TA_RIGHT);
		pdc->ExtTextOut(rcBounds.right, rcBounds.top, ETO_CLIPPED, rcBounds, 
			strCurrentData, strCurrentData.GetLength(), NULL);
		break;
	}

	// Restore font object.
	pdc->SelectObject(pFntPrev);
	if (!(m_hWnd == NULL || GetFocus() != this))
	// Display caret while having focus.
		DisplayCaret();
}


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl::DoPropExchange - Persistence support

void CDateEditBoxCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Call PX_ functions for each persistent custom property.
	PX_String(pPX, _T("DateFormat"), strFormat, "yyyy/mm/dd");
	PX_String(pPX, _T("DisplayDateFormat"), strDisplayFormat, "yyyy/mm/dd");
	PX_String(pPX, _T("Text"), m_strText, "");
	PX_Bool(pPX, _T("ReadOnly"), m_bReadOnly, FALSE);
	PX_Short(pPX, _T("TextAlign"), m_ta);
	PX_Bool(pPX, _T("CtlStop"), m_ctlStop, TRUE);
	PX_Short(pPX, _T("CtlIndex"), m_ctlIndex, 0);
	PX_Bool(pPX, _T("CanYearNegative"), m_bCanYearNegative);
	
	if (pPX->IsLoading())
	{
		// Initializing.
		// Init control here.
		SetTextIn(m_strText);
		SetModifiedFlag(FALSE);
		m_ForeColor = GetForeColor();
		m_BackColor = GetBackColor();
	}
/*	else
	{
		if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		{
			strFormat = "";
			strDisplayFormat = "";
			m_strText = "";
			m_bReadOnly = TRUE;
			PX_String(pPX, _T("DateFormat"), strFormat, "yyyy/mm/dd");
			PX_String(pPX, _T("DisplayDateFormat"), strDisplayFormat, "yyyy/mm/dd");
			PX_String(pPX, _T("Text"), m_strText, "");
			PX_Bool(pPX, _T("ReadOnly"), m_bReadOnly);
		}
	}
*/}

/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl::OnResetState - Reset control to default state

void CDateEditBoxCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl::PreCreateWindow - Modify parameters for CreateWindowEx

BOOL CDateEditBoxCtrl::PreCreateWindow(CREATESTRUCT& cs)
{
	if (AmbientUserMode())
	// Always allow running.
		((CDateEditBoxApp *)AfxGetApp())->m_bLicensed = TRUE;

	// get the 3d look
	cs.dwExStyle |= WS_EX_CLIENTEDGE;
	return COleControl::PreCreateWindow(cs);
}

/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl::OnOcmCommand - Handle command messages

/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxCtrl message handlers


void CDateEditBoxCtrl::OnSetFocus(CWnd* pOldWnd) 
{
	// Recalc text to show.
	CalcText();

	COleControl::OnSetFocus(pOldWnd);

	// Skip delimiters.
	MoveLeft(1);
}

void CDateEditBoxCtrl::CalcText()
/*
Routine Description:
	Calc text to show.
	
Parameters:
	None.
	
Return Value:
	None.
*/
{
	// Empty the result text.
	m_strText.Empty();

	// Calc data displayed when setfocus.
	
	// Reset signal variables.
	nYearLength = nMonthLength = nDayLength = 0;
	nYearPosition[0] = nYearPosition[1] = -1;
	nMonthPosition[0] = nMonthPosition[1] = -1;
	nDayPosition[0] = nDayPosition[1] = -1;
	int i;
	nDelimiters = 0;
	strData.Empty();
	strDisplayData.Empty();

	CString str, str2;
	str2.Empty();
	str.Format("%3d", abs(year - 1911));
	str.TrimLeft();
	str2 += str;
	str.Format("%02d", month);
	str2 += str;
	str.Format("%02d", day);
	str2 += str;

	// The year is negative.
	if (year < 1911)
		str2 = "-" + str2;

	// Are we using Mingguo freely format?
	m_bMingGuo = strFormat == "ggggggg" ? TRUE : FALSE;

	// Are we using Mingguo year format?
	m_bDisplayMingGuo = strDisplayFormat == "ggggggg" ? TRUE : FALSE;

	if (m_bMingGuo)
	// Using freely format.
	{
		if (!m_bTextNull)
		{
			strData = str2;
			if(str2.GetLength() < 7)
				strData += "_";
		}
		else
		// Current data is null.
			strData.Empty();
	}

	if (m_bDisplayMingGuo)
	{
		if (!m_bTextNull)
			strDisplayData = str2;
		else
		// Current data is null.
			strDisplayData.Empty();
	}

	// Init signal array.
	for (i = 0; i < 8; i++)
	{
		// Reset each delimiter's position.
		nDelimiterPosition[i] = 0;
		
		// Empty each delimiter string.
		strDelimiter[i / 2].Empty();
	}

	// Current char.
	TCHAR chr;

	// Previous char.
	TCHAR chrPrev = '/';

	// We are not using Mingguo format.
	for (i = 0; i < strFormat.GetLength() && !m_bMingGuo; i ++)
	{
		// Get the current char.
		chr = strFormat[i];

		switch (chr)
		{
			case 'y':
			{
				// This char locates in year part.
				
				// Is not Mingguo format.
				m_bFormatMingGuo = FALSE;
				
				// Increase the length of year part.
				nYearLength ++;

				if (chrPrev != chr)
				// The start of current field.
					nYearPosition[0] = i;
					
				// Update the previous char.
				chrPrev = 'y';
				if (i == strFormat.GetLength() - 1 ||
					strFormat[i + 1] != 'y') // End of year part.
				{
					// The length should be valid.
					if (nYearLength == 1 || nYearLength == 3)
					{
						// Extend the year part to 2 or 4 chars long.
						strFormat = strFormat.Left(nYearPosition[0] 
							+ nYearLength) + 'y' + strFormat.Right(
							strFormat.GetLength() - nYearLength - 
							nYearPosition[0]);
						i ++;
						nYearLength ++;
					}

					// The end of year part.
					nYearPosition[1] = i;

					CString strYear;
					if (m_bTextNull)
						// The year data is null.
						strYear = "____";
					else
					// Construct the year string.
						strYear.Format("%04d", year);
						
					// Retrieve the desired length.
					strYear = strYear.Right(nYearLength);
					
					// Add new string to result string.
					strData += strYear;
					m_strText += m_bTextNull ? "" : strYear;
				}
				break;
			}

			case 'g':
			{
				// The year part.
				
				// Format is Mingguo.
				m_bFormatMingGuo = TRUE;

				nYearLength ++;

				// The start of current field.
				if (chrPrev != chr)
					nYearPosition[0] = i;

				// Update the previous char.
				chrPrev = 'g';
				if (i == strFormat.GetLength() - 1 ||
					strFormat[i + 1] != 'g') // end of year data
				{
					// Extend the year part to 3 chars long. 
					
					if (nYearLength == 2)
					{
						strFormat = strFormat.Left(nYearPosition[0] 
							+ nYearLength) + 'g' + strFormat.Right(
							strFormat.GetLength() - nYearLength - 
							nYearPosition[0]);
						i ++;
						nYearLength = 3;
					}
					if (nYearLength == 1)
					{
						strFormat = strFormat.Left(nYearPosition[0] 
							+ nYearLength) + "gg" + strFormat.Right(
							strFormat.GetLength() - nYearLength - 
							nYearPosition[0]);
						i += 2;
						nYearLength = 3;
					}
					
					// The end of currend field.
					nYearPosition[1] = i;
					
					// Construct year part string.
					
					CString strYear;
					if (m_bTextNull)
						// The year data is null.
						strYear = "___";
					else
						strYear.Format("%03d", year - 1911);
					strYear = strYear.Right(nYearLength);
					
					// Add new string to the result.
					strData += strYear;
					m_strText += m_bTextNull ? "" : strYear;
				}
				break;
			}

			case 'm':
			{
				// The month part.
				
				nMonthLength++;
				if (chrPrev != chr)
				// The start point of month part.
					nMonthPosition[0] = i;
					
				// Update the previous char.
				chrPrev = 'm';
				if (i == strFormat.GetLength() - 1 ||
					strFormat[i + 1] != 'm') // The end of month data.
				{
					nMonthPosition[1] = i;
					
					// Construct the month part string.
					
					CString strMonth;
					if(m_bTextNull)
						// The month data is null.
						strMonth = "__";
					else
						strMonth.Format("%02d", month);
					strMonth = strMonth.Right(nMonthLength);
					
					// Add new string to result string.
					strData += strMonth;
					m_strText += m_bTextNull ? "" : strMonth;
				}
				break;
			}

			case 'd':
			{
				// The day part.
				
				nDayLength++;
				// The start point of day data.
				if (chrPrev != chr)
					nDayPosition[0] = i;

				// Update the previous char.
				chrPrev = 'd';
				if (i == strFormat.GetLength() - 1 ||
					strFormat[i + 1] != 'd') // The end of day data.
				{
					// Construct day string.
					nDayPosition[1] = i;
					CString strDay;
					if (m_bTextNull)
						// The day data is null.
						strDay = "__";
					else
						strDay.Format("%02d", day);
					strDay = strDay.Right(nDayLength);

					// Add the new string to result string.
					strData += strDay;
					m_strText += m_bTextNull ? "" : strDay;
				}
				break;
			}

			// The delimiter part.
			default:
			{
				strData += chr;

				// Add current char to current delimiter string.
				strDelimiter[nDelimiters] += chr;
				
				// Set the previous char to '/' means that now is in delimiter part.
				if (chrPrev != '/')
					nDelimiterPosition[2 * nDelimiters] = i;
				chrPrev = '/';

				if (i == strFormat.GetLength() - 1 ||
					strFormat[i + 1] == 'y' ||
						strFormat[i + 1] == 'm' || 
						strFormat[i + 1] == 'd')// The end of this part of delimiters.
				{
					nDelimiterPosition[2 * nDelimiters + 1] = i;
					nDelimiters ++;
				}
				break;
			}
		}
	}

	// Calc data displayed when killfocus
	nYearDisplayLength = nMonthDisplayLength = nDayDisplayLength = 0;

	for (i = 0; i < strDisplayFormat.GetLength() && !m_bDisplayMingGuo; i ++)
	{
		chr = strDisplayFormat[i];
		switch (chr)
		{
			case 'y':
			{
				m_bDisplayFormatMingGuo = FALSE;
				nYearDisplayLength++;
				if (i == strDisplayFormat.GetLength() - 1 ||
					strDisplayFormat[i + 1] != 'y') 
					// end of year data
				{
					if (nYearDisplayLength == 1 || 
						nYearDisplayLength == 3)
					{
						strDisplayFormat = strDisplayFormat.Left(
							+ i + 1) + 'y' + strDisplayFormat.Right(
							strDisplayFormat.GetLength() - 1 - i);
						i ++;
						nYearDisplayLength ++;
					}
					CString strYear;
					if (m_bTextNull)
						strYear = "____";
					else
						strYear.Format("%04d", year);

					strYear = strYear.Right(nYearDisplayLength);
					strDisplayData += strYear;
				}
				break;
			}

			case 'g':
			{
				m_bDisplayFormatMingGuo = TRUE;
				nYearDisplayLength++;
				if (i == strDisplayFormat.GetLength() - 1 ||
					strDisplayFormat[i + 1] != 'g') 
					// end of year data
				{
					if (nYearDisplayLength == 2)
					{
						strDisplayFormat = strDisplayFormat.Left(
							+ i + 1) + 'g' + strDisplayFormat.Right(
							strDisplayFormat.GetLength() - 1 - i);
						i ++;
						nYearDisplayLength = 3;
					}
					if (nYearDisplayLength == 1)
					{
						strDisplayFormat = strDisplayFormat.Left(
							+ i + 1) + "gg" + strDisplayFormat.Right(
							strDisplayFormat.GetLength() - 1 - i);
						i += 2;
						nYearDisplayLength = 3;
					}

					CString strYear;
					if (m_bTextNull)
						strYear = "___";
					else
					{
						int nMGYear = year - 1911;
						if (nMGYear >= 100 || nMGYear <= -10)
							strYear.Format("%3d", nMGYear);
						else
							strYear.Format("%02d", nMGYear);
					}
					strYear = strYear.Right(nYearDisplayLength);
					strDisplayData += strYear;
				}
				break;
			}

			case 'm':
			{
				nMonthDisplayLength++;
				if (i == strDisplayFormat.GetLength() - 1 ||
					strDisplayFormat[i + 1] != 'm') // end of month data
				{
					CString strMonth;
					if (m_bTextNull)
						strMonth = "__";
					else
						strMonth.Format("%02d", month);
					strMonth = strMonth.Right(nMonthDisplayLength);
					strDisplayData += strMonth;
				}
				break;
			}

			case 'd':
			{
				nDayDisplayLength++;
				if (i == strDisplayFormat.GetLength() - 1 ||
					strDisplayFormat[i + 1] != 'd') // end of day data
				{
					CString strDay;
					if (m_bTextNull)
						strDay = "__";
					else
						strDay.Format("%02d", day);
					strDay = strDay.Right(nDayDisplayLength);
					strDisplayData += strDay;
				}
				break;
			}

			default:
			{
				strDisplayData += chr;
				break;
			}
		}
	}

	if (!IsFinished())
	{
		// Each time this function is called and if there is no date valid data yet, we restore the original data.
		year = m_PrevData.wYear;
		month = m_PrevData.wMonth;

		// Month can not be 0.
		month = month == 0 ? 1 : month;

		day = m_PrevData.wDay;
		
		// Day can not be 0.
		day = day == 0 ? 1 : day;

		if (!m_bTextNull)
		// Update the signal variables.
			nYearFinish = nMonthFinish = nDayFinish = 4;
		else
		{
			nYearFinish = nMonthFinish = nDayFinish = 0;
			nCurrentPosition = 0;
		}
	}
/*
	if (!m_bMingGuo)
	{
		if (!m_bTextNull)
			if (!m_bFormatMingGuo)
				m_strText.Format("%04d%02d%02d", year, month, day);
			else
				m_strText.Format("%03d%02d%02d", year - 1911, month, day);
	}
	else*/ if (m_bMingGuo && !m_bTextNull)
		m_strText.Format("%03d%02d%02d", year - 1911, month, day);

}

void CDateEditBoxCtrl::OnKillFocus(CWnd* pNewWnd) 
{
	// Calc current date data.
	CalcData();

	// Calc text to show.
	CalcText();

	// Hide the caret.
	HideCaret();

	// Should destroy the caret when lost focus.
	DestroyCaret();

	// Redraw.
	InvalidateControl();

	// Fire event.
	FireChange();
	
	COleControl::OnKillFocus(pNewWnd);
}

void CDateEditBoxCtrl::MoveRight(int nSteps)
/*
Routine Description:
	Move the caret right for specified steps.
		
Parameters:
	nSteps		The desired steps.
	
Return Value:
	None.
*/
{
	if (m_bMingGuo)
	{
		// It is freely Mingguo format.
		
		// Move current position right.
		nCurrentPosition += nSteps;
		
		// Assure that there is one char at least.
		if (strData.IsEmpty() || (strData.Right(1) != '_' && strData.GetLength() < 7))
			strData += CString("_");

		// Current position should be in valid range.
		if (nCurrentPosition >= strData.GetLength())
			nCurrentPosition = strData.GetLength() - 1;
		if (nCurrentPosition < 0)
			nCurrentPosition = 0;

		// Update the caret.
		InvalidateControl();
		return;
	}

	if (nSteps > 0 && (nCurrentPosition == nYearPosition[0] + nYearFinish)
		&& nCurrentPosition <= nYearPosition[1])
		return;

	if (nSteps > 0 && (nCurrentPosition == nMonthPosition[0] + nMonthFinish)
		&& nCurrentPosition <= nMonthPosition[1])
		return;
	
	if (nSteps > 0 && (nCurrentPosition == nDayPosition[0] + nDayFinish)
		&& nCurrentPosition <= nDayPosition[1])
		return;

	int nNewPosition = nCurrentPosition + nSteps;
	if (nNewPosition >= strFormat.GetLength())
		nNewPosition = strFormat.GetLength() - 1;

	while (nNewPosition <= strFormat.GetLength())
	{
		// Skip the delimiters.
		if (!( (nNewPosition >= nYearPosition[0] && nNewPosition <= nYearPosition[1])
			|| (nNewPosition >= nMonthPosition[0] && nNewPosition <= nMonthPosition[1])
			|| (nNewPosition >= nDayPosition[0] && nNewPosition <= nDayPosition[1])))
			nNewPosition ++;
		else
			break;
	}
	if (nNewPosition > strFormat.GetLength()) // Can not move ahead now.
	{
		// Move left instead.
		while (nNewPosition >= 0)
		{
			if (!( (nNewPosition >= nYearPosition[0] && nNewPosition <= nYearPosition[1])
				|| (nNewPosition >= nMonthPosition[0] && nNewPosition <= nMonthPosition[1])
				|| (nNewPosition >= nDayPosition[0] && nNewPosition <= nDayPosition[1])))
				nNewPosition --;
			else
				break;
		}
	}
	
	// Update the caret position.
	nCurrentPosition = nNewPosition;
	InvalidateControl();
}

void CDateEditBoxCtrl::MoveLeft(int nSteps)
/*
Routine Description:
	Move the caret in specified steps.
		
Parameters:
	nSteps		The desired steps.
	
Return Value:
	None.
*/
{
	if (m_bMingGuo)
	{
		// It is freely Mingguo format.
		
		// Move curren position left.
		nCurrentPosition -= nSteps;

		// Current position should be in valid range.
		if (nCurrentPosition >= strData.GetLength())
			nCurrentPosition = strData.GetLength() -1;
		if (nCurrentPosition < 0)
			nCurrentPosition = 0;

		// Update the caret position.
		InvalidateControl();
		return;
	}

	if (nSteps < 0 && (nCurrentPosition == nYearPosition[0] + nYearFinish)
		&& nCurrentPosition <= nYearPosition[1])
		return;

	if (nSteps < 0 && (nCurrentPosition == nMonthPosition[0] + nMonthFinish)
		&& nCurrentPosition <= nMonthPosition[1])
		return;
	
	if (nSteps < 0 && (nCurrentPosition == nDayPosition[0] + nDayFinish)
		&& nCurrentPosition <= nDayPosition[1])
		return;

	int nNewPosition = nCurrentPosition - nSteps;
	if (nNewPosition < 0)
		nNewPosition = 0;

	// Skip the delimiters.
	while (nNewPosition >= 0)
	{
		if (!( (nNewPosition >= nYearPosition[0] && nNewPosition <= nYearPosition[1])
			|| (nNewPosition >= nMonthPosition[0] && nNewPosition <= nMonthPosition[1])
			|| (nNewPosition >= nDayPosition[0] && nNewPosition <= nDayPosition[1])))
			nNewPosition --;
		else
			break;
	}
	if (nNewPosition < 0) // Can not move left, move right instead.
	{
		while (nNewPosition <= strFormat.GetLength())
		{
			if (!( (nNewPosition >= nYearPosition[0] && nNewPosition <= nYearPosition[1])
				|| (nNewPosition >= nMonthPosition[0] && nNewPosition <= nMonthPosition[1])
				|| (nNewPosition >= nDayPosition[0] && nNewPosition <= nDayPosition[1])))
				nNewPosition ++;
			else
				break;
		}
	}

	// Update the caret position.
	nCurrentPosition = nNewPosition;
	InvalidateControl();
}

BSTR CDateEditBoxCtrl::GetDateFormat() 
/*
Routine Description:
	Get the input mask.
	
Parameters:
	None.
	
Return Value:
	The mask string.
*/
{
	CString strResult;
	strResult = strFormat;

	return strResult.AllocSysString();
}

void CDateEditBoxCtrl::SetDateFormat(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the input mask.
	
Parameters:
	lpszNewValue		The new mask string.
	
Return Value:
	None.
*/
{
	// Should have license to modify properties.
	if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		return;

	CString strNewValue(lpszNewValue);
	strNewValue.MakeLower();

	// Validate new mask.
	if (IsValidFormat(strNewValue))
	{
		if (strNewValue == "ggggggg")
		// It is freely Mingguo mask.
			m_bMingGuo = TRUE;
		else
			m_bMingGuo = FALSE;

		strFormat = strNewValue;

		// Update the text according to new mask.
		CalcText();
		InvalidateControl();

		SetModifiedFlag();
		BoundPropertyChanged(dispidDateFormat);
	}
	else
		AfxMessageBox("Invalid Format!");
}

BOOL CDateEditBoxCtrl::IsValidFormat(CString strNewFormat)
/*
Routine Description:
	Decide if the given format is valid.
		
Parameters:
	strNewFormat		The format string to be validated.
	
Return Value:
	If it is a valid format, return TRUE; Otherwise, return FALSE.
*/
{
	if (strNewFormat == "ggggggg")
	// It is freely Mingguo foramt.
		return TRUE;

	// The current char.
	TCHAR chr;
	
	// The previous char.
	TCHAR chrPrev = '/';
	int i = 0;
	
	// The length of each part.
	int nYearPosition = 0, nMonthPosition = 0, nDayPosition = 0;
	
	// If each part appears.
	BOOL bHasYear = FALSE, bHasMonth = FALSE, 
		bHasDay = FALSE;
	
	for (i = 0; i < strNewFormat.GetLength(); i ++)
	{
		// Get one char.
		chr = strNewFormat[i];

		switch (chr)
		{
			case 'y':
			{
				// It is in regular year part.
				if (chrPrev == chr)
				{
					// Year length can not exceed 4.
					if (++ nYearPosition > 4)
						return FALSE;
				}
				else
				{
					// The start point of a year part.
					
					if (bHasYear != TRUE)
					{
						// There is no start point yet. It's ok.
						
						// Update the previous char.
						chrPrev = 'y';
						
						// Year format now begins.
						bHasYear = TRUE;
						if (++ nYearPosition > 4)
							return FALSE;
					}
					else
					// There is already a start point of year part, it's a invalid format.
						return FALSE;
				}
			}
			break;

			case 'g':
			{
				// It is in Mingguo year part.
				if (chrPrev == chr)
				{
					// Year length can not exceed 4.
					if (++ nYearPosition > 3)
						return FALSE;
				}
				else
				{
					// It is the start point of a year part.
					
					if (bHasYear != TRUE)
					{
						// Update the previous char.
						chrPrev = 'g';
						
						// Year part now begins.
						bHasYear = TRUE;
						if (++ nYearPosition > 3)
							return FALSE;
					}
					else
					// There is already a start point now, so it's a invalid format.
						return FALSE;
				}
			}
			break;

			case 'm':
			{
				// It is in month part.
				
				if (chrPrev == chr)
				{
					// Month length can not exceed 2.
					if (++ nMonthPosition > 2)
						return FALSE;
				}
				else
				{
					// It is the start point of a month part.
					if (bHasMonth != TRUE)
					{
						// Month now begins.
						bHasMonth = TRUE;
						if (++ nMonthPosition > 2)
							return FALSE;
					}
					else
					// There is already a start point of month part, so it's a invalid format.
						return FALSE;
				}
				
				// Update the previous char.
				chrPrev = 'm';
			}
			break;

			case 'd':
			{
				// It is in day part.
				
				if (chrPrev == chr)
				{
					// The day part length can exceed 2.
					if (++ nDayPosition > 2)
						return FALSE;
				}
				else
				{
					// It is the start point of a day part.
					
					if (bHasDay != TRUE)
					{
						// The day part begins.
						bHasDay = TRUE;
						if (++ nDayPosition > 2)
							return FALSE;
					}
					else
					// There is already a start of day part, so it's a invalid format.
						return FALSE;
				}
				
				// Update the previous char.
				chrPrev = 'd';
			}
			break;

			default:
			{
				// Previous character is delimiter
				chrPrev = '/';
				break;
			}
		}
	}

	// 3 parts can not be all empty.
	if (nYearPosition == 0 && nMonthPosition == 0 && nDayPosition == 0)
		return FALSE;
	else
		return TRUE;
}

BSTR CDateEditBoxCtrl::GetDisplayDateFormat() 
/*
Routine Description:
	Get the display mask while have no focus.
	
Parameters:
	None.
	
Return Value:
	The mask string.
*/
{
	CString strResult;
	strResult = strDisplayFormat;

	return strResult.AllocSysString();
}

void CDateEditBoxCtrl::SetDisplayDateFormat(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the mask while have no focus.
	
Parameters:
	lpszNewValue		The new mask string.
	
Return Value:
	None.
*/
{
	if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		return;

	CString strNewValue(lpszNewValue);
	strNewValue.MakeLower();
	if (IsValidFormat(strNewValue))
	{
		if (strNewValue == "ggggggg")
			m_bDisplayMingGuo = TRUE;
		else
			m_bDisplayMingGuo = FALSE;
		strDisplayFormat = strNewValue;
		CalcText();
		InvalidateControl();
		SetModifiedFlag();
		BoundPropertyChanged(dispidDisplayDateFormat);
	}
	else
		AfxMessageBox("Invalid DisplayFormat!");
}

BOOL CDateEditBoxCtrl::PreTranslateMessage(MSG* pMsg) 
{
	if (pMsg->hwnd == this->m_hWnd && pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_DOWN || pMsg->wParam == VK_UP || 
		pMsg->wParam == VK_LEFT || pMsg->wParam == VK_RIGHT))
	{
		OnKeyDown((int)pMsg->wParam, 1, 1);
		return TRUE;
	}

	return COleControl::PreTranslateMessage(pMsg);
}

void CDateEditBoxCtrl::OnFontChanged() 
{
	if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		return;

	COleControl::OnFontChanged();
}

void CDateEditBoxCtrl::OnForeColorChanged() 
{
	if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		return;

	COleControl::OnForeColorChanged();

	m_ForeColor = GetForeColor();

	BoundPropertyChanged(DISPID_FORECOLOR);
}

void CDateEditBoxCtrl::DisplayCaret()
/*
Routine Description:
	Display a caret.
	
Parameters:
	None.
	
Return Value:
	None.
*/
{
	// Hide caret first.
	HideCaret();

	CDC * pDC = GetDC();

	CFont* pFntPrev = SelectStockFont(pDC);
	// The caret should cover the next character.
	ptCaretPosition.x = (pDC->GetOutputTextExtent(strData,
		nCurrentPosition)).cx;
	// The length of the total text.
	CSize sz = pDC->GetOutputTextExtent(strData);
	CRect rect;
	GetClientRect(&rect);
	// Move the caret to the correct position according to the text
	// alignment.
	if (m_ta == Center)
		ptCaretPosition.x += (rect.Width() - sz.cx) / 2;
	if (m_ta == Right)
		ptCaretPosition.x += rect.Width() - sz.cx;

	CSize szCaret;
	// Should not exceed the array's range, or the program will crash.
	szCaret = pDC->GetOutputTextExtent(nCurrentPosition >= strData.GetLength() ? 'A' : CString(strData[nCurrentPosition]), 1);
	// To show the caret, need creating it first.
	::CreateCaret(this->m_hWnd, NULL, szCaret.cx, szCaret.cy);
	ShowCaret();
	SetCaretPos(ptCaretPosition);

	/*	CRect rect(ptCaretPosition, szCaret);
	CBrush br(RGB(125, 125, 125));
	pDC->FillRect(&rect, &br);*/
	
	pDC->SelectObject(pFntPrev);
	ReleaseDC(pDC);
}

// We use the "value" property to exchange data with outside world
BSTR CDateEditBoxCtrl::GetText() 
/*
Routine Description:
	Get the current text.
	
Parameters:
	None.
	
Return Value:
	Current text string.
*/
{
	CalcData();
	CalcText();
	return m_strText.AllocSysString();
}

void CDateEditBoxCtrl::SetText(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set current text..
	
Parameters:
	lpszNewValue		The new value.
	
Return Value:
	None.
*/
{
	// Can not modify properties having no license.
	if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		return;

	if (SetTextIn(lpszNewValue))
	{
		// The given text is valid.
		InvalidateControl();
		BoundPropertyChanged(dispidText);
		SetModifiedFlag();
		FireChange();
	}
	else
	{
		// The given text is invalid, use null data instead.
		SetTextIn("");
		InvalidateControl();
		SetModifiedFlag();
		BoundPropertyChanged(dispidText);
		FireChange();
	}
}

void CDateEditBoxCtrl::OnReadOnlyChanged() 
{
	if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		return;

	InvalidateControl();
	SetModifiedFlag();
	BoundPropertyChanged(dispidReadOnly);
}

short CDateEditBoxCtrl::GetTextAlign() 
{
	return m_ta;
}

void CDateEditBoxCtrl::SetTextAlign(short nNewValue) 
{
	if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		return;

	m_ta = nNewValue % 3;
	InvalidateControl();

	SetModifiedFlag();
	BoundPropertyChanged(dispidTextAlign);
}

BOOL CDateEditBoxCtrl::IsFinished()
/*
Routine Description:
	Decides if there is a valid date data.
	
Parameters:
	None.
	
Return Value:
	If there is, return TRUE; Otherwise, return FALSE.
*/
{
	return (nYearFinish >= nYearLength && nMonthFinish >= nMonthLength
		&& nDayFinish >= nDayLength);
}

void CDateEditBoxCtrl::OnLButtonDown(UINT nFlags, CPoint point) 
{
	// Avoid showing caret at the position of a delimiter by clicking mouse button.
	
	COleControl::OnLButtonDown(nFlags, point);

	int pos = nCurrentPosition;

	CDC * pDC = GetDC();
	pDC->DPtoLP(&point);
	CSize sz = pDC->GetOutputTextExtent(strData);
	CSize sz2 = pDC->GetOutputTextExtent("2");
	CRect rect(CPoint(0, 0), sz);
	if (rect.PtInRect(point))
	{
		int i = (int) (point.x / sz2.cx);
		int j;
		if (i > pos)
			for (j = 0; j <= i - pos; j ++)
				MoveRight(1);
		if (i < nCurrentPosition)
			for (j = 0; j < pos - i; j ++)
				MoveLeft(1);
	}

	ReleaseDC(pDC);
}

void CDateEditBoxCtrl::SetDate(short nYear, short nMonth, short nDay) 
/*
Routine Description:
	Modifys the underlying data.
		
Parameters:
	nYear		The year data.
	
	nMonth		The month data.
	
	nDay		The day data.
	
Return Value:
	None.
*/
{
	// Need license to modify data.
	if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		return;

	// Calc current data first.
	CalcText();

	// Validate the given data.
	
	COleDateTime date;
	int nNewYear = nYear;
	
	if (m_bMingGuo || m_bFormatMingGuo)
		nNewYear = nYear + 1911;
	date.SetDate(nNewYear, nMonth, nDay);
	if (date.GetStatus() == COleDateTime::valid)
	{
		m_bTextNull = FALSE;
		year = nNewYear;
		month = nMonth;
		day = nDay;
		m_PrevData.wYear = year;
		m_PrevData.wMonth = month;
		m_PrevData.wDay = day;
	}
	
	// Update the window text.
	CalcText();
	InvalidateControl();
}

void CDateEditBoxCtrl::CalcMingGuo()
/*
Routine Description:
	Calc the correct text should be displayed in freely Mingguo fromat.
	
Parameters:
	None.
	
Return Value:
	None.
*/
{
	int i;
	char chr;
	CString strYear, strMonth, strDay, strTotal;
	int nNewYear, nNewMonth, nNewDay;
	nNewYear = nNewMonth = nNewDay = 0;
	BOOL bValid = TRUE;
	strYear.Empty();
	strMonth.Empty();
	strDay.Empty();
	strTotal.Empty();

	if (strData.IsEmpty())
	// Text is null, it's valid. Return now.
		return;
	else
	{
		for (i = 0; i < strData.GetLength(); i++)
		{
			// Filter invalid chars out.
			chr = strData[i];
			if ((chr >= '0' && chr <= '9') || (chr == '-' &&
				strTotal.IsEmpty()))
				strTotal += chr;
		}

		// Verify the filtered text.
		if (strTotal.IsEmpty() || (strTotal[0] == '-' && strTotal.
			GetLength() < 6) || (strTotal[0] != '-' && strTotal.
			GetLength() < 5))
			return;
		
		// Get the 3 parts of data. The format is "yyymmdd" or "-yyymmdd".
		strDay = strTotal.Right(2);
		strMonth = strTotal.Mid(strTotal.GetLength() - 4, 2);
		strYear = strTotal.Left(strTotal.GetLength() - 4);
		
		// Filter the negative signal out.
		i = strYear[0] == '-' ? 1 : 0;

		// Calc each part of data.
		
		for (; i < strYear.GetLength(); i++)
			nNewYear = nNewYear * 10 + strYear[i] - '0';
		if (strYear[0] == '-')
		// Year is negative.
			nNewYear = -nNewYear;
		for (i = 0; i < 2; i++)
			nNewMonth = nNewMonth * 10 + strMonth[i] - '0';
		for (i = 0; i < 2; i++)
			nNewDay = nNewDay * 10 + strDay[i] - '0';

		// Validate the result data.
		COleDateTime date;
		date.SetDate(nNewYear + 1911, nNewMonth, nNewDay);
		if (date.GetStatus() == COleDateTime::valid)
		// The data is valid.
			bValid = TRUE;
		else
		// The data is invalid.
			return;
	}

	if (bValid)
	// Update window text now.
		SetDate(nNewYear, nNewMonth, nNewDay);
}

void CDateEditBoxCtrl::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	COleControl::OnLButtonDblClk(nFlags, point);

	// Let the user select a date via a calendar

	if (dlgDatePicker.DoModal() == IDOK && !m_bReadOnly)
	{
		year = dlgDatePicker.m_Date.GetYear();
		month = dlgDatePicker.m_Date.GetMonth();
		day = dlgDatePicker.m_Date.GetDay();
		m_bTextNull = FALSE;
		nYearFinish = nMonthFinish = nDayFinish = 4;
		nCurrentPosition = 0;
		CalcText();
		MoveRight();
		FireChange();
		InvalidateControl();
	}
}

BOOL CDateEditBoxCtrl::SetTextIn(CString strNewValue)
/*
Routine Description:
	Calc new text from given text and data type.
		
Parameters:
	strNewValue		The desired new text.
	
	nDataType		The type of underlying data.
	
Return Value:
	None.
*/
{
	CalcText();

	// Reset these variables.
	CString strData = strNewValue;
	int i;
	char chr;
	CString strYear, strMonth, strDay, strTotal;
	int nNewYear, nNewMonth, nNewDay;
	nNewYear = nNewMonth = nNewDay = 0;
	BOOL bValid = TRUE;
	strYear.Empty();
	strMonth.Empty();
	strDay.Empty();
	strTotal.Empty();

	if (strData.IsEmpty() || strData == '_')
	{
		// The new data is empty.
		
		m_bTextNull = TRUE;
		nYearFinish = nMonthFinish = nDayFinish = 0;
		nCurrentPosition = 0;
	}
	else
	{
		// Fillter invalid chars out.
		for (i = 0; i < strData.GetLength(); i++)
		{
			chr = strData[i];
			if ((chr >= '0' && chr <= '9') || (chr == '-' &&
				strTotal.IsEmpty()))
				strTotal += chr;
		}

		// The length should be enough.
		if (strTotal.IsEmpty() || (strTotal[0] == '-' && strTotal.
			GetLength() < 6) || (strTotal[0] != '-' && strTotal.
			GetLength() < 5))
			return FALSE;
		
		strDay = strTotal.Right(2);
		strMonth = strTotal.Mid(strTotal.GetLength() - 4, 2);
		strYear = strTotal.Left(strTotal.GetLength() - 4);
	
		// Is the year negative?
		i = strYear[0] == '-' ? 1 : 0;
		if (i == 1 && !m_bMingGuo)
			return FALSE;

		// Calc new data.
		for (; i < strYear.GetLength(); i++)
			nNewYear = nNewYear * 10 + strYear[i] - '0';
		if (strYear[0] == '-')
			nNewYear = -nNewYear;
		for (i = 0; i < 2; i++)
			nNewMonth = nNewMonth * 10 + strMonth[i] - '0';
		for (i = 0; i < 2; i++)
			nNewDay = nNewDay * 10 + strDay[i] - '0';

		// Verify new data.
		COleDateTime date;
		date.SetDate(nNewYear + ((m_bMingGuo || m_bFormatMingGuo)? 
			1911 : 0), nNewMonth, nNewDay);
		if (date.GetStatus() == COleDateTime::valid)
		// The new data is valid.
			bValid = TRUE;
		else
		// The new dta is invalid.
			return FALSE;
	}

	if (bValid)
	{
		if (!(strData.IsEmpty() || strData == '_'))
			// The new data is not null.
			SetDate(nNewYear, nNewMonth, nNewDay);
		else
		{
			// The new data is null.
			month = day = 1;
			CalcText();
		}

		m_strText = strData;
	}

	return TRUE;
}

void CDateEditBoxCtrl::OnBackColorChanged() 
{
	if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		return;

	COleControl::OnBackColorChanged();

	m_BackColor = GetBackColor();

	BoundPropertyChanged(DISPID_BACKCOLOR);
}

void CDateEditBoxCtrl::OnCanYearNegativeChanged() 
{
	if (!((CDateEditBoxApp *)AfxGetApp())->m_bLicensed)
		return;

	SetModifiedFlag();
	BoundPropertyChanged(dispidCanYearNegative);
}

void CDateEditBoxCtrl::OnKeyDownEvent(USHORT nChar, USHORT nShiftState) 
{
	switch (nChar)
	{
		case VK_HOME:
		{
			MoveLeft(nCurrentPosition);
		}
		break;

		case VK_UP:
		{
			MoveToNextControl(FALSE);
			break;
		}
		case VK_LEFT:
		{
			MoveLeft(1);
		}
		break;

		case VK_END:
		{
			int pos = nCurrentPosition;
			for (int i = 0; i < strData.GetLength() - pos - 1; i++)
				MoveRight(1);
		}
		break;

		case VK_RIGHT:
		{
			MoveRight(1);
		}
		break;

		case VK_DELETE:
		{
			OnKeyPressEvent(VK_DELETE);
		}
		break;

		case VK_DOWN:
		case VK_RETURN:
		{
			MoveToNextControl();
		}
		break;
	}

	COleControl::OnKeyDownEvent(nChar, nShiftState);
}

void CDateEditBoxCtrl::OnKeyPressEvent(USHORT nChar) 
{
	BOOL bFinish = TRUE;

	if (m_bReadOnly)
		return;

	if (m_bMingGuo)
	{
		// It is in freely Mingguo format.
		
		if ((nChar < '0' || nChar > '9') && nChar != 
			VK_BACK && nChar != VK_DELETE && nChar != '/' && nChar != '-')
			return;
		if (nChar == VK_BACK)
		{
			// Delete previous char.
			if (nCurrentPosition == 0 && strData.IsEmpty())
				return;

			// Move the right data in the same part left.
			strData = strData.Left(nCurrentPosition - 1) + strData.
				Right(strData.GetLength() - nCurrentPosition);
			if (strData.IsEmpty() || strData == '_')
			// The data is null.
				SetText("");
			if (strData.IsEmpty() || strData.Right(1) != '_')
			// Replace previous char with '_'.
				strData += CString("_");

			// Move the caret.
			MoveLeft(1);

			// Update the window text.
			InvalidateControl();
			return;
		}

		if (nChar == VK_DELETE)
		{
			// Delete current char.
			strData = strData.Left(nCurrentPosition) + strData.
				Right(strData.GetLength() - nCurrentPosition - 1);
			if (strData.IsEmpty() || strData == '_')
				SetText("");
			if (strData.IsEmpty() || strData.Right(1) != '_')
				strData += CString("_");

			// Move the caret left.
			MoveLeft();

			// Update the window text.
			InvalidateControl();
			return;
		}

		else
		// Replace current char with new char.
			if (strData.IsEmpty())
				strData += (char)nChar;
			else
				strData.SetAt(nCurrentPosition, (char)nChar);

		if (strData.IsEmpty() || (strData.Right(1) != '_' && strData.GetLength() < 7))
			strData += CString("_");

		// Now the value is accpeted.
		if (strData.GetLength() == 7 && strData.Right(1) != '_')
			FireChange();

		// Move the caret right.
		MoveRight(1);

		// Update the window text.
		InvalidateControl();
		return;
	}

	if ((nChar < '0' || nChar > '9') && nChar != VK_BACK && nChar !=
		VK_DELETE && nChar != '/' && nChar != '-') // Invalid characters.
		return;

	COleDateTime temptime;
	CString text;
	text = strData;

	int i = 0;
	int nNewYear = 0, nNewMonth = 0, nNewDay = 0;

	// Is new data valid?
	BOOL bVerify = TRUE;

	if ((nChar == VK_BACK || nChar == VK_DELETE) && IsFinished()
		&& !m_bTextNull)
	{
		// The data will be modified, save current data now.
		
		m_PrevData.wYear = year;
		m_PrevData.wMonth = month;
		m_PrevData.wDay = day;
	}

	// One part of data will be empty, so there will not be valid data.
	if (nChar == VK_BACK || nChar == VK_DELETE)
		bFinish = FALSE;

	int nNewChar = nChar - '0';

	if (nCurrentPosition >= nYearPosition[0] 
		&& nCurrentPosition <= nYearPosition[1]) 
	{
		// Is is in year part.
		
		// Is the year is negatve?
		BOOL bNegative = FALSE;

		// Only the first char in Mingguo foramt can be '-'.
		if (nChar == '-' && !(m_bCanYearNegative && m_bFormatMingGuo
			&& nCurrentPosition == nYearPosition[0]))
			return;

		if (!bFinish)
		{
			// Not finish yet.
			
			// Replace the rest chars with '_'.
			for (i = nCurrentPosition; i <= nYearPosition[1]; i++)
				strData.SetAt(i, '_');

			// The length of year part which has been filled.
			nYearFinish = nCurrentPosition - nYearPosition[0];
			if (nYearFinish == 0 && nMonthFinish == 0 && nDayFinish == 0)
			{
				// The data is null.
				SetText("");
				return;
			}

			i = 0;
			if (text[nYearPosition[0]] == '-')
			{
				// The year is negative.
				bNegative = TRUE;
				i ++;
			}

			// Calc year data.
			for (; i < nYearFinish; i++)
				nNewYear = nNewYear * 10 + text[nYearPosition[0] + i] - '0';
			if (bNegative)
			// Year is negative.
				nNewYear = -nNewYear;

			year = nNewYear + (m_bFormatMingGuo ? 1911 : 0);

			if (nChar == VK_BACK)
			// Move the caret left.
				MoveLeft(1);

			// Update the window text.
			InvalidateControl();
			return;
		}

		// Curent position is the only empty position in year part.
		if (nChar != '/')
		{
			i = 0;
			if (text[nYearPosition[0]] == '-' || (nCurrentPosition 
				== nYearPosition[0] && nChar == '-'))
			{
				// Year is negative.
				bNegative = TRUE;
				i ++;
			}
			if (nCurrentPosition == nYearPosition[0] && 
				nChar != '-' && bNegative)
			{
				// Year is not negative.
				bNegative = FALSE;
				i --;
			}

			// Calc the year data.
			for (; i < nCurrentPosition - nYearPosition[0]; i++)
				nNewYear = nNewYear * 10 + text[nYearPosition[0] + i] - '0';
			if (!bNegative || nCurrentPosition != nYearPosition[0])
				nNewYear = nNewYear * 10 + nNewChar;

			// The finished length increases.
			if (nCurrentPosition - nYearPosition[0] == nYearFinish)
				nYearFinish ++;
			if (nYearFinish >= nYearLength)
			{
				// The year part is finished.
				for (i = 1; i <= nYearPosition[1] - nCurrentPosition; i++)
					nNewYear = nNewYear * 10 + text[nCurrentPosition + i] - '0';
			}
			if (bNegative)
			// The year is negative.
				nNewYear = -nNewYear;

			// Validate the new data.
			nNewYear += m_bFormatMingGuo ? 1911 : 0;
			temptime.SetDate(nNewYear, month, day);
			if (temptime.GetStatus() == COleDateTime::invalid && (
				IsFinished() || (nCurrentPosition == nYearPosition[1]
				&& !m_bTextNull)))
			{
				// The new data is invalid.
				bVerify = FALSE;
				if (text[nCurrentPosition] == '_')
				// The finished length should be descreased.
					nYearFinish --;
			}
			else if (nYearLength != 0)
			{
				// The new data is valid.
				int i;
				for (i = 0; i < nYearLength; i++)
					year /= 10;
				for (i = 0; i < nYearLength; i++)
					year *= 10;
				year = nNewYear + (m_bFormatMingGuo ? 0 : year);
			}
		}
		else
		{
			// If the user press '/', I should supplement the space
			// he left, and move to next field.
			
			i = 0;
			if (text[nYearPosition[0]] == '-')
			{
				// The year is negative.
				bNegative = TRUE;
				i ++;
			}

			// Calc the new year data.
			for (; i < nYearLength && text[i + nYearPosition[0]] != '_' ; i++)
				nNewYear = nNewYear * 10 + text[nYearPosition[0] + i] - '0';
			if (bNegative)
				nNewYear = -nNewYear;
			nNewYear += m_bFormatMingGuo ? 1911 : 0;

			// Validate the new data.
			temptime.SetDate(nNewYear, month, day);

			// Now, variable i is the number of space the the user input.
			if (temptime.GetStatus() == COleDateTime::valid)
			{
				// The new data is valid.
				
				// Fill the year data.
				int j;
				for (j = 1; j <= i; j ++)
					strData.SetAt(nYearLength - j + nYearPosition[0], strData[i - j + nYearPosition[0]]);

				// We finished the year part.
				nYearFinish = 4;

				j = 0;
				if (bNegative)
				{	
					// The year is negative.
					strData.SetAt(nYearPosition[0], '-');
					j = 1;
				}
				
				// The left chars should be filled with '0'.
				for (; j < nYearLength - i; j ++)
					strData.SetAt(j + nYearPosition[0], '0');
				
				// Move the caret to next field.	
				i = nCurrentPosition;
				for (j = 0; j <= nYearPosition[1] - i; j ++)
					MoveRight(1);
			}
		}
	}

	else if (nCurrentPosition >= nMonthPosition[0] 
		&& nCurrentPosition <= nMonthPosition[1]) 
	{
		// It is in month part.
		
		if (!bFinish)
		{
			// We have not finished month part.
			
			// Fill all chars in month part with '_'.
			for (i = nMonthPosition[0]; i <= nMonthPosition[1]; i++)
				strData.SetAt(i, '_');
			nMonthFinish = 0;

			// The month is empty, use 1 instead.
			month = 1;

			// Move the caret left.
			MoveLeft(nCurrentPosition - nMonthPosition[0]);
			if (nChar == VK_BACK)
				MoveLeft(1);

			// Update the window text.
			InvalidateControl();
			return;
		}

		if (nChar != '/')
		{
			// Calc the new month data.
			
			for (i = 0; i < nCurrentPosition - nMonthPosition[0]; i++)
				nNewMonth = nNewMonth * 10 + text[nMonthPosition[0] + i] - '0';
			nNewMonth = nNewMonth * 10 + nNewChar;
			for (i = 1; i <= nMonthPosition[1] - nCurrentPosition; i++)
			{
				if (text[nCurrentPosition + i] == '_')
					nNewMonth = nNewMonth * 10;
				else
					nNewMonth = nNewMonth * 10 + text[nCurrentPosition + i] - '0';
			}

			// Validate the new month data.
			temptime.SetDate(year, nNewMonth, day);
			if (temptime.GetStatus() == COleDateTime::invalid)
			{
				// The new data is invalid, guess the valid text.
				
				bVerify = FALSE;

				//How about use this char as the month data instead?
				temptime.SetDate(year, nChar - '0', day);

				if (temptime.GetStatus() == COleDateTime::valid)
				{
					// That's ok, replace other digits with 0.
					for (i = nMonthPosition[0]; i <= nMonthPosition[1]; i++)
						strData.SetAt(i, '0');

					// If the new month is too big, only keep the lowerst
					// position number.
					nCurrentPosition = i - 1;
					month = nChar - '0';
					nMonthFinish = nMonthLength;
					bVerify = TRUE;
				}
				else if (nChar == '0' && nCurrentPosition == nMonthPosition[0])
				{
					// This is a '0' char in first position, move to the second position.
					bVerify = TRUE;
					nMonthFinish ++;
				}
			}

			else if (nMonthLength != 0)
			{
				// We may have finished month part.
				
				// Calc new month data.
				int i;
				for (i = 0; i < nMonthLength; i++)
					month /= 10;
				for (i = 0; i < nMonthLength; i++)
					month *= 10;
				month += nNewMonth;
				nMonthFinish ++;
			}
		}
		else
		{
			// If the user press '/', I should supplement the space
			// he left, and move to next field.
			for (i = 0; i < nMonthLength && text[i + nMonthPosition[0]] != '_' ; i++)
				nNewMonth = nNewMonth * 10 + text[nMonthPosition[0] + i] - '0';
			temptime.SetDate(year, nNewMonth, day);

			// Now, i is the number of space the the user input.
			if (temptime.GetStatus() == COleDateTime::valid)
			{
				int j;
				for (j = 1; j <= i; j ++)
					strData.SetAt(nMonthPosition[0] + nMonthLength - j, strData[i - j + nMonthPosition[0]]);
				nMonthFinish = 4;
				for (j = 0; j < nMonthLength - i; j ++)
					strData.SetAt(j + nMonthPosition[0], '0');
				i = nCurrentPosition;
				for (j = 0; j <= nMonthPosition[1] - i; j ++)
					MoveRight(1);

				month = nNewMonth;
			}
		}
	}

	// change the Day text
	else if (nCurrentPosition >= nDayPosition[0] 
		&& nCurrentPosition <= nDayPosition[1]) 
	{
		// It is in day part.
		if (!bFinish)
		{
			// We have not finished day part yet.
			// Replace all position in day part with '_'.
			for (i = nDayPosition[0]; i <= nDayPosition[1]; i++)
				strData.SetAt(i, '_');
			nDayFinish = 0;
			day = 1;

			// Move the caret left.
			MoveLeft(nCurrentPosition - nDayPosition[0]);
			if (nChar == VK_BACK)
				MoveLeft(1);
			InvalidateControl();
			return;
		}

		if (nChar != '/')
		{
			// Calc the new day data.
			
			for (i = 0; i < nCurrentPosition - nDayPosition[0]; i++)
				nNewDay = nNewDay * 10 + text[nDayPosition[0] + i] - '0';
			nNewDay = nNewDay * 10 + nNewChar;
			for (i = 1; i <= nDayPosition[1] - nCurrentPosition; i++)
			{
				if (text[nCurrentPosition + i] == '_')
					nNewDay = nNewDay * 10;
				else
					nNewDay = nNewDay * 10 + text[nCurrentPosition + i] - '0';
			}

			// Validate the new data.
			temptime.SetDate(year, month, nNewDay);
			if (temptime.GetStatus() == COleDateTime::invalid)
			{
				// The new data is invalid. We should guess the correct data now.
				
				bVerify = FALSE;

				// Eliminate the lower position number to make it valid.
				temptime.SetDate(year, month, (nChar - '0') * 10);
				if (temptime.GetStatus() == COleDateTime::valid && 
					nCurrentPosition == nDayPosition[0])
				{
					// Ok.
					for (i = nDayPosition[0]; i <= nDayPosition[1]; i++)
						strData.SetAt(i, '0');

					day = (nChar - '0') * 10;
					nCurrentPosition = nDayPosition[0];
					nDayFinish = nDayLength;
					bVerify = TRUE;
				}
				// If the method above does not work, eliminate the lower.
				else
				{
					temptime.SetDate(year, month, nChar - '0');

					if (temptime.GetStatus() == COleDateTime::valid)
					{
						// Ok.
						for (i = nDayPosition[0]; i <= nDayPosition[1]; 
							i++)
							strData.SetAt(i, '0');

						nCurrentPosition = i - 1;
						day = nChar - '0';
						nDayFinish = nDayLength;
						bVerify = TRUE;
					}
					else if (nChar == '0' && nCurrentPosition == nDayPosition[0])
					{
						// This is a '0' at the first position, move to the second position.
						bVerify = TRUE;
						nDayFinish ++;
					}
				}
			}
			else if (nDayLength != 0)
			{
				// We have finished the day part.
				
				int i;
				for (i = 0; i < nDayLength; i++)
					day /= 10;
				for (i = 0; i < nDayLength; i++)
					day *= 10;
				day += nNewDay;
				nDayFinish ++;
			}
		}
		else
		{
			// If the user press '/', I should supplement the space
			// he left, and move to next field.
			
			for (i = 0; i < nDayLength && text[i + nDayPosition[0]] != '_' ; i++)
				nNewDay = nNewDay * 10 + text[nDayPosition[0] + i] - '0';
			temptime.SetDate(year, month, nNewDay);

			// Now, i is the number of space the the user input.
			if (temptime.GetStatus() == COleDateTime::valid)
			{
				int j;
				for (j = 1; j <= i; j ++)
					strData.SetAt(nDayLength - j + nDayPosition[0], strData[i - j + nDayPosition[0]]);
				nDayFinish = 4;
				for (j = 0; j < nDayLength - i; j ++)
					strData.SetAt(j + nDayPosition[0], '0');
				i = nCurrentPosition;
				for (j = 0; j <= nDayPosition[1] - i; j ++)
					MoveRight(1);
				
				day = nNewDay;
			}
		}

	}

	if (bVerify)
	{
		// The new char is valid.
		if (nChar != '/')
		{
			// Update current character
			strData.SetAt(nCurrentPosition, (char)nChar);

			// Move the caret right
			MoveRight(1);

			// skip the delimiters.
			MoveRight();
		}

		if (IsFinished())
		{
			// We got a new date data.
			m_bTextNull = FALSE;
			FireChange();
		}

		// Update the window text.
		InvalidateControl();
	}
	
	COleControl::OnKeyPressEvent(nChar);
}

void CDateEditBoxCtrl::CalcData()
{
/*
Routine Description:
	Calc the data in special cases.	

Parameters:
	
Return Value:
	None.
*/
	if (!IsFinished())
	{
		// There is no valid data yet.
		year = m_PrevData.wYear;
		month = m_PrevData.wMonth;
		day = m_PrevData.wDay;
		if (!m_bTextNull)
			nYearFinish = nMonthFinish = nDayFinish = 4;
		else
		{
			nYearFinish = nMonthFinish = nDayFinish = 0;
			nCurrentPosition = 0;
		}
	}
	if (m_bMingGuo)
	// It is freely Mingguo format now.
		CalcMingGuo();
}

void CDateEditBoxCtrl::OnCtlIndexChanged() 
{
	SetModifiedFlag();
	BoundPropertyChanged(dispidCtlIndex);
}

void CDateEditBoxCtrl::OnCtlStopChanged() 
{
	SetModifiedFlag();
	BoundPropertyChanged(dispidCtlStop);
}

void CDateEditBoxCtrl::MoveToNextControl(BOOL bForward)
/*
Routine Description:
	Move to control on form which has the next ctlindex or tabindex.

Parameters:
	bForward		If move forward.
	
Return Value:
	None.
*/
{
	LPOLECONTAINER lpContainer;   
	LPENUMUNKNOWN lpEnumUnk;
	// Note that the IOleContainer interface is currently defined as
	// mandatory. It must be implemented by control containers,
	// in the OLE Control Containers Guidelines.
	HRESULT hr = m_pClientSite->GetContainer(&lpContainer);   
	if(FAILED(hr)) 
	{
#ifdef _DEBUG
		AfxMessageBox("Unable to get to the container");
#endif
		return;
	}   
	// OLECONTF_EMBEDDINGS is used to retrieve OLE Controls.
	// OLECONTF_OTHERS is used to retrieve other objects such as
	// Visual Basic internal controls   
	hr = lpContainer->EnumObjects(OLECONTF_EMBEDDINGS | OLECONTF_OTHERS, &lpEnumUnk);   
	if(FAILED(hr)) 
	{
		lpContainer->Release();      
		return;   
	}   

	LPUNKNOWN lpUnk;
	int nIndex = -1, nIndexLow = -1, nIndexLowest = -1, nIndexHigh = -1, nIndexHighest = -1;
	short nCtlIndex = -1, nCtlIndexLow, nCtlIndexLowest, nCtlIndexHigh, nCtlIndexHighest;
	nCtlIndexLow = nCtlIndexLowest = nCtlIndexHigh = nCtlIndexHighest = m_ctlIndex ;
	BOOL bDone = FALSE;
	LPDISPATCH lpDisp;

	while (lpEnumUnk->Next(1, &lpUnk, NULL) == S_OK && !bDone) 
	{
		nIndex ++;
		lpDisp = GetDispatchFromUnk(lpUnk);
		
		if(lpDisp) 
		{         
			nCtlIndex = GetCtlIndexProperty(lpDisp);
			if ((bForward && nCtlIndex == m_ctlIndex + 1) ||
				(!bForward && nCtlIndex == m_ctlIndex - 1 && nCtlIndex > -1))
				bDone = TRUE;
			else if (nCtlIndex != m_ctlIndex && nCtlIndex != -1)
			{
				if (nCtlIndex > nCtlIndexHighest)
				{
					nIndexHighest = nIndex;
					nCtlIndexHighest = nCtlIndex;
				}
				else if (nCtlIndex < nCtlIndexLowest &&  nCtlIndex != -1)
				{
					nIndexLowest = nIndex;
					nCtlIndexLowest = nCtlIndex;
				}
				if (nCtlIndex > m_ctlIndex && ((nCtlIndexHigh == m_ctlIndex && nCtlIndex > nCtlIndexHigh) || 
					(nCtlIndexHigh != m_ctlIndex && nCtlIndex < nCtlIndexHigh)))
				{
					nIndexHigh = nIndex;
					nCtlIndexHigh = nCtlIndex;
				}
				else if (nCtlIndex < m_ctlIndex && ((nCtlIndexLow == m_ctlIndex && nCtlIndex < nCtlIndexLow) ||
					(nCtlIndexLow != m_ctlIndex && nCtlIndex > nCtlIndexLow)))
				{
					nIndexLow = nIndex;
					nCtlIndexLow = nCtlIndex;
				}
			}

			lpDisp->Release();
		}      // Release interface pointers.
				
		lpUnk->Release();
	}	// End of While statement   

	lpEnumUnk->Reset();
	if (bDone)
		lpEnumUnk->Skip(nIndex);
	else if (bForward)
	{
		if(nIndexHigh > -1)
		{
			lpEnumUnk->Skip(nIndexHigh);
			bDone = TRUE;
		}
/*		else if (nIndexLowest > -1)
		{
			lpEnumUnk->Skip(nIndexLowest);
			bDone = TRUE;
		}
*/	}
	else
	{
		if (nIndexLow > -1)
		{
			lpEnumUnk->Skip(nIndexLow);
			bDone = TRUE;
		}
/*		else if(nIndexHighest > -1)
		{
			lpEnumUnk->Skip(nIndexHighest);
			bDone = TRUE;
		}
*/	}

	if (bDone && lpEnumUnk->Next(1, &lpUnk, NULL) == S_OK)
	{
		lpDisp = GetDispatchFromUnk(lpUnk);
		if(lpDisp) 
		{         
			VARIANT va;         
			VariantInit(&va);
			DISPID dispid;         
			DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
			LPWSTR lpName[1] = {(WCHAR *)L"SetFocus"};
			HRESULT hr = lpDisp->GetIDsOfNames(IID_NULL, lpName, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
			if(SUCCEEDED(hr))
			{
				lpDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, 
					NULL, NULL, NULL);
			}

			lpDisp->Release();
		}
				
		lpUnk->Release();
	}
	// Final clean up   
	
	lpEnumUnk->Release();
	lpContainer->Release();
}

short CDateEditBoxCtrl::GetCtlIndexProperty(LPDISPATCH lpDisp)
/*
Routine Description:
	Get ctlindex property value of one control.

Parameters:
	lpDisp		The IDispatch pointer of the control.

Return Value:
	The property value.
*/
{
	VARIANT va;         
	VariantInit(&va);
	DISPID dispid;         
	DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
	// Get the names of all the controls present in a VB form.
	LPWSTR lpName[1] = { (WCHAR *)L"CtlIndex" };
	HRESULT hr = lpDisp->GetIDsOfNames(IID_NULL, lpName, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if(SUCCEEDED(hr)) 
	{
		hr = lpDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET |
			DISPATCH_METHOD, &dispParams, &va, NULL, NULL);
		if(SUCCEEDED(hr))
		{
			if (!GetCtlStopProperty(lpDisp))
				return -1;

			switch (va.vt)
			{
				case VT_UI1:
					return (short)va.bVal;
					break;

				case VT_I2:
					return (short)va.iVal;
					break;

				case VT_I4:
					return (short)va.lVal;
					break;

				default:
					return -1;
			}
		}
		else
			return -1;
	}
	else
	{
		// Have no "CtlIndex" property, use "tag" value instead.

		lpName[0] = (WCHAR *)L"Tag";
		HRESULT hr = lpDisp->GetIDsOfNames(IID_NULL, lpName, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
		if(SUCCEEDED(hr))
		{
			hr = lpDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET |
				DISPATCH_METHOD, &dispParams, &va, NULL, NULL);
			if(SUCCEEDED(hr))
			{
				if (!GetTabStopProperty(lpDisp))
					return -1;
				
				if (va.vt == VT_BSTR)
				{
					CString strVal(va.bstrVal);
					if (!strVal.IsEmpty())
						return (short)atoi(strVal);
				}
				return -1;
			}
			else
				return -1;
		}
		else
			return -1;
	}
}

BOOL CDateEditBoxCtrl::GetCtlStopProperty(LPDISPATCH lpDisp)
/*
Routine Description:
	Get the "CtlStop" property of a control.

Parameters:
	lpDisp		The IDispatch pointer of the control.
	
Return Value:
	The property value.
*/
{
	VARIANT va;         
	VariantInit(&va);
	DISPID dispid;         
	DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
	// Get the names of all the controls present in a VB form.
	LPWSTR lpName[1] = { (WCHAR *)L"Enabled" };
	HRESULT hr = lpDisp->GetIDsOfNames(IID_NULL, lpName, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if(SUCCEEDED(hr)) 
	{
		hr = lpDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET |
			DISPATCH_METHOD, &dispParams, &va, NULL, NULL);
		if(SUCCEEDED(hr))
		{
			if (va.boolVal == 0)
				return va.boolVal;
		}
		else
			return FALSE;
	}
	else
		return FALSE;
	
	lpName[0] = (WCHAR *)L"CtlStop";
	hr = lpDisp->GetIDsOfNames(IID_NULL, lpName, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if(SUCCEEDED(hr)) 
	{
		hr = lpDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET |
			DISPATCH_METHOD, &dispParams, &va, NULL, NULL);
		if(SUCCEEDED(hr))
		{
			return va.boolVal;
		}
		else
			return FALSE;
	}
	else
		return FALSE;
}

BOOL CDateEditBoxCtrl::GetTabStopProperty(LPDISPATCH lpDisp)
/*
Routine Description:
	Get the "TabStop" property value of a control.

Parameters:
	lpDisp		The IDispatch pointer of the control.
	
Return Value:
	The property value.
*/
{
	VARIANT va;         
	VariantInit(&va);
	DISPID dispid;         
	DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
	// Get the names of all the controls present in a VB form.
	LPWSTR lpName[1] = { (WCHAR *)L"Enabled" };
	HRESULT hr = lpDisp->GetIDsOfNames(IID_NULL, lpName, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if(SUCCEEDED(hr)) 
	{
		hr = lpDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET |
			DISPATCH_METHOD, &dispParams, &va, NULL, NULL);
		if(SUCCEEDED(hr))
		{
			if (va.boolVal == 0)
				return va.boolVal;
		}
		else
			return FALSE;
	}
	else
		return FALSE;
	
	lpName[0] = (WCHAR *)L"TabStop";
	hr = lpDisp->GetIDsOfNames(IID_NULL, lpName, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	if(SUCCEEDED(hr)) 
	{
		hr = lpDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET |
			DISPATCH_METHOD, &dispParams, &va, NULL, NULL);
		if(SUCCEEDED(hr))
		{
			return va.boolVal;
		}
		else
			return FALSE;
	}
	else
		return FALSE;
}

LPDISPATCH CDateEditBoxCtrl::GetDispatchFromUnk(LPUNKNOWN lpUnk)
/*
Routine Description:
	Get control IDispatch control from IUnknown pointer.

Parameters:
	lpUnk		The IUnknown pointer.
	
Return Value:
	The result pointer.
*/
{
	LPOLEOBJECT lpObject = NULL;      
	LPOLECONTROLSITE lpTargetSite = NULL;
	LPOLECLIENTSITE lpClientSite = NULL;      
	LPDISPATCH lpDisp;
	
	HRESULT hr = lpUnk->QueryInterface(IID_IOleObject, (LPVOID*)&lpObject);
	if(SUCCEEDED(hr)) 
	{
		// This is an OLE control.
		// Navigate to the Extended Control because Visual Basic 4.0 uses
		// Extended controls.
		hr = lpObject->GetClientSite(&lpClientSite);
		if(SUCCEEDED(hr)) 
		{
			// You have the IOleClientSite interface
			hr = lpClientSite->QueryInterface(IID_IOleControlSite, (LPVOID*)&lpTargetSite);
			if(SUCCEEDED(hr)) 
			{
				// You have the IOleControlSite interface
				// Get the IDispatch for the extended control.
				// Note that Extended controls are optional in the OLE
				// specifications for OLE Control Containers.
				hr = lpTargetSite->GetExtendedControl(&lpDisp);
			}
		}    
	}      
	else 
	{
		// This is either an internal VB control or the
		// VB form itself.         
		hr = lpUnk->QueryInterface(IID_IDispatch, (LPVOID*)&lpDisp);
	}

	if(lpObject)     
		lpObject->Release();
	if(lpTargetSite) 
		lpTargetSite->Release();
	if(lpClientSite) 
		lpClientSite->Release();      
	
	if (SUCCEEDED(hr))
		return lpDisp;
	else
		return NULL;
}

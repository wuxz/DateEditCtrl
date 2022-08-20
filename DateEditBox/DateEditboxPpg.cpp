// EditboxPpg.cpp : Implementation of the CDateEditBoxPropPage property page class.

#include "stdafx.h"
#include "DateEditBox.h"
#include "DateEditBoxPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CDateEditBoxPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CDateEditBoxPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CDateEditBoxPropPage)
	ON_BN_CLICKED(IDC_RADIO1, OnRadio1)
	ON_BN_CLICKED(IDC_RADIO2, OnRadio2)
	ON_BN_CLICKED(IDC_RADIO3, OnRadio3)
	ON_BN_CLICKED(IDC_RADIO4, OnRadio4)
	ON_BN_CLICKED(IDC_RADIO5, OnRadio5)
	ON_BN_CLICKED(IDC_RADIO6, OnRadio6)
	ON_BN_CLICKED(IDC_RADIO7, OnRadio7)
	ON_BN_CLICKED(IDC_RADIO8, OnRadio8)
	ON_BN_CLICKED(IDC_RADIO9, OnRadio9)
	ON_BN_CLICKED(IDC_RADIO10, OnRadio10)
	ON_BN_CLICKED(IDC_RADIO11, OnRadio11)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CDateEditBoxPropPage, "DATEEDITBOX.EditboxPropPage.1",
	0xf6d0e408, 0x2d34, 0x11d2, 0x9a, 0x46, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxPropPage::CDateEditBoxPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CDateEditBoxPropPage

BOOL CDateEditBoxPropPage::CDateEditBoxPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_DATEEDITBOX_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxPropPage::CDateEditBoxPropPage - Constructor

CDateEditBoxPropPage::CDateEditBoxPropPage() :
	COlePropertyPage(IDD, IDS_DATEEDITBOX_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CDateEditBoxPropPage)
	m_DateFormat = _T("");
	m_Style = -1;
	m_DisplayDateFormat = _T("");
	m_bReadOnly = FALSE;
	m_TextAlign = -1;
	//}}AFX_DATA_INIT
	strFormat[0] = "yyyy/mm/dd";
	strFormat[1] = "yy/mm/dd";
	strFormat[2] = "mm dd'yy";
	strFormat[3] = "mm/dd";
	strFormat[4] = "mm/dd/yyyy";
	strFormat[5] = "dd/mm/yy";
	strFormat[6] = "dd/mm/yyyy";
	strFormat[7] = "mm/yyyy";
	strFormat[8] = "mm/dd/yy";
	strFormat[9] = "ggg/mm/dd";
	strFormat[10] = "ggggggg";
}


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxPropPage::DoDataExchange - Moves data between page and properties

void CDateEditBoxPropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CDateEditBoxPropPage)
	DDP_Text(pDX, IDC_EDIT1, m_DateFormat, _T("DateFormat") );
	DDX_Text(pDX, IDC_EDIT1, m_DateFormat);
	DDP_Radio(pDX, IDC_RADIO1, m_Style, _T("Style") );
	DDX_Radio(pDX, IDC_RADIO1, m_Style);
	DDP_Text(pDX, IDC_EDIT2, m_DisplayDateFormat, _T("DisplayDateFormat") );
	DDX_Text(pDX, IDC_EDIT2, m_DisplayDateFormat);
	DDP_Check(pDX, IDC_CHECK1, m_bReadOnly, _T("ReadOnly") );
	DDX_Check(pDX, IDC_CHECK1, m_bReadOnly);
	DDP_CBIndex(pDX, IDC_COMBO1, m_TextAlign, _T("TextAlign") );
	DDX_CBIndex(pDX, IDC_COMBO1, m_TextAlign);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CDateEditBoxPropPage message handlers

void CDateEditBoxPropPage::OnRadio1() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[0]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[0]);
}

void CDateEditBoxPropPage::OnRadio2() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[1]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[1]);
}

void CDateEditBoxPropPage::OnRadio3() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[2]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[2]);
}

void CDateEditBoxPropPage::OnRadio4() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[3]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[3]);
}

void CDateEditBoxPropPage::OnRadio5() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[4]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[4]);
}

void CDateEditBoxPropPage::OnRadio6() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[5]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[5]);
}

void CDateEditBoxPropPage::OnRadio7() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[6]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[6]);
}

void CDateEditBoxPropPage::OnRadio8() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[7]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[7]);
}

void CDateEditBoxPropPage::OnRadio9() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[8]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[8]);
}

void CDateEditBoxPropPage::OnRadio10() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[9]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[9]);
}

void CDateEditBoxPropPage::OnRadio11() 
{
	CWnd *wnd = GetDlgItem(IDC_EDIT1);
	if (wnd)
		wnd->SetWindowText(strFormat[10]);
	wnd = GetDlgItem(IDC_EDIT2);
	if (wnd)
		wnd->SetWindowText(strFormat[10]);	
}

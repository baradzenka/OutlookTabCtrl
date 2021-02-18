#include "stdafx.h"
#include "MainDialog.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 
Dialog::Dialog(CWnd *parent)
	: CDialog(IDD_MAIN_DIALOG,parent)
{
}

BEGIN_MESSAGE_MAP(Dialog, CDialog)
	ON_WM_CREATE()
	ON_WM_DESTROY()
	ON_WM_SIZE()
	ON_COMMAND(ID_CONTEXT_MENU_POP, OnContextMenuPop)
	ON_COMMAND(ID_CONTEXT_MENU_PUSH, OnContextMenuPush)
	ON_COMMAND(ID_CONTEXT_MENU_MANAGE, OnContextMenuManage)
	ON_BN_CLICKED(IDC_RADIO11,OnBnClickedRadio11)
	ON_BN_CLICKED(IDC_RADIO12,OnBnClickedRadio12)
	ON_BN_CLICKED(IDC_RADIO13,OnBnClickedRadio13)
	ON_BN_CLICKED(IDC_RADIO14,OnBnClickedRadio14)
	ON_BN_CLICKED(IDC_RADIO21,OnBnClickedRadio21)
	ON_BN_CLICKED(IDC_RADIO22,OnBnClickedRadio22)
	ON_BN_CLICKED(IDC_RADIO23,OnBnClickedRadio23)
	ON_BN_CLICKED(IDC_RADIO31,OnBnClickedRadio31)
	ON_BN_CLICKED(IDC_RADIO32,OnBnClickedRadio32)
	ON_BN_CLICKED(IDC_RADIO33,OnBnClickedRadio33)
	ON_BN_CLICKED(IDC_RADIO34,OnBnClickedRadio34)
	ON_BN_CLICKED(IDC_RADIO41,OnBnClickedRadio41)
	ON_BN_CLICKED(IDC_RADIO42,OnBnClickedRadio42)
	ON_BN_CLICKED(IDC_CHECK11,OnBnClickedCheck11)
	ON_BN_CLICKED(IDC_CHECK12,OnBnClickedCheck12)
	ON_BN_CLICKED(IDC_CHECK21,OnBnClickedCheck21)
	ON_BN_CLICKED(IDC_CHECK22,OnBnClickedCheck22)
	ON_BN_CLICKED(IDC_CHECK23,OnBnClickedCheck23)
	ON_BN_CLICKED(IDC_CHECK24,OnBnClickedCheck24)
END_MESSAGE_MAP()


int Dialog::OnCreate(LPCREATESTRUCT lpCreateStruct)
{	if(CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;
		// 
	if( !m_OutlookTabCtrl.Create(this,WS_CHILD | WS_VISIBLE,CRect(0,0,0,0),100) )
		return -1;
	m_Style2.Install(&m_OutlookTabCtrl);
		// 
	m_OutlookTabCtrl.SetAbilityManager(this);
	m_OutlookTabCtrl.SetNotifyManager(this);
		// 
	m_OutlookTabCtrl.CreateStripeImageList(NULL,IDB_STRIPE_NORMAL,IDB_STRIPE_DISABLE,true,24);
	m_OutlookTabCtrl.CreateButtonImageList(NULL,IDB_BUTTON_NORMAL,IDB_BUTTON_DISABLE,true,16);
		// 
	m_OutlookTabCtrl.SetCaptionFont( GetTahomaBoldFont() );
	m_OutlookTabCtrl.SetStripesFont( GetTahomaBoldFont() );
	m_OutlookTabCtrl.SetButtonsFont( GetTahomaFont() );
		// 
	m_OutlookTabCtrl.SetCursor(IDC_CURSOR1);
		// 
	if( !m_List1.Create(WS_CHILD | WS_CLIPCHILDREN | LVS_REPORT,CRect(0,0,0,0),&m_OutlookTabCtrl,100) ||
		!m_List2.Create(WS_CHILD | WS_CLIPCHILDREN | LVS_REPORT,CRect(0,0,0,0),&m_OutlookTabCtrl,101) ||
		!m_List3.Create(WS_CHILD | WS_CLIPCHILDREN | LVS_REPORT,CRect(0,0,0,0),&m_OutlookTabCtrl,102) ||
		!m_List4.Create(WS_CHILD | WS_CLIPCHILDREN | LVS_REPORT,CRect(0,0,0,0),&m_OutlookTabCtrl,103) ||
		!m_List5.Create(WS_CHILD | WS_CLIPCHILDREN | LVS_REPORT,CRect(0,0,0,0),&m_OutlookTabCtrl,104) ||
		!m_List6.Create(WS_CHILD | WS_CLIPCHILDREN | LVS_REPORT,CRect(0,0,0,0),&m_OutlookTabCtrl,105) ||
		!m_List7.Create(WS_CHILD | WS_CLIPCHILDREN | LVS_REPORT,CRect(0,0,0,0),&m_OutlookTabCtrl,106) ||
		!m_List8.Create(WS_CHILD | WS_CLIPCHILDREN | LVS_REPORT,CRect(0,0,0,0),&m_OutlookTabCtrl,107) )
		return -1;
		// 
	m_List1.InsertColumn(0,_T("Mail"),LVCFMT_LEFT,100);
	m_List2.InsertColumn(0,_T("Calendar"),LVCFMT_LEFT,100);
	m_List3.InsertColumn(0,_T("Contacts"),LVCFMT_LEFT,100);
	m_List4.InsertColumn(0,_T("Tasks"),LVCFMT_LEFT,100);
	m_List5.InsertColumn(0,_T("Business Affairs"),LVCFMT_LEFT,100);
	m_List6.InsertColumn(0,_T("Notes"),LVCFMT_LEFT,100);
	m_List7.InsertColumn(0,_T("Folder List"),LVCFMT_LEFT,100);
	m_List8.InsertColumn(0,_T("Shortcuts"),LVCFMT_LEFT,100);
		// 
	m_List1.InsertItem(0,_T("11111111111"));
	m_List1.InsertItem(0,_T("22222222222"));
	m_List1.InsertItem(0,_T("33333333333"));
		// 
	if( !m_OutlookTabCtrl.AddItem(m_List1,_T("Mail"),0,0) ||
		!m_OutlookTabCtrl.AddItem(m_List2,_T("Calendar"),1,1) ||
		!m_OutlookTabCtrl.AddItem(m_List3,_T("Contacts"),2,2) ||
		!m_OutlookTabCtrl.AddItem(m_List4,_T("Tasks"),3,3) ||
		!m_OutlookTabCtrl.AddItem(m_List5,_T("Business Affairs"),-1,-1) ||
		!m_OutlookTabCtrl.AddItem(m_List6,_T("Notes"),4,4) ||
		!m_OutlookTabCtrl.AddItem(m_List7,_T("Folder List"),5,5) ||
		!m_OutlookTabCtrl.AddItem(m_List8,_T("Shortcuts"),6,6) )
		return -1;
		// 
	HANDLE itemContact = m_OutlookTabCtrl.GetItemByIndex(2);   // 'Contacts' item.
	m_OutlookTabCtrl.DisableItem(itemContact,true);   // just for demonstration the disable item.
		// 
	if( !m_OutlookTabCtrl.LoadState(AfxGetApp(),_T("OutlookTabCtrl"),_T("State")) )
	{	m_OutlookTabCtrl.PushVisibleItem();
		m_OutlookTabCtrl.PushVisibleItem();
	}
	m_OutlookTabCtrl.Update();
		// 
	return 0;
}
// 
void Dialog::OnDestroy()
{	m_OutlookTabCtrl.SaveState( AfxGetApp(), _T("OutlookTabCtrl"),_T("State"));
		// 
	CDialog::OnDestroy();
}
/////////////////////////////////////////////////////////////////////////////
// 
BOOL Dialog::OnInitDialog()
{	CDialog::OnInitDialog();
		// 
	GetDlgItem(IDC_OUTLOOKTABCTRL_BASE)->GetWindowRect(&m_rcInit/*out*/);
	ScreenToClient(&m_rcInit);
	SetOutlookTabCtrlPos();
	GetDlgItem(IDC_OUTLOOKTABCTRL_BASE)->ShowWindow(SW_HIDE);
		//
	SetButtonCheck(IDC_RADIO12,true);
	SetButtonCheck(IDC_RADIO22,true);
	SetButtonCheck(IDC_RADIO31,true);
	SetButtonCheck(IDC_RADIO41,true);
	SetButtonCheck(IDC_CHECK11,true);
	SetButtonCheck(IDC_CHECK12,true);
	SetButtonCheck(IDC_CHECK21,true);
	SetButtonCheck(IDC_CHECK22,true);
	SetButtonCheck(IDC_CHECK23,true);
	m_OutlookTabCtrl.HideEmptyButtonsArea(true);
	EnableControl(IDC_CHECK24,false);
	SetButtonCheck(IDC_CHECK24,true);
		// 
	return TRUE;
}
/////////////////////////////////////////////////////////////////////////////
// 
void Dialog::OnSize(UINT nType,int cx,int cy)
{	SetOutlookTabCtrlPos();
		// 
	CDialog::OnSize(nType,cx,cy);
}
/////////////////////////////////////////////////////////////////////////////
// 
LOGFONT const *Dialog::GetTahomaFont() const
{	static const LOGFONT lf = {-11/*8pt*/,0,0,0,FW_NORMAL,FALSE,FALSE,FALSE,DEFAULT_CHARSET,OUT_TT_PRECIS,
		CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,VARIABLE_PITCH | FF_DONTCARE,_T("Tahoma")};
	return &lf;
}
LOGFONT const *Dialog::GetTahomaBoldFont() const
{	static const LOGFONT lf = {-11/*8pt*/,0,0,0,FW_BOLD,FALSE,FALSE,FALSE,DEFAULT_CHARSET,OUT_TT_PRECIS,
		CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,VARIABLE_PITCH | FF_DONTCARE,_T("Tahoma")};
	return &lf;
}
/////////////////////////////////////////////////////////////////////////////
// 
void Dialog::SetOutlookTabCtrlPos()
{	CRect rc;
	GetClientRect(&rc);
	rc.DeflateRect(m_rcInit.left,m_rcInit.top,m_rcInit.top,m_rcInit.top);
	m_OutlookTabCtrl.MoveWindow(&rc);
}
/////////////////////////////////////////////////////////////////////////////
// 
void Dialog::SetButtonCheck(int id, bool check) const 
{	reinterpret_cast<CButton*>( GetDlgItem(id) )->SetCheck(check ? BST_CHECKED : BST_UNCHECKED);
}
bool Dialog::GetButtonCheck(int id) const
{	return reinterpret_cast<CButton*>( GetDlgItem(id) )->GetCheck() == BST_CHECKED;
}
// 
void Dialog::EnableControl(int id, bool enable) const
{	CWnd *wnd = GetDlgItem(id);
	enable ? wnd->ModifyStyle(WS_DISABLED,0) : wnd->ModifyStyle(0,WS_DISABLED);
	wnd->Invalidate();
}
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
//
bool Dialog::CanSelect(OutlookTabCtrl const *ctrl, HANDLE item) const
{	const CString text = ctrl->GetItemText(item);
	if(text=="Calendar")   // just for demonstration using of OutlookTabCtrl::Ability.
	{	::MessageBox(m_hWnd,_T("You can not select this item."),_T("OutlookTabCtrl::Ability"),MB_OK);
		return false;
	}
	return true;
}
/////////////////////////////////////////////////////////////////////////////
//
void Dialog::OnSelectionChanged(OutlookTabCtrl *ctrl)
{	CString text = ctrl->GetItemText( ctrl->GetSelectedItem() );
	if(text==_T("Tasks") || text==_T("Notes"))   // just for demonstration using of OutlookTabCtrl::Notify.
	{	text = _T("You selected item: \"") + text + _T('\"');
		::MessageBox(m_hWnd,text,_T("OutlookTabCtrl::Notify"),MB_OK);
	}
}
/////////////////////////////////////////////////////////////////////////////
//
void Dialog::OnRightButtonReleased(OutlookTabCtrl *ctrl, CPoint pt)
{	CMenu menu;
	menu.LoadMenu(IDR_CONTEXT_MENU);
	CMenu *popup = menu.GetSubMenu(0);
		// 
	if( !ctrl->CanVisibleItemPop() )
		popup->EnableMenuItem(ID_CONTEXT_MENU_POP,MF_BYCOMMAND | MF_GRAYED);
	if( !ctrl->CanVisibleItemPush() )
		popup->EnableMenuItem(ID_CONTEXT_MENU_PUSH,MF_BYCOMMAND | MF_GRAYED);
		// 
	ctrl->ClientToScreen(&pt);
	popup->TrackPopupMenu(TPM_LEFTALIGN,pt.x,pt.y,this);
}
/////////////////////////////////////////////////////////////////////////////
//
void Dialog::OnMenuButtonClicked(OutlookTabCtrl *ctrl, CRect const *rect)
{	OnRightButtonReleased(ctrl, CPoint(rect->right,rect->top) );
}
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
//
void Dialog::OnContextMenuPop()
{	m_OutlookTabCtrl.PopVisibleItem();
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
//
void Dialog::OnContextMenuPush()
{	m_OutlookTabCtrl.PushVisibleItem();
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
//
void Dialog::OnContextMenuManage()
{	if( ManageDialog(&m_OutlookTabCtrl).DoModal()==IDOK )
		m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
// Style: style1.
// 
void Dialog::OnBnClickedRadio11()
{	m_Style1.Install(&m_OutlookTabCtrl);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Style: style2.
// 
void Dialog::OnBnClickedRadio12()
{	m_Style2.Install(&m_OutlookTabCtrl);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Style: style3.
// 
void Dialog::OnBnClickedRadio13()
{	m_Style3.Install(&m_OutlookTabCtrl);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Style: style4.
// 
void Dialog::OnBnClickedRadio14()
{	m_Style4.Install(&m_OutlookTabCtrl);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
// Caption: none.
// 
void Dialog::OnBnClickedRadio21()
{	m_OutlookTabCtrl.ShowCaption(false);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Caption: top.
// 
void Dialog::OnBnClickedRadio22()
{	if( !m_OutlookTabCtrl.IsCaptionVisible() )
		m_OutlookTabCtrl.ShowCaption(true);
	m_OutlookTabCtrl.SetCaptionAlign(OutlookTabCtrl::CaptionAlignTop);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Caption: bottom.
// 
void Dialog::OnBnClickedRadio23()
{	if( !m_OutlookTabCtrl.IsCaptionVisible() )
		m_OutlookTabCtrl.ShowCaption(true);
	m_OutlookTabCtrl.SetCaptionAlign(OutlookTabCtrl::CaptionAlignBottom);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
// Layout: controls|splitter|tabs|buttons.
// 
void Dialog::OnBnClickedRadio31()
{	m_OutlookTabCtrl.SetLayout(OutlookTabCtrl::Layout1);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Layout: controls|splitter|buttons|tabs.
// 
void Dialog::OnBnClickedRadio32()
{	m_OutlookTabCtrl.SetLayout(OutlookTabCtrl::Layout2);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Layout: buttons|tabs|splitter|controls.
// 
void Dialog::OnBnClickedRadio33()
{	m_OutlookTabCtrl.SetLayout(OutlookTabCtrl::Layout3);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Layout: tabs|buttons|splitter|controls.
// 
void Dialog::OnBnClickedRadio34()
{	m_OutlookTabCtrl.SetLayout(OutlookTabCtrl::Layout4);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
// Buttons align: right.
// 
void Dialog::OnBnClickedRadio41()
{	m_OutlookTabCtrl.SetButtonsAlign(OutlookTabCtrl::ButtonsAlignRight);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Buttons align: left.
// 
void Dialog::OnBnClickedRadio42()
{	m_OutlookTabCtrl.SetButtonsAlign(OutlookTabCtrl::ButtonsAlignLeft);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
// Show border.
// 
void Dialog::OnBnClickedCheck11()
{	m_OutlookTabCtrl.ShowBorder( GetButtonCheck(IDC_CHECK11) );
	m_OutlookTabCtrl.Update();
	m_OutlookTabCtrl.SetWindowPos(NULL, 0,0,0,0, SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED);   // border update.
}
/////////////////////////////////////////////////////////////////////////////
// Active splitter.
// 
void Dialog::OnBnClickedCheck12()
{	m_OutlookTabCtrl.ActivateSplitter( GetButtonCheck(IDC_CHECK12) );
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
// Show button text.
// 
void Dialog::OnBnClickedCheck21()
{	m_OutlookTabCtrl.ShowButtonText( GetButtonCheck(IDC_CHECK21) );
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Show button separators.
// 
void Dialog::OnBnClickedCheck22()
{	m_OutlookTabCtrl.ShowButtonSeparator( GetButtonCheck(IDC_CHECK22) );
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Show menu button.
// 
void Dialog::OnBnClickedCheck23()
{	const bool show = GetButtonCheck(IDC_CHECK23);
	EnableControl(IDC_CHECK24,!show);
	m_OutlookTabCtrl.ShowMenuButton(show);
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
// Hide empty button bar.
// 
void Dialog::OnBnClickedCheck24()
{	m_OutlookTabCtrl.HideEmptyButtonsArea( GetButtonCheck(IDC_CHECK24) );
	m_OutlookTabCtrl.Update();
}
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////






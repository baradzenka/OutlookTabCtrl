/////////////////////////////////////////////////////////////////////////////
#include "stdafx.h"
#include "ManageDialog.h"
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
BEGIN_MESSAGE_MAP(ManageDialog, CDialog)
	ON_BN_CLICKED(ID_MANAGE_DIALOG_UP, OnBnClickedUp)
	ON_BN_CLICKED(ID_MANAGE_DIALOG_DOWN, OnBnClickedDown)
	ON_BN_CLICKED(ID_MANAGE_DIALOG_RESET, OnBnClickedReset)
END_MESSAGE_MAP()
/////////////////////////////////////////////////////////////////////////////
// 
ManageDialog::ManageDialog(OutlookTabCtrl *ctrl) : 
	CDialog(IDD_MANAGE_DIALOG)
{
	m_pTabCtrl = ctrl;
}
/////////////////////////////////////////////////////////////////////////////
// 
BOOL ManageDialog::OnInitDialog()
{	CDialog::OnInitDialog();
		// 
	m_ListCtrl.SubclassDlgItem(ID_MANAGE_DIALOG_LIST,this);
		// 
	CSize szTab, szButton;
	m_pTabCtrl->GetImageSize(&szTab/*out*/,&szButton/*out*/);
	m_bTabImages = (szTab.cy <= szButton.cy);
	CImageList normal, disable;
	(m_bTabImages ?
		m_pTabCtrl->GetStripeImageList( ::GetSysColor(COLOR_WINDOW), &normal/*out*/,&disable/*out*/) :
		m_pTabCtrl->GetButtonImageList( ::GetSysColor(COLOR_WINDOW), &normal/*out*/,&disable/*out*/));
	m_ImageList.Create(normal.GetSafeHandle() ? &normal : &disable);
		// 
	CRect rc;
	m_ListCtrl.GetClientRect(&rc);
	m_ListCtrl.InsertColumn(0,_T(""),LVCFMT_LEFT,rc.Width());
	m_ListCtrl.GetHeaderCtrl()->ModifyStyle(0,HDS_HIDDEN);
	m_ListCtrl.SetExtendedStyle(LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
	m_ListCtrl.SetImageList(&m_ImageList,LVSIL_SMALL);
	FillListCtrl();
		// 
	return TRUE;
}
/////////////////////////////////////////////////////////////////////////////
// 
void ManageDialog::FillListCtrl()
{	m_ListCtrl.DeleteAllItems();
		// 
	const int number = m_pTabCtrl->GetTotalNumberItems();
	for(int i=0; i<number; ++i)
	{	HANDLE h = m_pTabCtrl->GetItemHandleByIndex(i);
		int image;
		(m_bTabImages ? m_pTabCtrl->GetItemImage(h,&image/*out*/,NULL) : m_pTabCtrl->GetItemImage(h,NULL,&image/*out*/));
		const CString text = m_pTabCtrl->GetItemText(h);
		const bool visible = m_pTabCtrl->IsItemVisible(h);
		m_ListCtrl.InsertItem(i,text,image);
		m_ListCtrl.SetItemData(i,reinterpret_cast<DWORD_PTR>(h));
		m_ListCtrl.SetCheck(i,visible);	
	}
	if(number)
	{	const int idx = m_pTabCtrl->GetItemIndexByHandle( m_pTabCtrl->GetSelectedItem() );
		m_ListCtrl.SetItemState(idx,LVIS_SELECTED,LVIS_SELECTED);
	}
}
/////////////////////////////////////////////////////////////////////////////
// Up.
// 
void ManageDialog::OnBnClickedUp()
{	const int src = GetListCtrlItemSelection();
	const int dst = src-1;
		// 
	if(dst>=0)
	{	const CString text = m_ListCtrl.GetItemText(src,0);
		const int image = GetListCtrlItemImage(src);
		const bool check = m_ListCtrl.GetCheck(src)!=0;
		const DWORD_PTR data = m_ListCtrl.GetItemData(src);
			// 
		m_ListCtrl.DeleteItem(src);
		m_ListCtrl.InsertItem(dst,text,image);
		m_ListCtrl.SetCheck(dst,check);
		m_ListCtrl.SetItemData(dst,data);
		m_ListCtrl.SetItemState(dst,LVIS_SELECTED,LVIS_SELECTED);
	}
}
/////////////////////////////////////////////////////////////////////////////
// Down.
// 
void ManageDialog::OnBnClickedDown()
{	const int src = GetListCtrlItemSelection();
	if(src!=-1)
	{	const int dst = src+1;
			// 
		if(dst < m_ListCtrl.GetItemCount())
		{	const CString text = m_ListCtrl.GetItemText(src,0);
			const int image = GetListCtrlItemImage(src);
			const bool check = m_ListCtrl.GetCheck(src)!=0;
			const DWORD_PTR data = m_ListCtrl.GetItemData(src);
				// 
			m_ListCtrl.DeleteItem(src);
			m_ListCtrl.InsertItem(dst,text,image);
			m_ListCtrl.SetCheck(dst,check);
			m_ListCtrl.SetItemData(dst,data);
			m_ListCtrl.SetItemState(dst,LVIS_SELECTED,LVIS_SELECTED);
		}
	}
}
/////////////////////////////////////////////////////////////////////////////
// Reset.
// 
void ManageDialog::OnBnClickedReset()
{	FillListCtrl();
}
/////////////////////////////////////////////////////////////////////////////
// 
int ManageDialog::GetListCtrlItemSelection()
{	POSITION pos = m_ListCtrl.GetFirstSelectedItemPosition();
	return (pos ? m_ListCtrl.GetNextSelectedItem(pos) : -1);
}
/////////////////////////////////////////////////////////////////////////////
// 
int ManageDialog::GetListCtrlItemImage(int idx) const
{	LVITEM lv;
	lv.iItem = idx;
	lv.iSubItem = 0;
	lv.mask = LVIF_IMAGE;
	m_ListCtrl.GetItem(&lv/*out*/);
	return lv.iImage;
}
/////////////////////////////////////////////////////////////////////////////
// 
void ManageDialog::OnOK()
{	std::vector<HANDLE> items;
	typedef std::pair<HANDLE,bool> itemAndVisibl;
	std::vector<itemAndVisibl> itemsVisibl;
		// 
	for(int i=0; i<m_ListCtrl.GetItemCount(); ++i)
	{	HANDLE h = reinterpret_cast<HANDLE>( m_ListCtrl.GetItemData(i) );
		const bool check = m_ListCtrl.GetCheck(i)!=0;
		items.push_back(h);
		itemsVisibl.push_back( itemAndVisibl(h,check) );
	}
		// 
	m_pTabCtrl->SetItemsOrder(items);
		// 
	for(std::vector<itemAndVisibl>::const_iterator i=itemsVisibl.begin(), e=itemsVisibl.end(); i!=e; ++i)
		m_pTabCtrl->ShowItem(i->first,i->second);
	m_pTabCtrl->Update();
		// 
	CDialog::OnOK();
}
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////






























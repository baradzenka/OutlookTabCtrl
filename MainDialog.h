#pragma once

#include "ManageDialog.h"


// 
class Dialog : public CDialog,
	public OutlookTabCtrl::Ability,
	public OutlookTabCtrl::Notify
{
public:
	Dialog(CWnd *parent);

private:   // OutlookTabCtrl::Ability.
	bool CanSelect(OutlookTabCtrl const *ctrl, HANDLE item) override;

private:   // OutlookTabCtrl::Notify.
	void OnSelectionChanged(OutlookTabCtrl *ctrl) override;
	void OnRightButtonReleased(OutlookTabCtrl *ctrl, CPoint pt) override;
	void OnMenuButtonClicked(OutlookTabCtrl *ctrl, CRect const *rect) override;

private:
	OutlookTabCtrl m_OutlookTabCtrl;
	OutlookTabCtrlCustom1 m_Style1;
	OutlookTabCtrlCustom2 m_Style2;
	OutlookTabCtrlCustom3 m_Style3;
	OutlookTabCtrlCustom4 m_Style4;
	CListCtrl m_List1,m_List2,m_List3,m_List4,m_List5,m_List6,m_List7,m_List8;

private:
	LOGFONT const *GetTahomaFont() const;
	LOGFONT const *GetTahomaBoldFont() const;
	void SetOutlookTabCtrlPos();
	void SetButtonCheck(int id, bool check) const;
	bool GetButtonCheck(int id) const;
	void EnableControl(int id, bool enable) const;

protected:
	DECLARE_MESSAGE_MAP()
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	BOOL OnInitDialog() override;
	afx_msg void OnDestroy();
	afx_msg void OnSize(UINT nType,int cx,int cy);
	afx_msg void OnContextMenuPop();
	afx_msg void OnContextMenuPush();
	afx_msg void OnContextMenuManage();
	afx_msg void OnBnClickedRadio11();
	afx_msg void OnBnClickedRadio12();
	afx_msg void OnBnClickedRadio13();
	afx_msg void OnBnClickedRadio14();
	afx_msg void OnBnClickedRadio21();
	afx_msg void OnBnClickedRadio22();
	afx_msg void OnBnClickedRadio23();
	afx_msg void OnBnClickedRadio31();
	afx_msg void OnBnClickedRadio32();
	afx_msg void OnBnClickedRadio33();
	afx_msg void OnBnClickedRadio34();
	afx_msg void OnBnClickedRadio41();
	afx_msg void OnBnClickedRadio42();
	afx_msg void OnBnClickedCheck11();
	afx_msg void OnBnClickedCheck12();
	afx_msg void OnBnClickedCheck21();
	afx_msg void OnBnClickedCheck22();
	afx_msg void OnBnClickedCheck23();
	afx_msg void OnBnClickedCheck24();
};

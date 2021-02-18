//==========================================================
// Author: Baradzenka Aleh (baradzenka@gmail.com)
//==========================================================
// 
#pragma once
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
#pragma warning(push)
#pragma warning(disable : 4458)   // declaration of 'nativeCap' hides class member.
#include <gdiplus.h>
#pragma warning(pop)
// 
#include <vector>
// 
#if (!defined(_MSC_VER) && __cplusplus < 201103L) || (defined(_MSC_VER) && _MSC_VER < 1900)   // C++11 is not supported.
	#define override
#endif
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
// 
class OutlookTabCtrl : public CWnd
{	DECLARE_DYNCREATE(OutlookTabCtrl)

///////////////////////////////////////
// PUBLIC
///////////////////////////////////////
public:
	struct Ability
	{	virtual bool CanSelect(OutlookTabCtrl const * /*ctrl*/, HANDLE /*item*/) const { return true; }
	};
	struct Notify
	{	virtual void OnSelectionChanged(OutlookTabCtrl * /*ctrl*/) {}
		virtual void OnRightButtonReleased(OutlookTabCtrl * /*ctrl*/, CPoint /*pt*/) {}   // for context menu.
		virtual void OnMenuButtonClicked(OutlookTabCtrl * /*ctrl*/, CRect const * /*rect*/) {}
	};
	struct ToolTip
	{	virtual CToolTipCtrl *CreateToolTip(OutlookTabCtrl * /*ctrl*/) { return NULL; }
		virtual void DestroyToolTip(CToolTipCtrl * /*tooltip*/) {}
		virtual bool HasButtonTooltip(OutlookTabCtrl const * /*ctrl*/, HANDLE /*item*/) const { return true; }
	};
	struct Draw
	{	virtual void DrawBegin(OutlookTabCtrl const * /*ctrl*/, CDC * /*dc*/) const {}
		virtual void DrawBorder(OutlookTabCtrl const * /*ctrl*/, CDC * /*dc*/, CRect const * /*rect*/) const {}
		virtual void DrawCaption(OutlookTabCtrl const * /*ctrl*/, CDC * /*dc*/, CRect const * /*rect*/) const {}
		virtual void DrawEmptyWindowsArea(OutlookTabCtrl const * /*ctrl*/, CDC * /*dc*/, CRect const * /*rect*/) const {}
		virtual void DrawSplitter(OutlookTabCtrl const * /*ctrl*/, CDC * /*dc*/, CRect const * /*rect*/) const {}
		virtual void DrawStripe(OutlookTabCtrl const * /*ctrl*/, CDC * /*dc*/, HANDLE /*item*/, bool /*drawSeparator*/) const {}
		virtual void DrawButton(OutlookTabCtrl const * /*ctrl*/, CDC * /*dc*/, HANDLE /*item*/) const {}
		virtual void DrawButtonsBackground(OutlookTabCtrl const * /*ctrl*/, CDC * /*dc*/, CRect const * /*rect*/) const {}
		virtual void DrawButtonMenu(OutlookTabCtrl const * /*ctrl*/, CDC * /*dc*/, CRect const * /*rect*/) const {}
		virtual void DrawEnd(OutlookTabCtrl const * /*ctrl*/, CDC * /*dc*/) const {}
	};
	interface IRecalc
	{	virtual int GetBorderWidth(OutlookTabCtrl const *ctrl, IRecalc const *base) const = 0;
		virtual int GetCaptionHeight(OutlookTabCtrl const *ctrl, IRecalc const *base) const = 0;
		virtual int GetSplitterHeight(OutlookTabCtrl const *ctrl, IRecalc const *base) const = 0;
		virtual int GetStripeHeight(OutlookTabCtrl const *ctrl, IRecalc const *base) const = 0;
		virtual int GetButtonHeight(OutlookTabCtrl const *ctrl, IRecalc const *base) const = 0;
		virtual int GetMinButtonWidth(OutlookTabCtrl const *ctrl, IRecalc const *base) const = 0;
		virtual int GetMenuButtonWidth(OutlookTabCtrl const *ctrl, IRecalc const *base) const = 0;
	};

public:
	OutlookTabCtrl();
	~OutlookTabCtrl();

public:
	void Update();
		// 
	bool Create(CWnd *parent, DWORD style, RECT const &rect, UINT id);
	HANDLE AddItem(HWND wnd, TCHAR const *text, int imageStripe, int imageButton);
	HANDLE InsertItem(HANDLE before, HWND wnd, TCHAR const *text, int imageStripe, int imageButton);
	void RemoveItem(HANDLE before, HANDLE src);
	void DeleteItem(HANDLE item);
	void DeleteAllItems();
		// 
	void SetAbilityManager(Ability *p/*or null*/);
	Ability *GetAbilityManager() const;
	void SetNotifyManager(Notify *p/*or null*/);
	Notify *GetNotifyManager() const;
	void SetToolTipManager(ToolTip *p/*or null*/);
	ToolTip *GetToolTipManager() const;
	void SetDrawManager(Draw *p/*or null*/);
	Draw *GetDrawManager() const;
	void SetRecalcManager(IRecalc *p/*or null*/);   // NULL for default manager.
	IRecalc *GetRecalcManager() const;
		// 
	enum CaptionAlign
	{	CaptionAlignTop, CaptionAlignBottom
	};
	void SetCaptionAlign(CaptionAlign align);
	CaptionAlign GetCaptionAlign() const;
		// 
	enum Layout  // from top to bottom.
	{	Layout1,   // controls|splitter|stripes|buttons.
		Layout2,   // controls|splitter|buttons|stripes.
		Layout3,   // buttons|stripes|splitter|controls.
		Layout4   // stripes|buttons|splitter|controls.
	};
	void SetLayout(Layout layout);
	Layout GetLayout() const;
		// 
	enum ButtonsAlign
	{	ButtonsAlignRight,   // rise from right to left.
		ButtonsAlignLeft   // rise from left to right.
	};
	void SetButtonsAlign(ButtonsAlign align);
	ButtonsAlign GetButtonsAlign() const;
		// 
	bool CreateStripeImageList(HINSTANCE hinstRes/*or null*/, UINT resNormalID/*or 0*/, UINT resDisableID/*or 0*/, bool pngImage, int imageWidth, COLORREF clrTransp=CLR_NONE);
	void GetStripeImages(Gdiplus::Bitmap **normal/*out,or null*/, Gdiplus::Bitmap **disable/*out,or null*/) const;
	bool GetStripeImageLists(COLORREF clrBack/*or CLR_NONE*/, CImageList *normal/*out,or null*/, CImageList *disable/*out,or null*/) const;
	bool CreateButtonImageList(HINSTANCE hinstRes/*or null*/, UINT resNormalID/*or 0*/, UINT resDisableID/*or 0*/, bool pngImage, int imageWidth, COLORREF clrTransp=CLR_NONE);
	void GetButtonImages(Gdiplus::Bitmap **normal/*out,or null*/, Gdiplus::Bitmap **disable/*out,or null*/) const;
	bool GetButtonImageLists(COLORREF clrBack/*or CLR_NONE*/, CImageList *normal/*out,or null*/, CImageList *disable/*out,or null*/) const;
	void GetImageSizes(CSize *szStripe/*out,or null*/, CSize *szButton/*out,or null*/) const;
	void GetImageTranspColor(COLORREF *clrStripe/*out,or null*/, COLORREF *clrButton/*out,or null*/) const;
		// 
	bool SetCaptionFont(CFont *font);
	bool SetCaptionFont(LOGFONT const *logfont);
	CFont *GetCaptionFont();
		// 
	bool SetStripesFont(CFont *font);
	bool SetStripesFont(LOGFONT const *logfont);
	CFont *GetStripesFont();
		// 
	bool SetButtonsFont(CFont *font);
	bool SetButtonsFont(LOGFONT const *logfont);
	CFont *GetButtonsFont();
		// 
	bool SetCursor(UINT drag);
	bool SetCursor(HMODULE module, UINT drag);
	HCURSOR GetCursor() const;
		// 
	void SetItemText(HANDLE item, TCHAR const *text);
	CString GetItemText(HANDLE item) const;
	void SetItemImage(HANDLE item, int imageStripe, int imageButton);
	void GetItemImage(HANDLE item, int *imageStripe/*out,or null*/, int *imageButton/*out,or null*/) const;
	void SetItemWindow(HANDLE item, HWND wnd);
	HWND GetItemWindow(HANDLE item) const;
	void SetItemData(HANDLE item, __int64 data);
	__int64 GetItemData(HANDLE item) const;
		// 
	int GetNumberVisibleStripes() const;
	int GetNumberVisibleButtons() const;
	int GetTotalNumberVisibleItems() const;   // GetNumberVisibleStripes()+GetNumberVisibleButtons().
	int GetNumberStripes() const;
	int GetNumberButtons() const;
	int GetTotalNumberItems() const;   // visible + invisible = GetNumberStripes() + GetNumberButtons().
		// 
	HANDLE GetTopVisibleItem() const;
	HANDLE GetBottomVisibleItem() const;
	HANDLE GetTopVisibleAndEnableItem() const;
		// 
	void SelectItem(HANDLE item);   // select item.
	HANDLE GetSelectedItem() const;	
	void ShowItem(HANDLE item, bool show);   // show/hide item.
	bool IsItemVisible(HANDLE item) const;
	void DisableItem(HANDLE item, bool disable);   // disable/enable item.
	bool IsItemDisabled(HANDLE item) const;
	bool IsItemInStripeState(HANDLE item) const;
	bool IsItemInButtonState(HANDLE item) const;
		//
	struct ItemState
	{	bool selected;
		bool highlighted;
		bool pushed;
	};
	ItemState GetItemState(HANDLE item) const;
	ItemState GetMenuButtonState() const;
	bool IsAnyStripeOrButtonPushed() const;
		// 
	HANDLE HitTest(CPoint point) const;   // get item in the given point.
	HANDLE GetItemByIndex(int idx) const;   // top item has index 0 (from top to bottom).
	int GetIndexByHandle(HANDLE item) const;
	int GetVisibleIndexByHandle(HANDLE item) const;
	bool IsItemExist(HANDLE item) const;
		// 
	RECT GetItemRect(HANDLE item) const;   // stripe or button.
	RECT GetCaptionRect() const;
	RECT GetSplitterRect() const;
	RECT GetStripesAreaRect() const;
	RECT GetButtonsAreaRect() const;
	RECT GetMenuButtonRect() const;
		// 
	bool CanVisibleItemPush() const;
	void PushVisibleItem();   // from stripe to button state (for 1 item).
	bool CanVisibleItemPop() const;
	void PopVisibleItem();   // from button to stripe state (for 1 item).
	void SetNumberVisibleItemsInStripeState(int count);
	void SetItemsOrder(std::vector<HANDLE> const &items);   // set a new order of the current items.
		// 
	void ShowBorder(bool show);
	bool IsBorderVisible() const;
	void ShowCaption(bool show);
	bool IsCaptionVisible() const;
	void ActivateSplitter(bool active);
	bool IsSplitterActive() const;
	void ShowMenuButton(bool show);
	bool IsMenuButtonVisible() const;
	void HideEmptyButtonsArea(bool hide);
	bool IsEmptyButtonsAreaHide() const;
	void ShowButtonText(bool show);
	bool IsButtonTextVisible() const;
	void ShowButtonSeparator(bool show);
	bool IsButtonSeparatorVisible() const;
		// 
	bool IsButtonsAreaVisible() const;
		//
	int GetBorderWidth() const;
	int GetCaptionHeight() const;
	int GetSplitterHeight() const;
	int GetStripeHeight() const;
	int GetButtonHeight() const;
	int GetMinButtonWidth() const;
	int GetMenuButtonWidth() const;
		// 
	bool LoadState(CWinApp *app, TCHAR const *section, TCHAR const *entry);   // load state from registry.
	bool SaveState(CWinApp *app, TCHAR const *section, TCHAR const *entry) const;   // save state in registry.
	bool LoadState(CArchive *ar);
	bool SaveState(CArchive *ar) const;

///////////////////////////////////////
// PRIVATE
///////////////////////////////////////
private:
	struct Private;
	Private &p;

///////////////////////////////////////
// PROTECTED
///////////////////////////////////////
protected:
	DECLARE_MESSAGE_MAP()
	BOOL Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD style, const RECT &rect, CWnd *parentWnd, UINT id, CCreateContext *context = NULL) override;
	BOOL PreTranslateMessage(MSG *pMsg) override;
	afx_msg void OnDestroy();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnPaint();
	afx_msg void OnNcLButtonDown(UINT nHitTest, CPoint point);
	afx_msg void OnNcLButtonDblClk(UINT nHitTest, CPoint point);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags,CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg void OnCaptureChanged(CWnd *pWnd);
	afx_msg BOOL OnSetCursor(CWnd *pWnd, UINT nHitTest, UINT message);
	afx_msg LRESULT OnNcHitTest(CPoint point);
};
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
// 
struct OutlookTabCtrlCustomBase : 
	OutlookTabCtrl::ToolTip,
	OutlookTabCtrl::Draw,
	OutlookTabCtrl::IRecalc
{
	void Install(OutlookTabCtrl *ctrl);

protected:
		// OutlookTabCtrl::ToolTip.
	CToolTipCtrl *CreateToolTip(OutlookTabCtrl *ctrl) override;
	void DestroyToolTip(CToolTipCtrl *tooltip) override;
	bool HasButtonTooltip(OutlookTabCtrl const *ctrl, HANDLE item) const override;

		// OutlookTabCtrl::Draw.
	void DrawBorder(OutlookTabCtrl const *ctrl, CDC *dc, CRect const *rect) const override;
	void DrawCaption(OutlookTabCtrl const *ctrl, CDC *dc, CRect const *rect) const override;
	void DrawEmptyWindowsArea(OutlookTabCtrl const *ctrl, CDC *dc, CRect const *rect) const override;
	void DrawSplitter(OutlookTabCtrl const *ctrl, CDC *dc, CRect const *rect) const override;
	void DrawStripe(OutlookTabCtrl const *ctrl, CDC *dc, HANDLE item, bool drawSeparator) const override;
	void DrawButton(OutlookTabCtrl const *ctrl, CDC * dc, HANDLE item) const override;
	void DrawButtonsBackground(OutlookTabCtrl const *ctrl, CDC *dc, CRect const *rect) const override;
	void DrawButtonMenu(OutlookTabCtrl const *ctrl, CDC *dc, CRect const *rect) const override;

		// OutlookTabCtrl::IRecalc.
	int GetBorderWidth(OutlookTabCtrl const *ctrl, IRecalc const *base) const override;
	int GetCaptionHeight(OutlookTabCtrl const *ctrl, IRecalc const *base) const override;
	int GetSplitterHeight(OutlookTabCtrl const *ctrl, IRecalc const *base) const override;
	int GetStripeHeight(OutlookTabCtrl const *ctrl, IRecalc const *base) const override;
	int GetButtonHeight(OutlookTabCtrl const *ctrl, IRecalc const *base) const override;
	int GetMinButtonWidth(OutlookTabCtrl const *ctrl, IRecalc const *base) const override;
	int GetMenuButtonWidth(OutlookTabCtrl const *ctrl, IRecalc const *base) const override;

	virtual COLORREF GetEmptyWindowsAreaColor() const { return ::GetSysColor(COLOR_WINDOW); }
	virtual COLORREF GetBorderColor() const { return ::GetSysColor(COLOR_BTNSHADOW); }
	virtual COLORREF GetCaptionColor() const { return ::GetSysColor(COLOR_BTNSHADOW); }
	virtual COLORREF GetCaptionTextColor() const { return ::GetSysColor(COLOR_CAPTIONTEXT); }
	virtual COLORREF GetSplitterBackColor() const { return ::GetSysColor(COLOR_BTNFACE); }
	virtual COLORREF GetSplitterDotsColor() const { return ::GetSysColor(COLOR_3DDKSHADOW); }
	virtual COLORREF GetSeparationLineColor() const { return ::GetSysColor(COLOR_BTNSHADOW); }
	virtual COLORREF GetStripeTextColor(OutlookTabCtrl::ItemState const * /*state*/) const { return ::GetSysColor(COLOR_WINDOWTEXT); }
	virtual COLORREF GetDisabledStripeTextColor(OutlookTabCtrl::ItemState const * /*state*/) const { return ::GetSysColor(COLOR_GRAYTEXT); }
	virtual COLORREF GetButtonTextColor(OutlookTabCtrl::ItemState const *state) const { return GetStripeTextColor(state); }
	virtual COLORREF GetDisabledButtonTextColor(OutlookTabCtrl::ItemState const *state) const { return GetDisabledStripeTextColor(state); }
	virtual COLORREF GetMenuButtonImageColor() const { return ::GetSysColor(COLOR_WINDOWTEXT); }

	virtual int GetCaptionTextLeftMargin() const { return 7; }
	virtual int GetStripeContentLeftMargin() const { return 5; }
	virtual int GetStripeImageTextGap() const { return 5; }
	virtual int GetButtonContentLeftMargin() const { return 5; }
	virtual int GetButtonImageTextGap() const { return 5; }

	virtual void GetHighlightStateOfItem(OutlookTabCtrl const *ctrl, OutlookTabCtrl::ItemState const *state, bool *selectLight/*out*/, bool *selectDark/*out*/) const;
	virtual void DrawBackground(OutlookTabCtrl const *ctrl, CDC *dc, OutlookTabCtrl::ItemState const *state, CRect const *rect) const;
	COLORREF MixingColors(COLORREF src, COLORREF dst, int percent) const;
	CSize GetTextSize(CFont *font, CString const &text) const;
	void DrawSplitterDots(CDC *dc, CRect const *rect, int count, int size, COLORREF color) const;
	void DrawImage(CDC *dc, Gdiplus::Bitmap *bmp, CPoint const &ptDst, int image, CSize const &szSrc, COLORREF clrTransp) const;

	static const int m_iMenuImageWidth, m_iMenuImageHeight;
};
/////////////////////////////////////////////////////////////////////////////
// 
class OutlookTabCtrlCustomTooltip : public OutlookTabCtrlCustomBase
{
	#ifdef AFX_TOOLTIP_TYPE_ALL
		CToolTipCtrl *CreateToolTip(OutlookTabCtrl *ctrl) override
		{	CToolTipCtrl *tooltip = NULL;
			return (CTooltipManager::CreateToolTip(tooltip/*out*/,ctrl,AFX_TOOLTIP_TYPE_TAB) ? tooltip : NULL);
		}
		void DestroyToolTip(CToolTipCtrl *tooltip) override
		{	CTooltipManager::DeleteToolTip(tooltip);
		}
	#endif
};
/////////////////////////////////////////////////////////////////////////////
// 
class OutlookTabCtrlCustom2 : public OutlookTabCtrlCustomTooltip
{	void DrawBackground(OutlookTabCtrl const *ctrl, CDC *dc, OutlookTabCtrl::ItemState const *state, CRect const *rect) const override;

	COLORREF GetEmptyWindowsAreaColor() const override { return RGB(255,255,255); }
	COLORREF GetBorderColor() const override { return RGB(83,83,83); }
	COLORREF GetCaptionColor() const override { return RGB(160,160,160); }
	COLORREF GetCaptionTextColor() const override { return RGB(255,255,255); }
	COLORREF GetSplitterBackColor() const override { return RGB(180,180,180); }
	COLORREF GetSplitterDotsColor() const override { return RGB(35,35,35); }
	COLORREF GetSeparationLineColor() const override { return RGB(160,160,160); }
	COLORREF GetStripeTextColor(OutlookTabCtrl::ItemState const * /*state*/) const override { return RGB(0,0,0); }
	COLORREF GetDisabledStripeTextColor(OutlookTabCtrl::ItemState const * /*state*/) const override { return RGB(109,109,109); }
	COLORREF GetMenuButtonImageColor() const override { return RGB(0,0,0); }
};
/////////////////////////////////////////////////////////////////////////////
// 
class OutlookTabCtrlCustom3 : public OutlookTabCtrlCustomTooltip
{		// OutlookTabCtrl::Draw.
	void DrawSplitter(OutlookTabCtrl const *ctrl, CDC *dc, CRect const *rect) const override;

		// OutlookTabCtrl::IRecalc.
	int GetSplitterHeight(OutlookTabCtrl const *ctrl, IRecalc const * /*base*/) const override { return (ctrl->IsSplitterActive() ? 7 : 5); }

	COLORREF GetEmptyWindowsAreaColor() const override { return RGB(255,255,255); }
	COLORREF GetBorderColor() const override { return RGB(77,115,61); }
	COLORREF GetCaptionColor() const override { return RGB(160,160,160); }
	COLORREF GetCaptionTextColor() const override { return RGB(255,255,255); }
	COLORREF GetSeparationLineColor() const override { return RGB(77,115,61); }
	COLORREF GetStripeTextColor(OutlookTabCtrl::ItemState const *state) const override { return (!state->highlighted ? RGB(60,60,60) : RGB(0,0,0)); }
	COLORREF GetDisabledStripeTextColor(OutlookTabCtrl::ItemState const * /*state*/) const override { return RGB(128,128,128); }
	COLORREF GetMenuButtonImageColor() const override { return RGB(77,115,61); }

	void DrawBackground(OutlookTabCtrl const *ctrl, CDC *dc, OutlookTabCtrl::ItemState const *state, CRect const *rect) const override;
	void DrawGradient(CDC *dc, CRect const *rc, COLORREF clrTop, COLORREF clrBottom) const;
	void DrawSplitterDots(CDC *dc, CRect const *rect, int count, int size, int offset, COLORREF clrTopDot, COLORREF clrBottomDot) const;
};
/////////////////////////////////////////////////////////////////////////////
// 
class OutlookTabCtrlCustom4 : public OutlookTabCtrlCustomTooltip
{	void DrawBackground(OutlookTabCtrl const *ctrl, CDC *dc, OutlookTabCtrl::ItemState const *state, CRect const *rect) const override;

	COLORREF GetEmptyWindowsAreaColor() const override { return RGB(255,255,255); }
	COLORREF GetBorderColor() const override { return RGB(41,57,85); }
	COLORREF GetCaptionColor() const override { return RGB(41,57,85); }
	COLORREF GetCaptionTextColor() const override { return RGB(250,250,250); }
	COLORREF GetSplitterBackColor() const override { return RGB(41,57,85); }
	COLORREF GetSplitterDotsColor() const override { return RGB(230,230,230); }
	COLORREF GetSeparationLineColor() const override { return RGB(41,57,85); }
	COLORREF GetStripeTextColor(OutlookTabCtrl::ItemState const * /*state*/) const override { return RGB(255,255,255); }
	COLORREF GetDisabledStripeTextColor(OutlookTabCtrl::ItemState const * /*state*/) const override { return RGB(160,160,160); }
	COLORREF GetMenuButtonImageColor() const override { return RGB(235,235,235); }
};
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
/*
  To use OutlookTabCtrl along with a style class:
    OutlookTabCtrlEx<OutlookTabCtrlCustomTooltip> m_OutlookTabCtrl;
  OutlookTabCtrlCustom2, OutlookTabCtrlCustom3 or OutlookTabCtrlCustom4 can be used instead of OutlookTabCtrlCustomTooltip.  
*/ 
template<typename STYLE>
struct OutlookTabCtrlEx : OutlookTabCtrl
{	OutlookTabCtrlEx()
	{	style.Install(this);
	}
	STYLE style;
};
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////




























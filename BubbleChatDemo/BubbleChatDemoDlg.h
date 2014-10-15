
// BubbleChatDemoDlg.h : header file
//

#pragma once

#include "BubbleRichEditCtrl.h"

// CBubbleChatDemoDlg dialog
class CBubbleChatDemoDlg : public CDialogEx
{
// Construction
public:
	CBubbleChatDemoDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_BUBBLECHATDEMO_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()


private:

	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

private:
	BubbleRichEditCtrl* m_pRichEdit;
};

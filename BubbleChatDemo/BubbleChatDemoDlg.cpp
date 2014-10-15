
// BubbleChatDemoDlg.cpp : implementation file
//

#include "stdafx.h"
#include "BubbleChatDemo.h"
#include "BubbleChatDemoDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CBubbleChatDemoDlg dialog




CBubbleChatDemoDlg::CBubbleChatDemoDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CBubbleChatDemoDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pRichEdit = NULL;
}

void CBubbleChatDemoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CBubbleChatDemoDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
END_MESSAGE_MAP()


// CBubbleChatDemoDlg message handlers

BOOL CBubbleChatDemoDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	if (!m_pRichEdit)
	{
		m_pRichEdit = new BubbleRichEditCtrl;
		m_pRichEdit->Attach(m_hWnd, CRect(30, 30, 500, 300), ES_MULTILINE | ES_READONLY);
	}

	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CBubbleChatDemoDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CBubbleChatDemoDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



LRESULT CBubbleChatDemoDlg::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	// TODO: Add your specialized code here and/or call the base class
	if (!m_pRichEdit)
	{
		return __super::WindowProc(message, wParam, lParam);
	}

		switch(message)
		{
		   case WM_SETCURSOR:
			{
				POINT pt;
				GetCursorPos(&pt);
				::ScreenToClient(m_hWnd, &pt);
            if(!m_pRichEdit->DoSetCursor(NULL, &pt) )
            {
               SetCursor(LoadCursor(NULL, IDC_ARROW));
            }
				break;
			}
		   case WM_RBUTTONDOWN:
			   {
				   m_pRichEdit->OutPutText(L"quick brown fox jumps over the lazy dog", false);
				   break;;
			   }

		   case WM_DESTROY:
			{
				// Release the controls
				m_pRichEdit->Release();



   			break;
			}

		   case WM_PAINT:
			{
				PAINTSTRUCT ps;

				HDC hdc;
				RECT rc;

				hdc = ::BeginPaint(m_hWnd, &ps);



				// Force clear out of window area being repainted
				::GetClientRect(m_hWnd, &rc);
				::FillRect(hdc, &rc, (HBRUSH)CreateSolidBrush(RGB(255,255,0)));
				//ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);

				// Because we are not really passing the message through Window's DefWinProc
				// we can doctor the the parameters.
				WPARAM w = (WPARAM) hdc;
				LPARAM l = (LPARAM) &ps.rcPaint;

				if ((0 == ps.rcPaint.bottom) && (0 == ps.rcPaint.right))
				{
					l = 0;
				}

				// First give the message to whoever has the focus.
				if (m_pRichEdit)
				{
					m_pRichEdit->drawOther(hdc);
					m_pRichEdit->TxWindowProc(m_hWnd, message, w, l);
				}

				::EndPaint(m_hWnd, &ps);
				return 0;
            //break ;
			}

		case WM_TIMER:
			lParam = 0;
         break ;
		default:
//DEF_PROCESSING:
			// First give the message to whoever has the focus.
			if (m_pRichEdit && this->GetFocus())
			{
				LRESULT lres = m_pRichEdit->TxWindowProc(m_hWnd, message, wParam, lParam);
				if (lres != S_MSG_KEY_IGNORED)
				{
					return lres;
				}
			}
			break;
      }


	return CDialogEx::WindowProc(message, wParam, lParam);
}

#include "StdAfx.h"
#include "BubbleRichEditCtrl.h"


BOOL fInAssert = FALSE;

// HIMETRIC units per inch (used for conversion)
#define HIMETRIC_PER_INCH 2540

// Convert Himetric along the X axis to X pixels
LONG HimetricXtoDX(LONG xHimetric, LONG xPerInch)
{
	return (LONG) MulDiv(xHimetric, xPerInch, HIMETRIC_PER_INCH);
}

// Convert Himetric along the Y axis to Y pixels
LONG HimetricYtoDY(LONG yHimetric, LONG yPerInch)
{
	return (LONG) MulDiv(yHimetric, yPerInch, HIMETRIC_PER_INCH);
}

// Convert Pixels on the X axis to Himetric
LONG DXtoHimetricX(LONG dx, LONG xPerInch)
{
	return (LONG) MulDiv(dx, HIMETRIC_PER_INCH, xPerInch);
}

// Convert Pixels on the Y axis to Himetric
LONG DYtoHimetricY(LONG dy, LONG yPerInch)
{
	return (LONG) MulDiv(dy, HIMETRIC_PER_INCH, yPerInch);
}

// These constants are for backward compatibility. They are the 
// sizes used for initialization and reset in RichEdit 1.0

const LONG cInitTextMax = (32 * 1024) - 1;
const LONG cResetTextMax = (64 * 1024);


#define ibPed 0
#define SYS_ALTERNATE 0x20000000




INT	cxBorder, cyBorder;			// GetSystemMetricx(SM_CXBORDER)...
INT	cxDoubleClk, cyDoubleClk;   // Double click distances
INT	cxHScroll, cxVScroll;	    // Width/height of scrlbar arw bitmap
INT	cyHScroll, cyVScroll;	    // Width/height of scrlbar arw bitmap
INT	dct;						// Double Click Time in milliseconds
INT     nScrollInset;
COLORREF crAuto = 0;

LONG BubbleRichEditCtrl::xWidthSys = 0;    		    // average char width of system font
LONG BubbleRichEditCtrl::yHeightSys = 0;			// height of system font
LONG BubbleRichEditCtrl::yPerInch = 0;				// y pixels per inch
LONG BubbleRichEditCtrl::xPerInch = 0;				// x pixels per inch


EXTERN_C const IID IID_ITextServices = { // 8d33f740-cf58-11ce-a89d-00aa006cadc5
	0x8d33f740,
	0xcf58,
	0x11ce,
	{0xa8, 0x9d, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5}
};

EXTERN_C const IID IID_ITextHost = { /* c5bdd8d0-d26e-11ce-a89e-00aa006cadc5 */
	0xc5bdd8d0,
	0xd26e,
	0x11ce,
	{0xa8, 0x9e, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5}
};

EXTERN_C const IID IID_ITextDocument = { /*8CC497C0-A1DF-11ce-8098-00AA0047BE5D*/
	0x8CC497C0,
	0xA1DF,
	0x11CE,
	{0x80,0x98,	0x00,0xAA,0x00,0x47,0xBE,0x5D}
};

EXTERN_C const IID IID_ITextEditControl = { /* f6642620-d266-11ce-a89e-00aa006cadc5 */
	0xf6642620,
	0xd266,
	0x11ce,
	{0xa8, 0x9e, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5}
};

void GetSysParms(void)
{
	crAuto		= GetSysColor(COLOR_WINDOWTEXT);
	cxBorder	= GetSystemMetrics(SM_CXBORDER);	// Unsizable window border
	cyBorder	= GetSystemMetrics(SM_CYBORDER);	//  widths
	cxHScroll	= GetSystemMetrics(SM_CXHSCROLL);	// Scrollbar-arrow bitmap 
	cxVScroll	= GetSystemMetrics(SM_CXVSCROLL);	//  dimensions
	cyHScroll	= GetSystemMetrics(SM_CYHSCROLL);	//
	cyVScroll	= GetSystemMetrics(SM_CYVSCROLL);	//
	cxDoubleClk	= GetSystemMetrics(SM_CXDOUBLECLK);
	cyDoubleClk	= GetSystemMetrics(SM_CYDOUBLECLK);
	dct			= GetDoubleClickTime();

	nScrollInset =
		GetProfileIntA( "windows", "ScrollInset", DD_DEFSCROLLINSET );
}

HRESULT InitDefaultCharFormat(CHARFORMATW * pcf, HFONT hfont) 
{
	HWND hwnd;
	LOGFONT lf;
	HDC hdc;
	LONG yPixPerInch;

	// Get LOGFONT for default font
	if (!hfont)
		hfont = (HFONT)GetStockObject(SYSTEM_FONT);

	// Get LOGFONT for passed hfont
	if (!GetObject(hfont, sizeof(LOGFONT), &lf))
		return E_FAIL;

	// Set CHARFORMAT structure
	pcf->cbSize = sizeof(CHARFORMAT2);

	hwnd = GetDesktopWindow();
	hdc = GetDC(hwnd);
	yPixPerInch = GetDeviceCaps(hdc, LOGPIXELSY);
	pcf->yHeight = lf.lfHeight * LY_PER_INCH / yPixPerInch;
	ReleaseDC(hwnd, hdc);

	pcf->yOffset = 0;
	pcf->crTextColor = crAuto;

	pcf->dwEffects = CFM_EFFECTS | CFE_AUTOBACKCOLOR;
	pcf->dwEffects &= ~(CFE_PROTECTED | CFE_LINK);

	if(lf.lfWeight < FW_BOLD)
		pcf->dwEffects &= ~CFE_BOLD;
	if(!lf.lfItalic)
		pcf->dwEffects &= ~CFE_ITALIC;
	if(!lf.lfUnderline)
		pcf->dwEffects &= ~CFE_UNDERLINE;
	if(!lf.lfStrikeOut)
		pcf->dwEffects &= ~CFE_STRIKEOUT;

	pcf->dwMask = CFM_ALL | CFM_BACKCOLOR;
	pcf->bCharSet = lf.lfCharSet;
	pcf->bPitchAndFamily = lf.lfPitchAndFamily;
//#ifdef UNICODE
//	_tcscpy(pcf->szFaceName, lf.lfFaceName);
//#else
//	//need to thunk pcf->szFaceName to a standard char string.in this case it's easy because our thunk is also our copy
//	MultiByteToWideChar(CP_ACP, 0, lf.lfFaceName, LF_FACESIZE, pcf->szFaceName, LF_FACESIZE) ;
//#endif

	wcscpy(pcf->szFaceName, L"Î¢ÈíÑÅºÚ");

	return S_OK;
}

HRESULT InitDefaultParaFormat(PARAFORMAT * ppf) 
{	
	memset(ppf, 0, sizeof(PARAFORMAT));

	ppf->cbSize = sizeof(PARAFORMAT);
	ppf->dwMask = PFM_ALL;
	ppf->wAlignment = PFA_LEFT;
	ppf->cTabCount = 1;
	ppf->rgxTabs[0] = lDefaultTab;

	return S_OK;
}



LRESULT MapHresultToLresult(HRESULT hr, UINT msg)
{
	LRESULT lres = hr;

	switch(msg)
	{
	case EM_GETMODIFY:


		lres = (hr == S_OK) ? -1 : 0;

		break;

		// These messages must return TRUE/FALSE rather than an hresult.
	case EM_UNDO:
	case WM_UNDO:
	case EM_CANUNDO:
	case EM_CANPASTE:
	case EM_LINESCROLL:

		// Hresults are backwards from TRUE and FALSE so we need
		// to do that remapping here as well.

		lres = (hr == S_OK) ? TRUE : FALSE;

		break;

	case EM_EXLINEFROMCHAR:
	case EM_LINEFROMCHAR:

		// If success, then hr a number. If error, it s/b 0.
		lres = SUCCEEDED(hr) ? (LRESULT) hr : 0;
		break;

	case EM_LINEINDEX:

		// If success, then hr a number. If error, it s/b -1.
		lres = SUCCEEDED(hr) ? (LRESULT) hr : -1;
		break;	

	default:
		lres = (LRESULT) hr;		
	}

	return lres;
}


BOOL GetIconic(HWND hwnd) 
{
	while(hwnd)
	{
		if(::IsIconic(hwnd))
			return TRUE;
		hwnd = GetParent(hwnd);
	}
	return FALSE;
}



BubbleRichEditCtrl::BubbleRichEditCtrl(void)
{
	ZeroMemory(&pnc, sizeof(BubbleRichEditCtrl) - offsetof(BubbleRichEditCtrl, pnc));
	cchTextMost = cInitTextMax;
	laccelpos = -1;
}


BubbleRichEditCtrl::~BubbleRichEditCtrl(void)
{	
	// Revoke our drop target
	RevokeDragDrop();
	pserv->OnTxInPlaceDeactivate();
	pserv->Release();
	if (mMemDC)
	{
		DeleteDC(mMemDC);
	}
}

////////////////////// Create/Init/Destruct Commands ///////////////////////
HRESULT CreateHost(HWND hwnd, const CREATESTRUCT *pcs, PNOTIFY_CALL pnc, BubbleRichEditCtrl **pptec)
{
	//HRESULT hr = E_FAIL;
	//GdiSetBatchLimit(1);
	//BubbleRichEditCtrl *phost = new BubbleRichEditCtrl();
	//if(phost)
	//{
	//	if (phost->Init(hwnd, pcs, pnc))
	//	{
	//		*pptec = phost;
	//		hr = S_OK;
	//	}
	//}
	//if (FAILED(hr))
	//{
	//	delete phost;
	//}
	return TRUE;
}

/*
 *	BubbleRichEditCtrl::Init
 *
 *	Purpose:
 *		Initializes this BubbleRichEditCtrl
 *
 *	TODO:	Format Cache logic needs to be cleaned up. 
 */
BOOL BubbleRichEditCtrl::Init(/*HWND h_wnd, const CREATESTRUCT *pcs,PNOTIFY_CALL p_nc*/)
{
    HDC hdc;
    HFONT hfontOld;
    TEXTMETRIC tm;
	IUnknown *pUnk;
	HRESULT hr;

	// Initialize Reference count
	cRefs = 1;

	// Set up the notification callback
	//pnc = p_nc;	
	pnc = NULL;

	hwnd = mHwnd;

	// Create and cache CHARFORMAT for this control
	if(FAILED(InitDefaultCharFormat(&cf, NULL)))
		goto err;
		
	// Create and cache PARAFORMAT for this control
	if(FAILED(InitDefaultParaFormat(&pf)))
		goto err;

 	// edit controls created without a window are multiline by default
	// so that paragraph formats can be
	mStyle = mStyle|ES_MULTILINE;
	fHidden = TRUE;
	// edit controls are rich by default
	fRich = TRUE;
	mStyle = mStyle | (ES_DISABLENOSCROLL | ES_AUTOVSCROLL | ES_READONLY);

	if(!(mStyle & ES_LEFT))
	{
		if(mStyle & ES_CENTER)
			pf.wAlignment = PFA_CENTER;
		else if(mStyle & ES_RIGHT)
			pf.wAlignment = PFA_RIGHT;
	}

    // Init system metrics
	hdc = GetDC(hwnd);
    if(!hdc)
        goto err;

	//[[
	mMemDC = ::CreateCompatibleDC(hdc);
	
	// TODO: bitmap size should be dynamic, what is the max resolution of screen? 
	mBgBitMap = ::CreateCompatibleBitmap(hdc, 1000, 1000);  
	SelectObject(mMemDC, mBgBitMap);
	RECT rt = {0, 0, 1000, 1000};
	//FillRect(mMemDC, &rcClient, (HBRUSH)GetStockObject(BLACK_BRUSH));
	//]]

   	hfontOld = (HFONT)SelectObject(hdc, GetStockObject(SYSTEM_FONT));

	if(!hfontOld)
		goto err;

	GetTextMetrics(hdc, &tm);
	SelectObject(hdc, hfontOld);

	xWidthSys = (INT) tm.tmAveCharWidth;
    yHeightSys = (INT) tm.tmHeight;
	xPerInch = GetDeviceCaps(hdc, LOGPIXELSX); 
	yPerInch =	GetDeviceCaps(hdc, LOGPIXELSY); 

	ReleaseDC(hwnd, hdc);

	// At this point the border flag is set and so is the pixels per inch
	// so we can initalize the inset.
	SetDefaultInset();

	fInplaceActive = TRUE;

	fTransparent = TRUE;
	// Create Text Services component
	if(FAILED(CreateTextServices(NULL, this, &pUnk)))
		goto err;

	hr = pUnk->QueryInterface(IID_ITextServices,(void **)&pserv);
	
	// Whether the previous call succeeded or failed we are done
	// with the private interface.
	pUnk->Release();

	if(FAILED(hr))
	{
		goto err;
	}


	//rcClient.left = pcs->x;
	//rcClient.top = pcs->y;
	//rcClient.right = pcs->x + pcs->cx;
	//rcClient.bottom = pcs->y + pcs->cy;

	rcClient = mRERect;
	mMaxREWidth = mRERect.right - mRERect.left;

	// The extent matches the full client rectangle in HIMETRIC
	sizelExtent.cx = DXtoHimetricX(mRERect.left - 2 * HOST_BORDER, xPerInch);
	sizelExtent.cy = DYtoHimetricY(mRERect.top - 2 * HOST_BORDER, yPerInch);

	// notify Text Services that we are in place active
	if(FAILED(pserv->OnTxInPlaceActivate(&rcClient)))
	{
		ASSERT(0);
		return FALSE;
	}

	if (!(mStyle & ES_READONLY))
	{
		// This isn't a read only window so we need a drop target.
		RegisterDragDrop();
	}
	
	hr = pserv->TxSendMessage(EM_GETOLEINTERFACE, 0, (LPARAM)&pRichEditOle, NULL);
	hr = pRichEditOle->QueryInterface(IID_ITextDocument, (void **)&pDocument);
	mITextServices = pserv;
	mIRichEditOle = pRichEditOle;
	mIDocument = pDocument;

	return TRUE;

err:
	return FALSE;
}


/////////////////////////////////  IUnknown ////////////////////////////////
HRESULT BubbleRichEditCtrl::QueryInterface(REFIID riid, void **ppvObject)
{
	HRESULT hr = E_NOINTERFACE;
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_ITextEditControl))
	{
		*ppvObject = (ITextEditControl *) this;
		AddRef();
		hr = S_OK;
	}
	else if (IsEqualIID(riid, IID_IUnknown) 
		|| IsEqualIID(riid, IID_ITextHost)) 
	{
		AddRef();
		*ppvObject = (ITextHost *) this;
		hr = S_OK;
	}

	return hr;
}

ULONG BubbleRichEditCtrl::AddRef(void)
{
	return ++cRefs;
}

ULONG BubbleRichEditCtrl::Release(void)
{
	ULONG c_Refs = --cRefs;

	if (c_Refs == 0)
	{
		delete this;
	}

	return c_Refs;
}


//////////////////////////////// Properties ////////////////////////////////


TXTEFFECT BubbleRichEditCtrl::TxGetEffects() const
{
	return (mStyle & ES_SUNKEN) ? TXTEFFECT_SUNKEN : TXTEFFECT_NONE;
}


//////////////////////////// System API wrapper ////////////////////////////



///////////////////////  Windows message dispatch methods  ///////////////////////////////


LRESULT BubbleRichEditCtrl::TxWindowProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam)
{
	LRESULT	lres = 0;
	HRESULT hr;

	switch(msg)
	{
	case WM_NCCALCSIZE:
		// we can't rely on WM_WININICHANGE so we use WM_NCCALCSIZE since
		// changing any of these should trigger a WM_NCCALCSIZE
		GetSysParms();
		break;

	case WM_KEYDOWN:
		lres = OnKeyDown((WORD) wparam, (DWORD) lparam);
		if(lres != 0)
			goto serv;
		break;		   

	case WM_CHAR:
		lres = OnChar((WORD) wparam, (DWORD) lparam);
		if(lres != 0)
			goto serv;
		break;

	case WM_SYSCOLORCHANGE:
		OnSysColorChange();

		// Notify the text services that there has been a change in the 
		// system colors.
		goto serv;

	case WM_GETDLGCODE:
		lres = OnGetDlgCode(wparam, lparam);
		break;

	case EM_HIDESELECTION:
		if((BOOL)lparam)
		{
			DWORD dwPropertyBits = 0;

			if((BOOL)wparam)
			{
				mStyle &= ~(DWORD) ES_NOHIDESEL;
				dwPropertyBits = TXTBIT_HIDESELECTION;
			}
			else
				mStyle |= ES_NOHIDESEL;

			// Notify text services of change in status.
			pserv->OnTxPropertyBitsChange(TXTBIT_HIDESELECTION, 
				dwPropertyBits);
		}

		goto serv;



    case EM_LIMITTEXT:

        lparam = wparam;

        // Intentionally fall through. These messages are duplicates.

	case EM_EXLIMITTEXT:

        if (lparam == 0)
        {
            // 0 means set the control to the maximum size. However, because
            // 1.0 set this to 64K will keep this the same value so as not to
            // supprise anyone. Apps are free to set the value to be above 64K.
            lparam = (LPARAM) cResetTextMax;
        }

		cchTextMost = (LONG) lparam;
		pserv->OnTxPropertyBitsChange(TXTBIT_MAXLENGTHCHANGE, 
					TXTBIT_MAXLENGTHCHANGE);

		break;

	case EM_SETREADONLY:
		OnSetReadOnly(BOOL(wparam));
		lres = 1;
		break;

	case EM_GETEVENTMASK:
		lres = OnGetEventMask();
		break;

	case EM_SETEVENTMASK:
		OnSetEventMask((DWORD) lparam);
		goto serv;

	case EM_GETOPTIONS:
		lres = OnGetOptions();
		break;

	case EM_SETOPTIONS:
		OnSetOptions((WORD) wparam, (DWORD) lparam);
		lres = (mStyle & ECO_STYLES);
		if(fEnableAutoWordSel)
			lres |= ECO_AUTOWORDSELECTION;
		break;

	case WM_SETFONT:
		lres = OnSetFont((HFONT) wparam);
		break;

	case EM_SETRECT:
        OnSetRect((LPRECT)lparam);
        break;
        
	case EM_GETRECT:
        OnGetRect((LPRECT)lparam);
        break;

	case EM_SETBKGNDCOLOR:

		lres = (LRESULT) crBackground;
		fNotSysBkgnd = !wparam;
		crBackground = (COLORREF) lparam;

		if(wparam)
			//crBackground = GetSysColor(COLOR_WINDOW);
			crBackground = RGB(120,120,120);

		// Notify the text services that color has changed
		pserv->TxSendMessage(WM_SYSCOLORCHANGE, 0, 0, 0);

		if(lres != (LRESULT) crBackground)
			TxInvalidateRect(NULL, TRUE);

		break;

	case EM_SETCHARFORMAT:

		if(!FValidCF((CHARFORMAT *) lparam))
		{
			return 0;
		}

		if(wparam & SCF_SELECTION)
			goto serv;								// Change selection format
		OnSetCharFormat((CHARFORMAT *) lparam);		// Change default format
		break;

	case EM_SETPARAFORMAT:
		if(!FValidPF((PARAFORMAT *) lparam))
		{
			return 0;
		}

		// check to see if we're setting the default.
		// either SCF_DEFAULT will be specified *or* there is no
		// no text in the document (richedit1.0 behaviour).
		if (!(wparam & SCF_DEFAULT))
		{
			hr = pserv->TxSendMessage(WM_GETTEXTLENGTH, 0, 0, 0);

			if (hr == 0)
			{
				wparam |= SCF_DEFAULT;
			}
		}

		if(wparam & SCF_DEFAULT)
		{								
			OnSetParaFormat((PARAFORMAT *) lparam);	// Change default format
		}
		else
		{
			goto serv;								// Change selection format
		}
		break;

    case WM_SETTEXT:

        // For RichEdit 1.0, the max text length would be reset by a settext so 
        // we follow pattern here as well.

		hr = pserv->TxSendMessage(msg, wparam, lparam, 0);

        if (SUCCEEDED(hr))
        {
            // Update succeeded.
            LONG cNewText = _tcslen((LPCTSTR) lparam);

            // If the new text is greater than the max set the max to the new
            // text length.
            if (cNewText > cchTextMost)
            {
                cchTextMost = cNewText;                
            }

			lres = 1;
        }

        break;

	case WM_SIZE:
		lres = OnSize(hwnd, wparam, LOWORD(lparam), HIWORD(lparam));
		break;

	case WM_WINDOWPOSCHANGING:
		lres = ::DefWindowProc(hwnd, msg, wparam, lparam);

		if(TxGetEffects() == TXTEFFECT_SUNKEN)
			OnSunkenWindowPosChanging(hwnd, (WINDOWPOS *) lparam);
		break;

	case WM_SETCURSOR:
		//Only set cursor when over us rather than a child; this
		//			helps prevent us from fighting it out with an inplace child
		if((HWND)wparam == hwnd)
		{
			if(!(lres = ::DefWindowProc(hwnd, msg, wparam, lparam)))
			{
				POINT pt;
				GetCursorPos(&pt);
				::ScreenToClient(hwnd, &pt);
				pserv->OnTxSetCursor(
					DVASPECT_CONTENT,	
					-1,
					NULL,
					NULL,
					NULL,
					NULL,
					NULL,			// Client rect - no redraw 
					pt.x, 
					pt.y);
				lres = TRUE;
			}
		}
		break;

	case WM_SHOWWINDOW:
		hr = OnTxVisibleChange((BOOL)wparam);
		break;

	case WM_NCPAINT:

		lres = ::DefWindowProc(hwnd, msg, wparam, lparam);

		if(TxGetEffects() == TXTEFFECT_SUNKEN)
		{
			HDC hdc = GetDC(hwnd);

			if(hdc)
			{
				DrawSunkenBorder(hwnd, hdc);
				ReleaseDC(hwnd, hdc);
			}
		}
		break;

	case WM_PAINT:
		{
			// Put a frame around the control so it can be seen
			HDC hdc = (HDC)wparam;
			RECT rcClient;
			GetControlRect(&rcClient);
			//FrameRect((HDC) wparam, &rcClient, (HBRUSH) GetStockObject(BLACK_BRUSH));
			//FillRect(hdc, &rcClient, (HBRUSH)CreateSolidBrush(RGB(255,0,0)));
			//BOOL res = BitBlt(hdc, rcClient.left, rcClient.right, rcClient.right - rcClient.left, rcClient.bottom - rcClient.top, mMemDC, 0, 0, SRCCOPY);
			//if (!res)
			//{
			//	int err = GetLastError();
			//	::MessageBoxA(NULL, NULL, NULL, 0);
			//}

			DrawBubbleToBg(hdc);

			RECT *prc = NULL;
			LONG lViewId = TXTVIEW_ACTIVE;

			if (!fInplaceActive)
			{
				GetControlRect(&rcClient);
				prc = &rcClient;
				lViewId = TXTVIEW_INACTIVE;
			}
			
			// Remember wparam is actually the hdc and lparam is the update
			// rect because this message has been preprocessed by the window.
			pserv->TxDraw(
	    		DVASPECT_CONTENT,  		// Draw Aspect
				/*-1*/0,						// Lindex
				NULL,					// Info for drawing optimazation
				NULL,					// target device information
	        	(HDC) wparam,			// Draw device HDC
	        	NULL, 				   	// Target device HDC
				(RECTL *) prc,			// Bounding client rectangle
				NULL, 					// Clipping rectangle for metafiles
				(RECT *) lparam,		// Update rectangle
				NULL, 	   				// Call back function
				NULL,					// Call back parameter
				lViewId);				// What view of the object				

			if(TxGetEffects() == TXTEFFECT_SUNKEN)
				DrawSunkenBorder(hwnd, (HDC) wparam);

			//SetTransparent((HDC) wparam, )

			// .,.. IRichEditOle 
		}

		break;

	default:
serv:
		{
			hr = pserv->TxSendMessage(msg, wparam, lparam, &lres);

			if (hr == S_FALSE)
			{
				lres = ::DefWindowProc(hwnd, msg, wparam, lparam);
			}
		}
	}

	return lres;
}
	

///////////////////////////////  Keyboard Messages  //////////////////////////////////


LRESULT BubbleRichEditCtrl::OnKeyDown(WORD vkey, DWORD dwFlags)
{
	switch(vkey)
	{
	case VK_ESCAPE:
		if(fInDialogBox)
		{
			PostMessage(hwndParent, WM_CLOSE, 0, 0);
			return 0;
		}
		break;
	
	case VK_RETURN:
		if(fInDialogBox && !(GetKeyState(VK_CONTROL) & 0x8000) 
				&& !(mStyle & ES_WANTRETURN))
		{
			// send to default button
			LRESULT id;
			HWND hwndT;

			id = SendMessage(hwndParent, DM_GETDEFID, 0, 0);
			if(LOWORD(id) &&
				(hwndT = GetDlgItem(hwndParent, LOWORD(id))))
			{
				SendMessage(hwndParent, WM_NEXTDLGCTL, (WPARAM) hwndT, (LPARAM) 1);
				if(GetFocus() != hwnd)
					PostMessage(hwndT, WM_KEYDOWN, (WPARAM) VK_RETURN, 0);
			}
			return 0;
		}
		break;

	case VK_TAB:
		if(fInDialogBox) 
		{
			SendMessage(hwndParent, WM_NEXTDLGCTL, 
							!!(GetKeyState(VK_SHIFT) & 0x8000), 0);
			return 0;
		}
		break;
	}
	return 1;
}

#define CTRL(_ch) (_ch - 'A' + 1)

LRESULT BubbleRichEditCtrl::OnChar(WORD vkey, DWORD dwFlags)
{
	switch(vkey)
	{
	// Ctrl-Return generates Ctrl-J (LF), treat it as an ordinary return
	case CTRL('J'):
	case VK_RETURN:
		if(fInDialogBox && !(GetKeyState(VK_CONTROL) & 0x8000)
				 && !(mStyle & ES_WANTRETURN))
			return 0;
		break;

	case VK_TAB:
		if(fInDialogBox && !(GetKeyState(VK_CONTROL) & 0x8000))
			return 0;
	}
	
	return 1;
}


////////////////////////////////////  View rectangle //////////////////////////////////////


void BubbleRichEditCtrl::OnGetRect(LPRECT prc)
{
    RECT rcInset;

	// Get view inset (in HIMETRIC)
    TxGetViewInset(&rcInset);

	// Convert the himetric inset to pixels
	rcInset.left = HimetricXtoDX(rcInset.left, xPerInch);
	rcInset.top = HimetricYtoDY(rcInset.top , yPerInch);
	rcInset.right = HimetricXtoDX(rcInset.right, xPerInch);
	rcInset.bottom = HimetricYtoDY(rcInset.bottom, yPerInch);
    
	// Get client rect in pixels
    TxGetClientRect(prc);

	// Modify the client rect by the inset 
    prc->left += rcInset.left;
    prc->top += rcInset.top;
    prc->right -= rcInset.right;
    prc->bottom -= rcInset.bottom;
}

void BubbleRichEditCtrl::OnSetRect(LPRECT prc)
{
	RECT rcClient;
	
	if(!prc)
	{
		SetDefaultInset();
	}	
	else	
    {
    	// For screen display, the following intersects new view RECT
    	// with adjusted client area RECT
    	TxGetClientRect(&rcClient);

        // Adjust client rect
        // Factors in space for borders
        if(fBorder)
        {																					  
    	    rcClient.top		+= yHeightSys / 4;
    	    rcClient.bottom 	-= yHeightSys / 4 - 1;
    	    rcClient.left		+= xWidthSys / 2;
    	    rcClient.right	-= xWidthSys / 2;
        }
	
        // Ensure we have minimum width and height
        rcClient.right = max(rcClient.right, rcClient.left + xWidthSys);
        rcClient.bottom = max(rcClient.bottom, rcClient.top + yHeightSys);

        // Intersect the new view rectangle with the 
        // adjusted client area rectangle
        if(!IntersectRect(&rcViewInset, &rcClient, prc))
    	    rcViewInset = rcClient;

        // compute inset in pixels
        rcViewInset.left -= rcClient.left;
        rcViewInset.top -= rcClient.top;
        rcViewInset.right = rcClient.right - rcViewInset.right;
        rcViewInset.bottom = rcClient.bottom - rcViewInset.bottom;

		// Convert the inset to himetric that must be returned to ITextServices
        rcViewInset.left = DXtoHimetricX(rcViewInset.left, xPerInch);
        rcViewInset.top = DYtoHimetricY(rcViewInset.top, yPerInch);
        rcViewInset.right = DXtoHimetricX(rcViewInset.right, xPerInch);
        rcViewInset.bottom = DYtoHimetricY(rcViewInset.bottom, yPerInch);
    }

    pserv->OnTxPropertyBitsChange(TXTBIT_VIEWINSETCHANGE, 
    	TXTBIT_VIEWINSETCHANGE);
}



////////////////////////////////////  System notifications  //////////////////////////////////


void BubbleRichEditCtrl::OnSysColorChange()
{
	crAuto = GetSysColor(COLOR_WINDOWTEXT);
	if(!fNotSysBkgnd)
		crBackground = GetSysColor(COLOR_WINDOW);
	TxInvalidateRect(NULL, TRUE);
}

LRESULT BubbleRichEditCtrl::OnGetDlgCode(WPARAM wparam, LPARAM lparam)
{
	LRESULT lres = DLGC_WANTCHARS | DLGC_WANTARROWS | DLGC_WANTTAB;

	if(mStyle & ES_MULTILINE)
		lres |= DLGC_WANTALLKEYS;

	if(!(mStyle & ES_SAVESEL))
		lres |= DLGC_HASSETSEL;

	if(lparam)
		fInDialogBox = TRUE;

	if(lparam &&
		((WORD) wparam == VK_BACK))
	{
		lres |= DLGC_WANTMESSAGE;
	}

	return lres;
}


/////////////////////////////////  Other messages  //////////////////////////////////////


LRESULT BubbleRichEditCtrl::OnGetOptions() const
{
	LRESULT lres = (mStyle & ECO_STYLES);

	if(fEnableAutoWordSel)
		lres |= ECO_AUTOWORDSELECTION;
	
	return lres;
}

void BubbleRichEditCtrl::OnSetOptions(WORD wOp, DWORD eco)
{
	const BOOL fAutoWordSel = !!(eco & ECO_AUTOWORDSELECTION);
	DWORD mStyleNew = mStyle;
	DWORD dw_Style = 0 ;

	DWORD dwChangeMask = 0;

	// single line controls can't have a selection bar
	// or do vertical writing
	if(!(dw_Style & ES_MULTILINE))
	{
#ifdef DBCS
		eco &= ~(ECO_SELECTIONBAR | ECO_VERTICAL);
#else
		eco &= ~ECO_SELECTIONBAR;
#endif
	}
	dw_Style = (eco & ECO_STYLES);

	switch(wOp)
	{
	case ECOOP_SET:
		mStyleNew = ((mStyleNew) & ~ECO_STYLES) | mStyle;
		fEnableAutoWordSel = fAutoWordSel;
		break;

	case ECOOP_OR:
		mStyleNew |= dw_Style;
		if(fAutoWordSel)
			fEnableAutoWordSel = TRUE;
		break;

	case ECOOP_AND:
		mStyleNew &= (dw_Style | ~ECO_STYLES);
		if(fEnableAutoWordSel && !fAutoWordSel)
			fEnableAutoWordSel = FALSE;
		break;

	case ECOOP_XOR:
		mStyleNew ^= dw_Style;
		fEnableAutoWordSel = (!fEnableAutoWordSel != !fAutoWordSel);
		break;
	}

	if(fEnableAutoWordSel != (unsigned)fAutoWordSel)
	{
		dwChangeMask |= TXTBIT_AUTOWORDSEL; 
	}

	if(mStyleNew != dw_Style)
	{
		DWORD dwChange = mStyleNew ^ dw_Style;
#ifdef DBCS
		USHORT	usMode;
#endif

		mStyle = mStyleNew;
		SetWindowLong(hwnd, GWL_STYLE, mStyleNew);

		if(dwChange & ES_NOHIDESEL)	
		{
			dwChangeMask |= TXTBIT_HIDESELECTION;
		}

		if(dwChange & ES_READONLY)
		{
			dwChangeMask |= TXTBIT_READONLY;

			// Change drop target state as appropriate.
			if (mStyleNew & ES_READONLY)
			{
				RevokeDragDrop();
			}
			else
			{
				RegisterDragDrop();
			}
		}

		if(dwChange & ES_VERTICAL)
		{
			dwChangeMask |= TXTBIT_VERTICAL;
		}

		// no action require for ES_WANTRETURN
		// no action require for ES_SAVESEL
		// do this last
		if(dwChange & ES_SELECTIONBAR)
		{
			lSelBarWidth = 212;
			dwChangeMask |= TXTBIT_SELBARCHANGE;
		}
	}

	if (dwChangeMask)
	{
		DWORD dwProp = 0;
		TxGetPropertyBits(dwChangeMask, &dwProp);
		pserv->OnTxPropertyBitsChange(dwChangeMask, dwProp);
	}
}

void BubbleRichEditCtrl::OnSetReadOnly(BOOL fReadOnly)
{
	DWORD dwUpdatedBits = 0;

	if(fReadOnly)
	{
		mStyle |= ES_READONLY;

		// Turn off Drag Drop 
		RevokeDragDrop();
		dwUpdatedBits |= TXTBIT_READONLY;
	}
	else
	{
		mStyle &= ~(DWORD) ES_READONLY;

		// Turn drag drop back on
		RegisterDragDrop();	
	}

	pserv->OnTxPropertyBitsChange(TXTBIT_READONLY, dwUpdatedBits);
}

void BubbleRichEditCtrl::OnSetEventMask(DWORD mask)
{
	LRESULT lres = (LRESULT) dwEventMask;
	dwEventMask = (DWORD) mask;

}


LRESULT BubbleRichEditCtrl::OnGetEventMask() const
{
	return (LRESULT) dwEventMask;
}

/*
 *	BubbleRichEditCtrl::OnSetFont(hfont)
 *
 *	Purpose:	
 *		Set new font from hfont
 *
 *	Arguments:
 *		hfont	new font (NULL for system font)
 */
BOOL BubbleRichEditCtrl::OnSetFont(HFONT hfont)
{
	if(SUCCEEDED(InitDefaultCharFormat(&cf, hfont)))
	{
		pserv->OnTxPropertyBitsChange(TXTBIT_CHARFORMATCHANGE, 
			TXTBIT_CHARFORMATCHANGE);
		return TRUE;
	}

	return FALSE;
}

/*
 *	BubbleRichEditCtrl::OnSetCharFormat(pcf)
 *
 *	Purpose:	
 *		Set new default CharFormat
 *
 *	Arguments:
 *		pch		ptr to new CHARFORMAT
 */
BOOL BubbleRichEditCtrl::OnSetCharFormat(CHARFORMAT *pcf)
{
	pserv->OnTxPropertyBitsChange(TXTBIT_CHARFORMATCHANGE, 
		TXTBIT_CHARFORMATCHANGE);

	return TRUE;
}

/*
 *	BubbleRichEditCtrl::OnSetParaFormat(ppf)
 *
 *	Purpose:	
 *		Set new default ParaFormat
 *
 *	Arguments:
 *		pch		ptr to new PARAFORMAT
 */
BOOL BubbleRichEditCtrl::OnSetParaFormat(PARAFORMAT *pPF)
{
	pf = *pPF;									// Copy it

	pserv->OnTxPropertyBitsChange(TXTBIT_PARAFORMATCHANGE, 
		TXTBIT_PARAFORMATCHANGE);

	return TRUE;
}



////////////////////////////  Event firing  /////////////////////////////////



void * BubbleRichEditCtrl::CreateNmhdr(UINT uiCode, LONG cb)
{
	NMHDR *pnmhdr;

	pnmhdr = (NMHDR*) new char[cb];
	if(!pnmhdr)
		return NULL;

	memset(pnmhdr, 0, cb);

	pnmhdr->hwndFrom = hwnd;
	pnmhdr->idFrom = GetWindowLong(hwnd, GWL_ID);
	pnmhdr->code = uiCode;

	return (VOID *) pnmhdr;
}


////////////////////////////////////  Helpers  /////////////////////////////////////////
void BubbleRichEditCtrl::SetDefaultInset()
{
    // Generate default view rect from client rect.
    if(fBorder)
    {
        // Factors in space for borders
  	    rcViewInset.top = DYtoHimetricY(yHeightSys / 4, yPerInch);
   	 rcViewInset.bottom	= DYtoHimetricY(yHeightSys / 4 - 1, yPerInch);
   	 rcViewInset.left = DXtoHimetricX(xWidthSys / 2, xPerInch);
   	 rcViewInset.right = DXtoHimetricX(xWidthSys / 2, xPerInch);
    }
    else
    {
		rcViewInset.top = rcViewInset.left =
		rcViewInset.bottom = rcViewInset.right = 0;
	}
}


/////////////////////////////////  Far East Support  //////////////////////////////////////

HIMC BubbleRichEditCtrl::TxImmGetContext(void)
{
#if 0
	HIMC himc;

	himc = ImmGetContext( hwnd );

	return himc;
#endif // 0

	return NULL;
}

void BubbleRichEditCtrl::TxImmReleaseContext(HIMC himc)
{
	ImmReleaseContext( hwnd, himc );
}

void BubbleRichEditCtrl::RevokeDragDrop(void)
{
	if (fRegisteredForDrop)
	{
		::RevokeDragDrop(hwnd);
		fRegisteredForDrop = FALSE;
	}
}

void BubbleRichEditCtrl::RegisterDragDrop(void)
{
	IDropTarget *pdt;

	if(!fRegisteredForDrop && pserv->TxGetDropTarget(&pdt) == NOERROR)
	{
		HRESULT hr = ::RegisterDragDrop(hwnd, pdt);

		if(hr == NOERROR)
		{	
			fRegisteredForDrop = TRUE;
		}

		pdt->Release();
	}
}

VOID DrawRectFn(
	HDC hdc, 
	RECT *prc, 
	INT icrTL, 
	INT icrBR,
	BOOL fBot, 
	BOOL fRght)
{
	COLORREF cr;
	COLORREF crSave;
	RECT rc;

	cr = GetSysColor(icrTL);
	crSave = SetBkColor(hdc, cr);

	// top
	rc = *prc;
	rc.bottom = rc.top + 1;
	ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);

	// left
	rc.bottom = prc->bottom;
	rc.right = rc.left + 1;
	ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);

	if(icrTL != icrBR)
	{
		cr = GetSysColor(icrBR);
		SetBkColor(hdc, cr);
	}

	// right
	rc.right = prc->right;
	rc.left = rc.right - 1;
	if(!fBot)
		rc.bottom -= cyHScroll;
	if(fRght)
	{
		ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);
	}

	// bottom
	if(fBot)
	{
		rc.left = prc->left;
		rc.top = rc.bottom - 1;
		if(!fRght)
			rc.right -= cxVScroll;
		ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, NULL, 0, NULL);
	}
	SetBkColor(hdc, crSave);
}

#define cmultBorder 1

VOID BubbleRichEditCtrl::OnSunkenWindowPosChanging(HWND hwnd, WINDOWPOS *pwndpos)
{
	if(fVisible)
	{
		RECT rc;
		HWND hwndParent;

		GetWindowRect(hwnd, &rc);
		InflateRect(&rc, cxBorder * cmultBorder, cyBorder * cmultBorder);
		hwndParent = GetParent(hwnd);
		MapWindowPoints(HWND_DESKTOP, hwndParent, (POINT *) &rc, 2);
		InvalidateRect(hwndParent, &rc, FALSE);
	}
}


VOID BubbleRichEditCtrl::DrawSunkenBorder(HWND hwnd, HDC hdc)
{
	RECT rc;
	RECT rcParent;
	DWORD dwScrollBars;
	HWND hwndParent;

	GetWindowRect(hwnd, &rc);
    hwndParent = GetParent(hwnd);
	rcParent = rc;
	MapWindowPoints(HWND_DESKTOP, hwndParent, (POINT *)&rcParent, 2);
	InflateRect(&rcParent, cxBorder, cyBorder);
	OffsetRect(&rc, -rc.left, -rc.top);

	// draw inner rect
	TxGetScrollBars(&dwScrollBars);
	DrawRectFn(hdc, &rc, COLOR_WINDOWFRAME, COLOR_BTNFACE,
		!(dwScrollBars & WS_HSCROLL), !(dwScrollBars & WS_VSCROLL));

	// draw outer rect
	hwndParent = GetParent(hwnd);
	hdc = GetDC(hwndParent);
	DrawRectFn(hdc, &rcParent, COLOR_BTNSHADOW, COLOR_BTNHIGHLIGHT,
		TRUE, TRUE);
	ReleaseDC(hwndParent, hdc);
}

LRESULT BubbleRichEditCtrl::OnSize(HWND hwnd, WORD fwSizeType, int nWidth, int nHeight)
{
	// Update our client rectangle
	rcClient.right = rcClient.left + nWidth / 2;
	rcClient.bottom = rcClient.top + nHeight / 2;

	mMaxREWidth = nWidth / 2;

	if(!fVisible)
	{
		fIconic = GetIconic(hwnd);
		if(!fIconic)
			fResized = TRUE;
	}
	else
	{
		if(GetIconic(hwnd))
		{
			fIconic = TRUE;
		}
		else
		{
			pserv->OnTxPropertyBitsChange(TXTBIT_CLIENTRECTCHANGE, 
				TXTBIT_CLIENTRECTCHANGE);

			if(fIconic)
			{
				InvalidateRect(hwnd, NULL, FALSE);
				fIconic = FALSE;
			}
			
			if(TxGetEffects() == TXTEFFECT_SUNKEN)	// Draw borders
				DrawSunkenBorder(hwnd, NULL);
		}
	}
	return 0;
}

HRESULT BubbleRichEditCtrl::OnTxVisibleChange(BOOL fVisible)
{
	fVisible = fVisible;

	if(!fVisible && fResized)
	{
		RECT rc;
		// Control was resized while hidden,
		// need to really resize now
		TxGetClientRect(&rc);
		fResized = FALSE;
		pserv->OnTxPropertyBitsChange(TXTBIT_CLIENTRECTCHANGE, 
			TXTBIT_CLIENTRECTCHANGE);
	}

	return S_OK;
}



//////////////////////////// ITextHost Interface  ////////////////////////////

HDC BubbleRichEditCtrl::TxGetDC()
{
	return ::GetDC(hwnd);
}


int BubbleRichEditCtrl::TxReleaseDC(HDC hdc)
{
	return ::ReleaseDC (hwnd, hdc);
}


BOOL BubbleRichEditCtrl::TxShowScrollBar(INT fnBar,	BOOL fShow)
{
	return ::ShowScrollBar(hwnd, fnBar, fShow);
}

BOOL BubbleRichEditCtrl::TxEnableScrollBar (INT fuSBFlags, INT fuArrowflags)
{
	return ::EnableScrollBar(hwnd, fuSBFlags, fuArrowflags) ;//SB_HORZ, ESB_DISABLE_BOTH);
}


BOOL BubbleRichEditCtrl::TxSetScrollRange(INT fnBar, LONG nMinPos, INT nMaxPos, BOOL fRedraw)
{
	return ::SetScrollRange(hwnd, fnBar, nMinPos, nMaxPos, fRedraw);
}


BOOL BubbleRichEditCtrl::TxSetScrollPos (INT fnBar, INT nPos, BOOL fRedraw)
{
	return ::SetScrollPos(hwnd, fnBar, nPos, fRedraw);
}

void BubbleRichEditCtrl::TxInvalidateRect(LPCRECT prc, BOOL fMode)
{
	::InvalidateRect(hwnd, prc, fMode);
}

void BubbleRichEditCtrl::TxViewChange(BOOL fUpdate) 
{
	::UpdateWindow (hwnd);
}


BOOL BubbleRichEditCtrl::TxCreateCaret(HBITMAP hbmp, INT xWidth, INT yHeight)
{
	return ::CreateCaret (hwnd, hbmp, xWidth, yHeight);
}


BOOL BubbleRichEditCtrl::TxShowCaret(BOOL fShow)
{
	if(fShow)
		return ::ShowCaret(hwnd);
	else
		return ::HideCaret(hwnd);
}

BOOL BubbleRichEditCtrl::TxSetCaretPos(INT x, INT y)
{
	return ::SetCaretPos(x, y);
}


BOOL BubbleRichEditCtrl::TxSetTimer(UINT idTimer, UINT uTimeout)
{
	fTimer = TRUE;
	return ::SetTimer(hwnd, idTimer, uTimeout, NULL);
}


void BubbleRichEditCtrl::TxKillTimer(UINT idTimer)
{
	::KillTimer(hwnd, idTimer);
	fTimer = FALSE;
}

void BubbleRichEditCtrl::TxScrollWindowEx (INT dx, INT dy, LPCRECT lprcScroll,	LPCRECT lprcClip,	HRGN hrgnUpdate, LPRECT lprcUpdate,	UINT fuScroll)	
{
	::ScrollWindowEx(hwnd, dx, dy, lprcScroll, lprcClip, hrgnUpdate, lprcUpdate, fuScroll);
}

void BubbleRichEditCtrl::TxSetCapture(BOOL fCapture)
{
	if (fCapture)
		::SetCapture(hwnd);
	else
		::ReleaseCapture();
}

void BubbleRichEditCtrl::TxSetFocus()
{
	::SetFocus(hwnd);
}

void BubbleRichEditCtrl::TxSetCursor(HCURSOR hcur,	BOOL fText)
{
	::SetCursor(hcur);
}

BOOL BubbleRichEditCtrl::TxScreenToClient(LPPOINT lppt)
{
	return ::ScreenToClient(hwnd, lppt);	
}

BOOL BubbleRichEditCtrl::TxClientToScreen(LPPOINT lppt)
{
	return ::ClientToScreen(hwnd, lppt);
}

HRESULT BubbleRichEditCtrl::TxActivate(LONG *plOldState)
{
    return S_OK;
}

HRESULT BubbleRichEditCtrl::TxDeactivate(LONG lNewState)
{
    return S_OK;
}
    

HRESULT BubbleRichEditCtrl::TxGetClientRect(LPRECT prc)
{
	*prc = rcClient;

	GetControlRect(prc);

	return NOERROR;
}


HRESULT BubbleRichEditCtrl::TxGetViewInset(LPRECT prc) 
{

    *prc = rcViewInset;
    
    return NOERROR;	
}

HRESULT BubbleRichEditCtrl::TxGetCharFormat(const CHARFORMATW **ppCF)
{
	*ppCF = &cf;
	return NOERROR;
}

HRESULT BubbleRichEditCtrl::TxGetParaFormat(const PARAFORMAT **ppPF)
{
	*ppPF = &pf;
	return NOERROR;
}


COLORREF BubbleRichEditCtrl::TxGetSysColor(int nIndex) 
{
	if (nIndex == COLOR_WINDOW)
	{
		if(!fNotSysBkgnd)
			return GetSysColor(COLOR_WINDOW);
		return crBackground;
	}

	return GetSysColor(nIndex);
}



HRESULT BubbleRichEditCtrl::TxGetBackStyle(TXTBACKSTYLE *pstyle)
{
	*pstyle = !fTransparent ? TXTBACK_OPAQUE : TXTBACK_TRANSPARENT;
	return NOERROR;
}


HRESULT BubbleRichEditCtrl::TxGetMaxLength(DWORD *pLength)
{
	*pLength = cchTextMost;
	return NOERROR;
}



HRESULT BubbleRichEditCtrl::TxGetScrollBars(DWORD *pdwScrollBar)
{
	*pdwScrollBar =  mStyle & (WS_VSCROLL | WS_HSCROLL | ES_AUTOVSCROLL | 
						ES_AUTOHSCROLL | ES_DISABLENOSCROLL);

	return NOERROR;
}


HRESULT BubbleRichEditCtrl::TxGetPasswordChar(TCHAR *pch)
{
#ifdef UNICODE
	*pch = chPasswordChar;
#else
	WideCharToMultiByte(CP_ACP, 0, &chPasswordChar, 1, pch, 1, NULL, NULL) ;
#endif
	return NOERROR;
}

HRESULT BubbleRichEditCtrl::TxGetAcceleratorPos(LONG *pcp)
{
	*pcp = laccelpos;
	return S_OK;
} 										   

HRESULT BubbleRichEditCtrl::OnTxCharFormatChange(const CHARFORMATW *pcf)
{
	return S_OK;
}


HRESULT BubbleRichEditCtrl::OnTxParaFormatChange(const PARAFORMAT *ppf)
{
	pf = *ppf;
	return S_OK;
}


HRESULT BubbleRichEditCtrl::TxGetPropertyBits(DWORD dwMask, DWORD *pdwBits) 
{
	DWORD dwProperties = 0;

	if (fRich)
	{
		dwProperties = TXTBIT_RICHTEXT;
	}

	if (mStyle & ES_MULTILINE)
	{
		dwProperties |= TXTBIT_MULTILINE;
	}

	if (mStyle & ES_READONLY)
	{
		dwProperties |= TXTBIT_READONLY;
	}


	if (mStyle & ES_PASSWORD)
	{
		dwProperties |= TXTBIT_USEPASSWORD;
	}

	if (!(mStyle & ES_NOHIDESEL))
	{
		dwProperties |= TXTBIT_HIDESELECTION;
	}

	if (fEnableAutoWordSel)
	{
		dwProperties |= TXTBIT_AUTOWORDSEL;
	}

	if (fVertical)
	{
		dwProperties |= TXTBIT_VERTICAL;
	}
					
	if (fWordWrap)
	{
		dwProperties |= TXTBIT_WORDWRAP;
	}

	if (fAllowBeep)
	{
		dwProperties |= TXTBIT_ALLOWBEEP;
	}

	if (fSaveSelection)
	{
		dwProperties |= TXTBIT_SAVESELECTION;
	}

	*pdwBits = dwProperties & dwMask; 
	return NOERROR;
}


HRESULT BubbleRichEditCtrl::TxNotify(DWORD iNotify, void *pv)
{
	if( iNotify == EN_REQUESTRESIZE )
	{
		RECT rc;
		REQRESIZE *preqsz = (REQRESIZE *)pv;
		
		GetControlRect(&rc);
		rc.bottom = rc.top + preqsz->rc.bottom + HOST_BORDER;
		rc.right  = rc.left + preqsz->rc.right + HOST_BORDER;
		rc.top -= HOST_BORDER;
		rc.left -= HOST_BORDER;
		
		SetClientRect(&rc, TRUE);
		
		return S_OK;
	} 

	// Forward this to the container
	if (pnc)
	{
		(*pnc)(iNotify);
	}

	return S_OK;
}



HRESULT BubbleRichEditCtrl::TxGetExtent(LPSIZEL lpExtent)
{

	// Calculate the length & convert to himetric
	*lpExtent = sizelExtent;

	return S_OK;
}

HRESULT	BubbleRichEditCtrl::TxGetSelectionBarWidth (LONG *plSelBarWidth)
{
	*plSelBarWidth = lSelBarWidth;
	return S_OK;
}


BOOL BubbleRichEditCtrl::GetReadOnly()
{
	return (mStyle & ES_READONLY) != 0;
}

void BubbleRichEditCtrl::SetReadOnly(BOOL fReadOnly)
{
	if (fReadOnly)
	{
		mStyle |= ES_READONLY;
	}
	else
	{
		mStyle &= ~ES_READONLY;
	}

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_READONLY, 
		fReadOnly ? TXTBIT_READONLY : 0);
}

BOOL BubbleRichEditCtrl::GetAllowBeep()
{
	return fAllowBeep;
}

void BubbleRichEditCtrl::SetAllowBeep(BOOL fAllowBeep)
{
	fAllowBeep = fAllowBeep;

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_ALLOWBEEP, 
		fAllowBeep ? TXTBIT_ALLOWBEEP : 0);
}

void BubbleRichEditCtrl::SetViewInset(RECT *prc)
{
	rcViewInset = *prc;

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_VIEWINSETCHANGE, 0);
}

WORD BubbleRichEditCtrl::GetDefaultAlign()
{
	return pf.wAlignment;
}


void BubbleRichEditCtrl::SetDefaultAlign(WORD wNewAlign)
{
	pf.wAlignment = wNewAlign;

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_PARAFORMATCHANGE, 0);
}

BOOL BubbleRichEditCtrl::GetRichTextFlag()
{
	return fRich;
}

void BubbleRichEditCtrl::SetRichTextFlag(BOOL fNew)
{
	fRich = fNew;

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_RICHTEXT, 
		fNew ? TXTBIT_RICHTEXT : 0);
}

LONG BubbleRichEditCtrl::GetDefaultLeftIndent()
{
	return pf.dxOffset;
}


void BubbleRichEditCtrl::SetDefaultLeftIndent(LONG lNewIndent)
{
	pf.dxOffset = lNewIndent;

	// Notify control of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_PARAFORMATCHANGE, 0);
}

void BubbleRichEditCtrl::SetClientRect(RECT *prc, BOOL fUpdateExtent) 
{
	// If the extent matches the client rect then we assume the extent should follow
	// the client rect.
	LONG lTestExt = DYtoHimetricY(
		(rcClient.bottom - rcClient.top)  - 2 * HOST_BORDER, yPerInch);

	if (fUpdateExtent 
		&& (sizelExtent.cy == lTestExt))
	{
		sizelExtent.cy = DXtoHimetricX((prc->bottom - prc->top) - 2 * HOST_BORDER, 
			xPerInch);
		sizelExtent.cx = DYtoHimetricY((prc->right - prc->left) - 2 * HOST_BORDER,
			yPerInch);
	}

	rcClient = *prc; 
}

BOOL BubbleRichEditCtrl::SetSaveSelection(BOOL f_SaveSelection)
{
	BOOL fResult = f_SaveSelection;

	fSaveSelection = f_SaveSelection;

	// notify text services of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_SAVESELECTION, 
		fSaveSelection ? TXTBIT_SAVESELECTION : 0);

	return fResult;		
}

HRESULT	BubbleRichEditCtrl::OnTxInPlaceDeactivate()
{
	HRESULT hr = pserv->OnTxInPlaceDeactivate();

	if (SUCCEEDED(hr))
	{
		fInplaceActive = FALSE;
	}

	return hr;
}

HRESULT	BubbleRichEditCtrl::OnTxInPlaceActivate(LPCRECT prcClient)
{
	fInplaceActive = TRUE;

	HRESULT hr = pserv->OnTxInPlaceActivate(prcClient);

	if (FAILED(hr))
	{
		fInplaceActive = FALSE;
	}

	return hr;
}

BOOL BubbleRichEditCtrl::DoSetCursor(RECT *prc, POINT *pt)
{
	RECT rc = prc ? *prc : rcClient;

	// Give some space for our border
	rc.top += HOST_BORDER;
	rc.bottom -= HOST_BORDER;
	rc.left += HOST_BORDER;
	rc.right -= HOST_BORDER;

	// Is this in our rectangle?
	if (PtInRect(&rc, *pt))
	{
		RECT *prcClient = (!fInplaceActive || prc) ? &rc : NULL;

		HDC hdc = GetDC(hwnd);

		pserv->OnTxSetCursor(
			DVASPECT_CONTENT,	
			-1,
			NULL,
			NULL,
			hdc,
			NULL,
			prcClient,
			pt->x, 
			pt->y);

		ReleaseDC(hwnd, hdc);

		return TRUE;
	}

	return FALSE;
}

void BubbleRichEditCtrl::GetControlRect(
	LPRECT prc			//@parm	Where to put client coordinates
)
{
	// Give some space for our border
	prc->top = rcClient.top + HOST_BORDER;
	prc->bottom = rcClient.bottom - HOST_BORDER;
	prc->left = rcClient.left + HOST_BORDER;
	prc->right = rcClient.right - HOST_BORDER;
}

void BubbleRichEditCtrl::SetTransparent(BOOL f_Transparent)
{
	fTransparent = f_Transparent;

	// notify text services of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_BACKSTYLECHANGE, 0);
}

LONG BubbleRichEditCtrl::SetAccelPos(LONG l_accelpos)
{
	LONG laccelposOld = l_accelpos;

	laccelpos = l_accelpos;

	// notify text services of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_SHOWACCELERATOR, 0);

	return laccelposOld;
}

WCHAR BubbleRichEditCtrl::SetPasswordChar(WCHAR ch_PasswordChar)
{
	WCHAR chOldPasswordChar = chPasswordChar;

	chPasswordChar = ch_PasswordChar;

	// notify text services of property change
	pserv->OnTxPropertyBitsChange(TXTBIT_USEPASSWORD, 
		(chPasswordChar != 0) ? TXTBIT_USEPASSWORD : 0);

	return chOldPasswordChar;
}

void BubbleRichEditCtrl::SetDisabled(BOOL fOn)
{
	cf.dwMask	  |= CFM_COLOR	   | CFM_DISABLED;
	cf.dwEffects |= CFE_AUTOCOLOR | CFE_DISABLED;

	if( !fOn )
	{
		cf.dwEffects &= ~CFE_DISABLED;
	}
	
	pserv->OnTxPropertyBitsChange(TXTBIT_CHARFORMATCHANGE, 
		TXTBIT_CHARFORMATCHANGE);
}

LONG BubbleRichEditCtrl::SetSelBarWidth(LONG l_SelBarWidth)
{
	LONG lOldSelBarWidth = lSelBarWidth;

	lSelBarWidth = l_SelBarWidth;

	if (lSelBarWidth)
	{
		mStyle |= ES_SELECTIONBAR;
	}
	else
	{
		mStyle &= (~ES_SELECTIONBAR);
	}

	pserv->OnTxPropertyBitsChange(TXTBIT_SELBARCHANGE, TXTBIT_SELBARCHANGE);

	return lOldSelBarWidth;
}

BOOL BubbleRichEditCtrl::GetTimerState()
{
	return fTimer;
}

void BubbleRichEditCtrl::OutPutText(const std::wstring& msg, bool bMine)
{
	//mMessageVec.push_back();
	formatText ft;
	ft.message = msg;

	std::wstring testmsg = msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	testmsg += msg;
	
	int nStart;
	int nEnd;
	SetSel(-1, -1);
	if (mMessageVec.size())
	{
		AppendText(L"\n");
	}
	
	AppendText(L"userName\n");

	GetSel(nStart, nEnd);
	ft.nStart = nEnd;
	AppendText(testmsg);
	SetSel(-1, -1);
	GetSel(nStart, nEnd);
	ft.nEnd = nEnd - 1;
	SetSel(ft.nStart, ft.nEnd + 1);
	long min;
	long max;
	long page;
	long pos;
	BOOL f;
	pserv->TxGetVScroll(&min, &max, &pos, &page, &f);
	long lx, ly;
	long tx, ty;
	long rx, ry;
	long bx, by;

	

	size_t fpos = testmsg.find('\n');
	size_t ls = 0;
	size_t le = testmsg.size();
	size_t lastPos = 0;
	
	while(fpos != std::wstring::npos)
	{
		if (fpos - lastPos > le - ls)
		{
			ls = lastPos;
			le = fpos;
		}
		lastPos = fpos + 1;
		fpos = testmsg.find('\n', lastPos);
	}

	CComPtr<ITextRange> rangeLong;
	pDocument->Range(ft.nStart + ls, ft.nStart + le, &rangeLong);
	long sl, st, el, et;
	HRESULT hr;
	if (FAILED(rangeLong->GetPoint(tomStart + TA_LEFT + TA_TOP, &sl, &st)))
	{
		__asm int 3
	}

	if (FAILED(rangeLong->GetPoint(tomEnd + TA_RIGHT + TA_TOP, &el, &et)))
	{
		__asm int 3
	}

	// one line 
	if (st == et)
	{
		ft.nMaxWidth = el - sl;
	}
	else
	{
		ft.nMaxWidth = mMaxREWidth;
	}

	{
		CComPtr<ITextRange> range;
		pDocument->Range(ft.nStart, ft.nEnd + 1, &range);
		HRESULT hr;
		long lx, ly;
		long tx, ty;
		long rx, ry;
		long bx, by;
		long sx, sy;
		long ex, ey;

		hr = range->GetPoint(tomStart + TA_LEFT + TA_TOP, &lx, &ly);
		hr = range->GetPoint(tomEnd + TA_RIGHT + TA_BOTTOM, &tx, &ty);
		hr = range->GetPoint(tomStart + TA_LEFT + TA_BASELINE, &rx, &ry);
		//hr = range->GetPoint(TA_BOTTOM, &bx, &by);

		//hr = range->GetPoint(tomStart, &sx, &sy);
		//hr = range->GetPoint(tomEnd, &ex, &ey);

		long min;
		long max;
		long page;
		long pos;
		BOOL f;

		RECT rt;
		GetWindowRect(hwnd,  &rt);

		rcClient;
		pserv->TxGetVScroll(&min, &max, &pos, &page, &f);
		ft.rect.top = pos + ly - rt.top;
		ft.rect.bottom = pos + ty - rt.top;
		ft.rect.left = lx - rt.left;
		ft.rect.right = ft.rect.left  + ft.nMaxWidth;
		int n = 0;
	}

	mMessageVec.push_back(ft);
}


/////////////////////////////////////////////////////////////
void BubbleRichEditCtrl::SetSel(int nStart, int nEnd)
{
	HRESULT hr = pserv->TxSendMessage(EM_SETSEL, nStart, nEnd, 0);
}

void BubbleRichEditCtrl::GetSel(int& nStart, int& nEnd)
{
	HRESULT hr = pserv->TxSendMessage(EM_GETSEL, (WPARAM)&nStart, (LPARAM)&nEnd, 0);
}

void BubbleRichEditCtrl::AppendText(const std::wstring& txt)
{
	SetSel(-1, -1);

	SETTEXTEX tx;
	tx.flags = ST_SELECTION;
	tx.codepage = 1200;

	HRESULT hr;
	hr = pserv->TxSendMessage(EM_SETTEXTEX, (WPARAM)&tx, (LPARAM)txt.c_str(), 0);
}

void BubbleRichEditCtrl::DrawBubbleToBg(HDC hdc)
{
	RECT rt = {0,0,1000,1000};
	FillRect(mMemDC, &rt, (HBRUSH)CreateSolidBrush(RGB(110,0,0)));
	if (!mMessageVec.size())
	{
		return ;
	}

	long min;
	long max;
	long page;
	long pos;
	BOOL f;

	pserv->TxGetVScroll(&min, &max, &pos, &page, &f);


	int index = 0;
	long noDisplayH = 0;
	for (int i = 0; i < mMessageVec.size(); i++)
	{
		if(pos < mMessageVec[i].rect.top)
		{
			index = i;
			break;
		}

	}


	long basetop = mMessageVec[index].rect.top;
	if (!index)
	{
		basetop = 0;
	}
	else
	{
		index -= 1;
		basetop = mMessageVec[index].rect.top;
	}


	for (int i = index; i < mMessageVec.size(); i++)
	{
		RECT rt = mMessageVec[i].rect;
		rt.top = rt.top - basetop;
		rt.bottom = rt.bottom - basetop + 2;
		//FrameRect(mMemDC, &rt, (HBRUSH)CreateSolidBrush(RGB(255,255,255)));

		HBRUSH hbrush = ::CreateSolidBrush( RGB(170,170,170) );
		//HBRUSH hbrushold = (HBRUSH )SelectObject( mMemDC, hbrush );

		FillRect( mMemDC, &rt, hbrush );
	}

	int w = rcClient.right - rcClient.left;
	int h = rcClient.bottom - rcClient.top;
	BitBlt(hdc, mRERect.left, mRERect.top, mRERect.Width(), mRERect.Height(), mMemDC, 0, pos - basetop, SRCCOPY);
}

void BubbleRichEditCtrl::drawOther( HDC hdc )
{
	//RECT rt = {0,0,1000,1000};
	//FillRect(mMemDC, &rt, (HBRUSH)CreateSolidBrush(RGB(110,0,0)));
	//hdc = GetDC(NULL);
	//BitBlt(hdc, 720, 0, 700, 500, mMemDC, 0, 0, SRCCOPY);
	//ReleaseDC(NULL, hdc);
}

BOOL BubbleRichEditCtrl::Attach(HWND parent, RECT* rect, DWORD mStyle)
{
	if (!IsWindow(parent))
	{
		return FALSE;
	}

	mHwnd = parent;
	GetWindowRect(parent, &mBgRect);
	mRERect = *rect;
	mStyle = GetWindowLong(parent, GWL_STYLE);

	GdiSetBatchLimit(1);
	Init();
	return TRUE;
}

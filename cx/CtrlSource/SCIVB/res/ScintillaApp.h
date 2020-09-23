#if !defined(AFX_SCINTILLAAPP_H__4A2DFF67_C146_11D3_9CDF_482471000000__INCLUDED_)
#define AFX_SCINTILLAAPP_H__4A2DFF67_C146_11D3_9CDF_482471000000__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// ScintillaApp.h : main header file for SCINTILLACTRL.OCX

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CScintillaApp : See Scintilla.cpp for implementation.

class CScintillaApp : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
private:
	HINSTANCE hSciLib;
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SCINTILLAAPP_H__4A2DFF67_C146_11D3_9CDF_482471000000__INCLUDED)

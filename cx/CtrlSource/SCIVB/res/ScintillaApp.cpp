// Scintilla.cpp : Implementation of CScintillaApp and DLL registration.

#include "stdafx.h"
#include "ScintillaApp.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


CScintillaApp NEAR theApp;

const GUID CDECL BASED_CODE _tlid =
		{ 0x4a2dff5f, 0xc146, 0x11d3, { 0x9c, 0xdf, 0x48, 0x24, 0x71, 0, 0, 0 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;


////////////////////////////////////////////////////////////////////////////
// CScintillaApp::InitInstance - DLL initialization

BOOL CScintillaApp::InitInstance()
{
	hSciLib = LoadLibrary(_T("SciLexer.dll"));
	BOOL bInit = COleControlModule::InitInstance();
	return bInit;
}


////////////////////////////////////////////////////////////////////////////
// CScintillaApp::ExitInstance - DLL termination

int CScintillaApp::ExitInstance()
{
	FreeLibrary(hSciLib);
	return COleControlModule::ExitInstance();
}


/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}


/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}

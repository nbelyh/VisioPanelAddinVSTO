//	vTestAddonExe.cpp - Generated by the "Microsoft Visio 2010 SDK 
//	Add-on Wizard"
//	
//	<summary>
//	This file contains the Visio add-on classes definitions.
//	</summary>
//	

#include "stdafx.h"

#include <tchar.h>
#include <strsafe.h>
#include "resource.h"
#include "vTestAddonExe.h"	//	The wizard-generated subclass 
						//	of VAddon

//	Global add-on declaration. The name resource ID that 
//	you give the VTestAddonExe in its constructor is the
//	text which shows up in the Visio Tools|Macros menu.

VTestAddonExe g_vTestAddonExeAddon(
		VAO_AOATTS_ISACTION|VAO_AOATTS_HASABOUT,
		VAO_ENABLEALWAYS,
		0,
		0,
		NULL,IDS_ADDONNAME);

//	Constructor
//
//	<summary>
//	Constructor for new VTestAddonExe objects.
//	</summary>
//	
//	Return Value:
//	none
//

VTestAddonExe::VTestAddonExe(
		VAO_ATTS atts,
		VAO_ENABMASK enabMask,
		VAO_INVOKEMASK invokeOnMask,
		VAO_NOTIFYMASK notifyOnMask,
		LPCTSTR lpNameU,
		UINT uIDLocalName) : 
	VAddon(atts, 
		enabMask, 
		invokeOnMask, 
		notifyOnMask, 
		lpNameU,
		uIDLocalName)
{

}

//	Destructor
//
//	<summary>
//	Destructor for VTestAddonExe objects.
//	</summary>
//	
//	Return Value:
//	none
//

VTestAddonExe::~VTestAddonExe()
{

}

//	Run - This method is called when the add-on is run
//
//	<summary>
//	The Run method is pure virtual, with no implementation
//	in the parent class. This method *MUST* be overridden,
//	but it doesn't *HAVE* to do anything. (As you can see 
//	below, we're just returning VAORC_SUCCESS.)
//	The code in Run gets executed when the user chooses the
//	Visio Tools menu item corresponding to your addon. It 
//	also gets executed when a shape sheet cell formula 
//	RUNADDON or RUNADDONWARGS gets evaluated.
//	</summary>
//	
//	Return Value:
//	VAORC - a valid Visio VAORC return code (see vao.h 
//	for details).
//

VAORC VTestAddonExe::Run(
	LPVAOV2LSTRUCT pV2L)
{
	Visio::IVApplicationPtr app;
	
	if (SUCCEEDED(GetApp(pV2L, app)))
	{
		MessageBox(GetActiveWindow(),
		_T("TestAddonExe add-on generated by Microsoft Visio 2010 SDK ")\
		_T("Add-on Wizard"),
		_T("TestAddonExe"),
		MB_OK);
	}

	return VAORC_SUCCESS;
}

//	IsEnabled - This method is called to find out if the 
//				add-on is enabled.
//
//	<summary>
//	Depending on the EnableMask for the add-on this 
//	method is called to find out whether the add-on 
//	is currently enabled.  See vao.h for more details.
//	</summary>
//	
//	Return Value:
//	VAORC - a valid Visio VAORC return code (see vao.h 
//	for details).
//

VAORC VTestAddonExe::IsEnabled(
	LPVAOV2LSTRUCT pV2L)
{
	return VAddon::IsEnabled(pV2L);
}

//	About - Display an About dialog box.
//
//	<summary>
//	This method displays an About dialog box.
//	</summary>
//	
//	Return Value:
//	VAORC - a valid Visio VAORC return code (see vao.h 
//	for details).
//

VAORC VTestAddonExe::About(
	LPVAOV2LSTRUCT /* pV2L */)
{
	TCHAR szText[] = _T("TestAddonExe Add-on generated by Microsoft Visio 2010 ")\
			_T("SDK Add-on Wizard.\n");

	TCHAR szAbout[] = _T("About ");

	TCHAR szCaption[_MAX_PATH];
	StringCchCopy(szCaption, sizeof(szCaption)/sizeof(TCHAR), szAbout);
	StringCchCat(szCaption, sizeof(szCaption)/sizeof(TCHAR),  GetName());

	MessageBox(GetActiveWindow(), szText, szCaption, MB_OK);

	return VAORC_SUCCESS;
}

//	Help - Display help for the add-on.
//
//	<summary>
//	This method displays help for the add-on.
//	</summary>
//	
//	Return Value:
//	VAORC - a valid Visio VAORC return code (see vao.h
//	for details).
//

VAORC VTestAddonExe::Help(
	LPVAOV2LSTRUCT /* pV2L */)
{
	TCHAR szText[] = _T("Add code to jump to TestAddonExe's ")\
		_T("help file.");

	TCHAR szHelp[] = _T(" Help");

	TCHAR szCaption[_MAX_PATH];
	StringCchCopy(szCaption, sizeof(szCaption)/sizeof(TCHAR), GetName());
	StringCchCat(szCaption, sizeof(szCaption)/sizeof(TCHAR),  szHelp);

	MessageBox(GetActiveWindow(), szText, szCaption, MB_OK);

	return VAORC_SUCCESS;
}

//	Load - Load the add-on.
//
//	<summary>
//	This method is called when Visio loads the add-on.
//	</summary>
//	
//	Return Value:
//	VAORC - a valid Visio VAORC return code (see vao.h 
//	for details).
//

VAORC VTestAddonExe::Load(
	WORD wVersion, 
	LPVOID p)
{
	return VAddon::Load(wVersion, p);
}

//	Unload - Unload the add-on.
//
//	<summary>
//	This method is called when it's time for an add-on to 
//	unload. This is the last opportunity to call automation
//	methods in Visio.
//	</summary>
//	
//	Return Value:
//	VAORC - a valid Visio VAORC return code (see vao.h 
//	for details).
//

VAORC VTestAddonExe::Unload(
	WORD wParam, 
	LPVOID p)
{
	m_app = NULL;
	return VAddon::Unload(wParam, p);
}

//	KillSession - Notify the add-on that it is time to 
//	shutdown.
//
//	<summary>
//	This method is called when an add-on, that returned 
//	VAORC_L2V_MODELESS from its run method, needs to 
//	shutdown. This typically happens when a user closes 
//	Visio.
//	</summary>
//	
//	Return Value:
//	VAORC - a valid Visio VAORC return code (see vao.h 
//	for details).
//

VAORC VTestAddonExe::KillSession(
	LPVAOV2LSTRUCT pV2L)
{
	return VAddon::KillSession(pV2L);
}

//	GetInstance - Get the HINSTANCE for the add-on.
//
//	<summary>
//	This method returns a HINSTANCE to use as the module 
//	instance handle.
//	</summary>
//	
//	Return Value:
//	HINSTANCE
//

HINSTANCE VTestAddonExe::GetInstance(
	long nFlags /* = 0L*/)
{
	return VAddon::GetInstance(nFlags);
}

//	GetResourceHandle - Get the HINSTANCE for the add-on 
//	resources.
//
//	<summary>
//	This method returns a HINSTANCE to use as the resource 
//	handle for the add-on.
//	</summary>
//	
//	Return Value:
//	HINSTANCE
//

HINSTANCE VTestAddonExe::GetResourceHandle()
{
	return VAddon::GetResourceHandle();
}

//	GetApp - Get a Visio::IVApplicationPtr</vistypelib> for the 
//	Visio application.
//
//	<summary>
//	This method gets a Visio::IVApplicationPtr</vistypelib> for the Visio 
//	application.  It first tries to attach to the currently
//	running Visio application.  If that fails this method 
//	will launch a new Visio application.
//	</summary>
//	
//	Return Value:
//	HRESULT - S_OK or some failure code if no application 
//	object was returned.
//



HRESULT VTestAddonExe::GetApp(
	LPVAOV2LSTRUCT pV2L, 
	Visio::IVApplicationPtr &app)
{
	if (!m_app.GetInterfacePtr())
	{
		if (NULL != pV2L && NULL != pV2L->lpApp)
		{
			LPUNKNOWN lpApp = pV2L->lpApp;
			m_app = lpApp;
		}
		else
		{
			if (!SUCCEEDED(m_app.GetActiveObject(
				Visio::CLSID_Application)))
			{
				m_app.CreateInstance(
					Visio::CLSID_Application);
			}
		}
	}

	if (m_app.GetInterfacePtr())
	{
		app = m_app;
		return NOERROR;
	}

	return E_FAIL;
}
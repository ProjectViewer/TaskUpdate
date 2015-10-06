// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "SeavusTaskUpdate_i.h"
#include "dllmain.h"

CSeavusTaskUpdateModule _AtlModule;

void CSeavusTaskUpdateModule::setInstance(HINSTANCE hInstance)
{
	m_hInstance = hInstance;
}

HINSTANCE CSeavusTaskUpdateModule::getInstance() const
{
	return m_hInstance;
}

class CSeavusTaskUpdateApp : public CWinApp
{
public:

// Overrides
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CSeavusTaskUpdateApp, CWinApp)
END_MESSAGE_MAP()

CSeavusTaskUpdateApp theApp;

BOOL CSeavusTaskUpdateApp::InitInstance()
{

	_AtlModule.setInstance(m_hInstance);
	return _AtlModule.DllMain(DLL_THREAD_ATTACH, NULL); 
}

int CSeavusTaskUpdateApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
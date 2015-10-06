// Connect.cpp : Implementation of CConnect

#include "stdafx.h"
#include "Connect.h"
#include "Utils.h"
#include "CollDlg.h"
#include "GetAssgnUpdatesDlg.h"
#include "GetTaskUpdatesDlg.h"
#include "OptionsPS.h"
#include "UpdatesManager.h"

EXTERN_C IMAGE_DOS_HEADER __ImageBase;

_ATL_FUNC_INFO OnClickButtonInfo = {CC_STDCALL,VT_EMPTY,2,{VT_DISPATCH,VT_BYREF | VT_BOOL}};
_ATL_FUNC_INFO OnSimpleEventInfo ={CC_STDCALL,VT_EMPTY,0};

CConnect*		CConnect::pTheConnect = NULL;
HRESULT			CConnect::FinalConstruct()
{
	pTheConnect = this;
	return S_OK;
}
void			CConnect::FinalRelease()
{
	pTheConnect = NULL;
}
bool CConnect::IsInEditingMode()
{
 	bool returnValue = false;

	for(int i=1; i<=m_spToolBarCtrl->GetCount(); ++i)
	{
		CComPtr <Office::CommandBar> spTollbar  = m_spToolBarCtrl->GetItem(i);//standar toolbar
		if(spTollbar->GetId()==9)
		{
			CComPtr<Office::CommandBarControl> spControl = spTollbar->GetControls()->GetItem(1);
			returnValue = (spControl->GetEnabled()==VARIANT_FALSE);
		}
	}
	return returnValue;
}
void __stdcall	CConnect::OnToolbarUpdate()
{
	MSProject::ProjectsPtr pProjects = m_MSPApplication->Projects;

    int reportingMode = TUAssignmentReportingMode;
	BSTR bstrFullName = NULL;
    bool writable = false;

	try
	{

		bstrFullName = (m_MSPApplication->ActiveProject->FullName);
        reportingMode = CTU_Utils::GetReportingModeFromFile(bstrFullName);
        VARIANT_BOOL readonly = m_MSPApplication->ActiveProject->ReadOnly;
        short sReadOnly = readonly;
        writable = bstrFullName && sReadOnly == 0;
	}
	catch(...)
	{
		bstrFullName = NULL;
	}

    if((pProjects && pProjects->GetCount()>0) && !IsInEditingMode() && writable)
	{	
		if(bstrFullName != NULL)
		{
			bool bIdFileCollaborative = CTU_Utils::IsFileTUActive(bstrFullName);

			if(bIdFileCollaborative)
			{
				EnableButton(m_spButtonGetUpd,true);
				EnableButton(m_spMenuGetUpd,true);
                EnableButton(m_spButtonOptions,reportingMode == TUTaskReportingMode);
                EnableButton(m_spMenuOptions,reportingMode == TUTaskReportingMode);
			}
			else
			{
				EnableButton(m_spButtonGetUpd,false);
				EnableButton(m_spMenuGetUpd,false);
                EnableButton(m_spButtonOptions,false);
                EnableButton(m_spMenuOptions,false);
			}
		}
		else
		{
			EnableButton(m_spButtonGetUpd,false);
			EnableButton(m_spMenuGetUpd,false);
            EnableButton(m_spButtonOptions,false);
            EnableButton(m_spMenuOptions,false);
		}

		EnableButton(m_spButtonMakeColl,true);
		EnableButton(m_spMenuMakeColl,true);
        EnableButton(m_spButtonHelp,true);
        EnableButton(m_spMenuHelp,true);
	}
	else
	{
		EnableButton(m_spButtonGetUpd,false);
		EnableButton(m_spMenuGetUpd,false);
		EnableButton(m_spButtonMakeColl,false);
		EnableButton(m_spMenuMakeColl,false);
        EnableButton(m_spButtonOptions,false);
        EnableButton(m_spMenuOptions,false);
        EnableButton(m_spButtonHelp,true);
        EnableButton(m_spMenuHelp,true);
	}
}

void __stdcall	CConnect::OnClickButtonMakeColl(IDispatch* /*Office::_CommandBarButton* */ Ctrl, VARIANT_BOOL * CancelDefault)
{
    ShowMakeCollaborative();
}
void __stdcall	CConnect::OnClickButtonGetUpd(IDispatch* /*Office::_CommandBarButton* */ Ctrl, VARIANT_BOOL * CancelDefault)
{
    ShowGetUpdates();
}
void __stdcall	CConnect::OnClickButtonOptions(IDispatch* /*Office::_CommandBarButton* */ Ctrl, VARIANT_BOOL * CancelDefault)
{
    ShowOptions();
}
void __stdcall	CConnect::OnClickButtonHelp(IDispatch* /*Office::_CommandBarButton* */ Ctrl, VARIANT_BOOL * CancelDefault)
{
    ShowHelp();
}
void __stdcall	CConnect::OnClickMenuMakeColl(IDispatch* /*Office::_CommandBarButton* */ Ctrl, VARIANT_BOOL * CancelDefault)
{
    ShowMakeCollaborative();
}
void __stdcall	CConnect::OnClickMenuGetUpd(IDispatch* /*Office::_CommandBarButton* */ Ctrl, VARIANT_BOOL * CancelDefault)
{
    ShowGetUpdates();
}
void __stdcall	CConnect::OnClickMenuOptions(IDispatch* /*Office::_CommandBarButton* */ Ctrl, VARIANT_BOOL * CancelDefault)
{
    ShowOptions();
}
void __stdcall	CConnect::OnClickMenuHelp(IDispatch* /*Office::_CommandBarButton* */ Ctrl, VARIANT_BOOL * CancelDefault)
{
    ShowHelp();
}

void    CConnect::ShowMakeCollaborative()
{
    Utils::initAppUnits();

    AFX_MANAGE_STATE(AfxGetStaticModuleState());//call this before creating dlg object
    showDialog<CCollDlg>(0, true);
}
void    CConnect::ShowGetUpdates()
{
    Utils::initAppUnits();

    MSProject::ProjectsPtr pProjects = m_MSPApplication->Projects;
    BSTR path = pProjects->Application->ActiveProject->FullName;
    int reportingMode = CTU_Utils::GetReportingModeFromFile(path);
    CString strGetViaCF = CTU_Utils::GetCustomPropertyValue(path, _T("ImportViaCustomFields"));
    bool bImportViaCF = (strGetViaCF == _T("1")) ? true : false;

    if (TUTaskReportingMode == reportingMode)
    {
        if (bImportViaCF)
        {
			AFX_MANAGE_STATE(AfxGetStaticModuleState());//call this before creating dlg object
			CWaitCursor wait;
            CUpdatesManager updatesManager;
			wait.Restore();

            bool processNotes = true;
            bool processCustomTexts = true;
            bool hasNoteReport = false;
            bool hasCustomTextReport = false;
            updatesManager.getNotesViaGantt(processNotes, hasNoteReport);   
            updatesManager.getCustomTextUpdatesViaGantt(processCustomTexts, hasCustomTextReport); 

            if ((hasNoteReport && processNotes) || (hasCustomTextReport && processCustomTexts))
            {
                MessageBox(NULL,
                    _T("Notes/Custom Text Fields has been successfully inserted in the project plan."),
                    _T("SeavusTaskUpdate"),
                    MB_OK | MB_ICONWARNING);
            }

            updatesManager.getUpdatesViaGantt();   
        }
        else
        {
			AFX_MANAGE_STATE(AfxGetStaticModuleState());//call this before creating dlg object
			CWaitCursor wait;
            CUpdatesManager* updatesManager = new CUpdatesManager();
			wait.Restore();

            bool processNotes = true;
            bool processCustomTexts = true;
            bool hasNoteReport = false;
            bool hasCustomTextReport = false;
            updatesManager->getNotesViaGantt(processNotes, hasNoteReport);  
            updatesManager->getCustomTextUpdatesViaGantt(processCustomTexts, hasCustomTextReport); 

            if ((hasNoteReport && processNotes) || (hasCustomTextReport && processCustomTexts))
            {
                MessageBox(NULL,
                    _T("Notes/Custom Text Fields has been successfully inserted in the project plan."),
                    _T("SeavusTaskUpdate"),
                    MB_OK | MB_ICONWARNING);
            }

            //AFX_MANAGE_STATE(AfxGetStaticModuleState());//call this before creating dlg object
            showDialog<CGetTaskUpdatesDlg>(updatesManager, 0);
        } 
    }
    else
    {
        AFX_MANAGE_STATE(AfxGetStaticModuleState());//call this before creating dlg object
        showDialog<CGetAssgnUpdatesDlg>(0, 0);  
    }
}
void    CConnect::ShowOptions()
{
    AFX_MANAGE_STATE(AfxGetStaticModuleState());//call this before creating dlg object
    showPropertySheet<COptionsPS>(true);
}
void    CConnect::ShowHelp()
{
    LPTSTR  strDLLPath1 = new TCHAR[_MAX_PATH];
    ::GetModuleFileName((HINSTANCE)&__ImageBase, strDLLPath1, _MAX_PATH);

    CString path = strDLLPath1;
    path = path.Left(path.ReverseFind(_T('\\')));
    path += _T("//Help.chm");

    ShellExecute(NULL, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
}

template<class T> 
void CConnect::showDialog(CUpdatesManager* updatesManager, bool bAskToSave)
{
	MSProject::ProjectsPtr pProjects = m_MSPApplication->Projects;
	if(pProjects && pProjects->GetCount()==0)
		return;
    try 
    {
        MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
        MSProject::_IProjectDocPtr project = app->GetActiveProject();  

        if (project && !Utils::isProjectSaved(project) && bAskToSave)
        {
            if(AfxMessageBox(_T("The project plan must be saved before “Make Collaboration” functionality can be called. \n\nDo you like to save project now?"), MB_YESNO) == IDNO)
            {
                // do nothing 
            }
            else if (Utils::saveProject(project))
            {
                T dlg(0, 0);
                dlg.DoModal();
            }
        }
        else 
        {
			AFX_MANAGE_STATE(AfxGetStaticModuleState());//call this before creating dlg object
			CWaitCursor wait;
            T dlg(updatesManager, 0);
			wait.Restore();

            dlg.DoModal();
        }
    }
    catch( _com_error &e )
    {
         // Crack _com_error
        _bstr_t bstrSource(e.Source());
        _bstr_t bstrDescription(e.Description());

        TRACE( "Exception thrown for classes generated by #import" );
        TRACE( "\tCode = %08lx\n",      e.Error());
        TRACE( "\tCode meaning = %s\n", e.ErrorMessage());
        TRACE( "\tSource = %s\n",       (LPCTSTR) bstrSource);
        TRACE( "\tDescription = %s\n",  (LPCTSTR) bstrDescription);
    }
    catch (...)
    {
        // do nothing
    }
}
template<class T> 
void CConnect::showPropertySheet(bool bAskToSave)
{
	MSProject::ProjectsPtr pProjects = m_MSPApplication->Projects;
	if(pProjects && pProjects->GetCount()==0)
		return;
    try 
    {
        MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
        MSProject::_IProjectDocPtr project = app->GetActiveProject();  

        T dlg(_T("Task Update Options"));
        dlg.DoModal();

    }
    catch( _com_error &e )
    {
         // Crack _com_error
        _bstr_t bstrSource(e.Source());
        _bstr_t bstrDescription(e.Description());

        TRACE( "Exception thrown for classes generated by #import" );
        TRACE( "\tCode = %08lx\n",      e.Error());
        TRACE( "\tCode meaning = %s\n", e.ErrorMessage());
        TRACE( "\tSource = %s\n",       (LPCTSTR) bstrSource);
        TRACE( "\tDescription = %s\n",  (LPCTSTR) bstrDescription);
    }
    catch (...)
    {
        // do nothing
    }
}
void	    CConnect::EnableButton(Office::_CommandBarButton* pButton, bool bEnable)
{		
	if(!pButton)
		return;
	if(bEnable && pButton->GetEnabled())//if the button is enabled there is no need to enable it again
		return;
	if(!bEnable && !pButton->GetEnabled())//if the button is disabled there is no need to disable it again
		return;

	pButton->PutVisible(VARIANT_FALSE); //set to visible false, then again to true, to refresh the buttons painting
	pButton->PutEnabled(bEnable ? VARIANT_TRUE : VARIANT_FALSE);
	pButton->PutVisible(VARIANT_TRUE); 

}
void		CConnect::DeleteControls()
{
	if(m_spButtonMakeColl)
		m_spButtonMakeColl->Delete();

	if(m_spButtonGetUpd)
		m_spButtonGetUpd->Delete();

    if(m_spButtonOptions)
        m_spButtonOptions->Delete();

    if(m_spButtonHelp)
        m_spButtonHelp->Delete();

	if(m_spToolbar)
		m_spToolbar->Delete();

	if(m_spMenuMakeColl)
		m_spMenuMakeColl->Delete();

	if(m_spMenuGetUpd)
		m_spMenuGetUpd->Delete();

    if(m_spMenuOptions)
        m_spMenuOptions->Delete();

    if(m_spMenuHelp)
        m_spMenuHelp->Delete();

	if(m_spMenuCtrl)
		m_spMenuCtrl->Delete();
}
void		CConnect::EventsAdvise()
{
	CommandButtonMakeCollEvents::DispEventAdvise((IDispatch*)m_spButtonMakeColl);
	CommandButtonGetUpdEvents::DispEventAdvise((IDispatch*)m_spButtonGetUpd);
    CommandButtonOptionsEvents::DispEventAdvise((IDispatch*)m_spButtonOptions);
    CommandButtonHelpEvents::DispEventAdvise((IDispatch*)m_spButtonHelp);
 	CommandMenuMakeCollEvents::DispEventAdvise((IDispatch*)m_spMenuMakeColl);
 	CommandMenuGetUpdEvents::DispEventAdvise((IDispatch*)m_spMenuGetUpd);
    CommandMenuOptionsEvents::DispEventAdvise((IDispatch*)m_spMenuOptions);
    CommandMenuHelpEvents::DispEventAdvise((IDispatch*)m_spMenuHelp);
	CommandUpdateToolbarEvents::DispEventAdvise((IDispatch*)m_spToolBarCtrl,&__uuidof(Office::_CommandBarsEvents));
}
void		CConnect::EventsUnadvise()
{
	CommandButtonMakeCollEvents::DispEventUnadvise((IDispatch*)m_spButtonMakeColl);
	CommandButtonGetUpdEvents::DispEventUnadvise((IDispatch*)m_spButtonGetUpd);
    CommandButtonOptionsEvents::DispEventUnadvise((IDispatch*)m_spButtonOptions);
    CommandButtonHelpEvents::DispEventUnadvise((IDispatch*)m_spButtonHelp);
	CommandMenuMakeCollEvents::DispEventUnadvise((IDispatch*)m_spMenuMakeColl);;
	CommandMenuGetUpdEvents::DispEventUnadvise((IDispatch*)m_spMenuGetUpd);
    CommandMenuOptionsEvents::DispEventUnadvise((IDispatch*)m_spMenuOptions);
    CommandMenuHelpEvents::DispEventUnadvise((IDispatch*)m_spMenuHelp);
	CommandUpdateToolbarEvents::DispEventUnadvise((IDispatch*)m_spToolBarCtrl,&__uuidof(Office::_CommandBarsEvents));
}
STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode /*ConnectMode*/, IDispatch *pAddInInst, SAFEARRAY ** custom)
{

	//clear toolbar and menu if they exist before connecting

	DeleteControls();

	HRESULT hr;
	pApplication->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pApplication);
	pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pAddInInstance);
	{
		CComQIPtr <MSProject::_MSProject> spApp(pApplication);
		m_MSPApplication = spApp;
	}

    _bstr_t appVersion = m_MSPApplication->Version;
    int iMajorVersion = Utils::GetProjectMajorVersion((BSTR)appVersion);

	eMSProjectVersion mspVersion(MSP_Unknown);
	if(iMajorVersion == 14)
		mspVersion = MSP_14;
	else if(iMajorVersion == 12)
		mspVersion = MSP_12;
	else 
		mspVersion = MSP_11;

	//ADDING TOOLBAR WITH 4 BUTTONs
	CComPtr < Office::CommandBar> spCmdBar;
	m_spToolBarCtrl = m_MSPApplication->GetCommandBars();

	CComVariant vName("Seavus Task Update");
	CComVariant vPos(1); 
	CComVariant vTemp(VARIANT_TRUE);     
	CComVariant vEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);            

	m_spToolbar = m_spToolBarCtrl->Add(vName, vPos, vEmpty, vTemp);

	CComPtr < Office::CommandBarControls> spBarControls;
	spBarControls = m_spToolbar->GetControls();

	CComVariant vToolbarBtnType(1/*Office::MsoControlType::msoControlButton*/);
	CComVariant vShow(VARIANT_TRUE);
	CComPtr < Office::CommandBarControl> spNewBarMakeColl; 
	CComPtr < Office::CommandBarControl> spNewBarGetUpd; 
    CComPtr < Office::CommandBarControl> spNewBarOptions; 
    CComPtr < Office::CommandBarControl> spNewBarHelp; 

	spNewBarMakeColl = spBarControls->Add(vToolbarBtnType,vEmpty,vEmpty,vEmpty,vShow); 
	spNewBarGetUpd = spBarControls->Add(vToolbarBtnType,vEmpty,vEmpty,vEmpty,vShow); 
    spNewBarOptions = spBarControls->Add(vToolbarBtnType,vEmpty,vEmpty,vEmpty,vShow); 
    spNewBarHelp = spBarControls->Add(vToolbarBtnType,vEmpty,vEmpty,vEmpty,vShow); 

	CComQIPtr < Office::_CommandBarButton> spCmdButtonMakeColl(spNewBarMakeColl);
	CComQIPtr < Office::_CommandBarButton> spCmdButtonGetUpd(spNewBarGetUpd);
    CComQIPtr < Office::_CommandBarButton> spCmdButtonOptions(spNewBarOptions);
    CComQIPtr < Office::_CommandBarButton> spCmdButtonHelp(spNewBarHelp);

    //make collaborative
	HBITMAP hBmpMakeColl =(HBITMAP)::LoadImage(_AtlModule.getInstance(),MAKEINTRESOURCE(IDB_MakeColl),IMAGE_BITMAP,0,0,LR_LOADMAP3DCOLORS);
	::OpenClipboard(NULL);
	::EmptyClipboard();
	::SetClipboardData(CF_BITMAP, (HANDLE)hBmpMakeColl);
	::CloseClipboard();
	::DeleteObject(hBmpMakeColl);     

	PICTDESC pdescMakeColl;
	memset(&pdescMakeColl, 0, sizeof(pdescMakeColl));
	pdescMakeColl.cbSizeofstruct = sizeof(pdescMakeColl);
	pdescMakeColl.picType = PICTYPE_BITMAP;
	pdescMakeColl.bmp.hbitmap = hBmpMakeColl;
	PicturePtr pdMakeColl ;
	OleCreatePictureIndirect(&pdescMakeColl, IID_IPictureDisp, TRUE,(void**)&pdMakeColl);

	spCmdButtonMakeColl->PutStyle(mspVersion==MSP_14 ? Office::msoButtonIconAndCaption : Office::msoButtonIcon);
	spCmdButtonMakeColl->PutPicture(pdMakeColl);

	HBITMAP hBmpMakeCollMask =(HBITMAP)::LoadImage(_AtlModule.getInstance(),MAKEINTRESOURCE(IDB_MakeCollMask),IMAGE_BITMAP,0,0,LR_DEFAULTCOLOR);
	::OpenClipboard(NULL);
	::EmptyClipboard();
	::SetClipboardData(CF_BITMAP, (HANDLE)hBmpMakeCollMask);
	::CloseClipboard();
	::DeleteObject(hBmpMakeCollMask);     
 
	PICTDESC pdescMakeCollMask;
	memset(&pdescMakeCollMask, 0, sizeof(pdescMakeCollMask));
	pdescMakeCollMask.cbSizeofstruct = sizeof(pdescMakeCollMask);
	pdescMakeCollMask.picType = PICTYPE_BITMAP;
	pdescMakeCollMask.bmp.hbitmap = hBmpMakeCollMask;
	CComPtr<IPictureDisp> pdMakeCollMask ;
	OleCreatePictureIndirect(&pdescMakeCollMask, IID_IPictureDisp, TRUE,(void**)&pdMakeCollMask);

	//spCmdButtonMakeColl->PutMask(pdMakeCollMask);
	spCmdButtonMakeColl->PutVisible(VARIANT_TRUE); 
	spCmdButtonMakeColl->PutCaption(OLESTR("Make Collaborative")); 
	spCmdButtonMakeColl->PutEnabled(VARIANT_TRUE);
	spCmdButtonMakeColl->PutTooltipText(OLESTR("Make Project File(s) Collaborative")); 
	spCmdButtonMakeColl->PutTag(OLESTR("Tag for Make Collaborative")); 

    //get updtaes
	HBITMAP hBmpGetUpd =(HBITMAP)::LoadImage(_AtlModule.getInstance(),MAKEINTRESOURCE(IDB_GetUpd),IMAGE_BITMAP,16,16,LR_DEFAULTCOLOR);
	::OpenClipboard(NULL);
	::EmptyClipboard();
	::SetClipboardData(CF_BITMAP, (HANDLE)hBmpGetUpd);
	::CloseClipboard();
	::DeleteObject(hBmpGetUpd);  

	PICTDESC pdescGetUpd;
	memset(&pdescGetUpd, 0, sizeof(pdescGetUpd));
	pdescGetUpd.cbSizeofstruct = sizeof(pdescGetUpd);
	pdescGetUpd.picType = PICTYPE_BITMAP;
	pdescGetUpd.bmp.hbitmap = hBmpGetUpd;
	PicturePtr pdGetUpd ;
	OleCreatePictureIndirect(&pdescGetUpd, IID_IPictureDisp, TRUE,(void**)&pdGetUpd);

	spCmdButtonGetUpd->PutStyle(mspVersion==MSP_14 ? Office::msoButtonIconAndCaption : Office::msoButtonIcon);
	spCmdButtonGetUpd->PutPicture(pdGetUpd);

	HBITMAP hBmpGetUpdMask =(HBITMAP)::LoadImage(_AtlModule.getInstance(),MAKEINTRESOURCE(IDB_GetUpdMask),IMAGE_BITMAP,16,16,LR_DEFAULTCOLOR);
	::OpenClipboard(NULL);
	::EmptyClipboard();
	::SetClipboardData(CF_BITMAP, (HANDLE)hBmpGetUpdMask);
	::CloseClipboard();
	::DeleteObject(hBmpGetUpdMask); 

	PICTDESC pdescGetUpdMask;
	memset(&pdescGetUpdMask, 0, sizeof(pdescGetUpdMask));
	pdescGetUpdMask.cbSizeofstruct = sizeof(pdescGetUpdMask);
	pdescGetUpdMask.picType = PICTYPE_BITMAP;
	pdescGetUpdMask.bmp.hbitmap = hBmpGetUpdMask;
	CComPtr<IPictureDisp> pdGetUpdMask ;
	OleCreatePictureIndirect(&pdescGetUpdMask, IID_IPictureDisp, TRUE,(void**)&pdGetUpdMask);

	//spCmdButtonGetUpd->PutMask(pdGetUpdMask);
	spCmdButtonGetUpd->PutVisible(VARIANT_TRUE); 
	spCmdButtonGetUpd->PutCaption(OLESTR("Get Updates")); 
	spCmdButtonGetUpd->PutEnabled(VARIANT_TRUE);
	spCmdButtonGetUpd->PutTooltipText(OLESTR("Get Updates")); 
	spCmdButtonGetUpd->PutTag(OLESTR("Tag for Get Updates")); 

    //options
    HBITMAP hBmpOptions =(HBITMAP)::LoadImage(_AtlModule.getInstance(),MAKEINTRESOURCE(IDB_Options),IMAGE_BITMAP,0,0,LR_LOADMAP3DCOLORS);
    ::OpenClipboard(NULL);
    ::EmptyClipboard();
    ::SetClipboardData(CF_BITMAP, (HANDLE)hBmpOptions);
    ::CloseClipboard();
    ::DeleteObject(hBmpOptions);     

    PICTDESC pdescOptions;
    memset(&pdescOptions, 0, sizeof(pdescOptions));
    pdescOptions.cbSizeofstruct = sizeof(pdescOptions);
    pdescOptions.picType = PICTYPE_BITMAP;
    pdescOptions.bmp.hbitmap = hBmpOptions;
    PicturePtr pdOptions ;
    OleCreatePictureIndirect(&pdescOptions, IID_IPictureDisp, TRUE,(void**)&pdOptions);

    spCmdButtonOptions->PutStyle(mspVersion==MSP_14 ? Office::msoButtonIconAndCaption : Office::msoButtonIcon);
    spCmdButtonOptions->PutPicture(pdOptions);

    HBITMAP hBmpOptionsMask =(HBITMAP)::LoadImage(_AtlModule.getInstance(),MAKEINTRESOURCE(IDB_OptionsMask),IMAGE_BITMAP,0,0,LR_DEFAULTCOLOR);
    ::OpenClipboard(NULL);
    ::EmptyClipboard();
    ::SetClipboardData(CF_BITMAP, (HANDLE)hBmpOptionsMask);
    ::CloseClipboard();
    ::DeleteObject(hBmpOptionsMask);     

    PICTDESC pdescOptionsMask;
    memset(&pdescOptionsMask, 0, sizeof(pdescOptionsMask));
    pdescOptionsMask.cbSizeofstruct = sizeof(pdescOptionsMask);
    pdescOptionsMask.picType = PICTYPE_BITMAP;
    pdescOptionsMask.bmp.hbitmap = hBmpOptionsMask;
    CComPtr<IPictureDisp> pdOptionsMask ;
    OleCreatePictureIndirect(&pdescOptionsMask, IID_IPictureDisp, TRUE,(void**)&pdOptionsMask);

    //spCmdButtonOptions->PutMask(pdOptionsMask);
    spCmdButtonOptions->PutVisible(VARIANT_TRUE); 
    spCmdButtonOptions->PutCaption(OLESTR("Options")); 
    spCmdButtonOptions->PutEnabled(VARIANT_TRUE);
    spCmdButtonOptions->PutTooltipText(OLESTR("Options")); 
    spCmdButtonOptions->PutTag(OLESTR("Tag for Options")); 

    //help
    HBITMAP hBmpHelp =(HBITMAP)::LoadImage(_AtlModule.getInstance(),MAKEINTRESOURCE(IDB_Help),IMAGE_BITMAP,0,0,LR_LOADMAP3DCOLORS);
    ::OpenClipboard(NULL);
    ::EmptyClipboard();
    ::SetClipboardData(CF_BITMAP, (HANDLE)hBmpHelp);
    ::CloseClipboard();
    ::DeleteObject(hBmpHelp);     

    PICTDESC pdescHelp;
    memset(&pdescHelp, 0, sizeof(pdescHelp));
    pdescHelp.cbSizeofstruct = sizeof(pdescHelp);
    pdescHelp.picType = PICTYPE_BITMAP;
    pdescHelp.bmp.hbitmap = hBmpHelp;
    PicturePtr pdHelp ;
    OleCreatePictureIndirect(&pdescHelp, IID_IPictureDisp, TRUE,(void**)&pdHelp);

    spCmdButtonHelp->PutStyle(mspVersion==MSP_14 ? Office::msoButtonIconAndCaption : Office::msoButtonIcon);
    spCmdButtonHelp->PutPicture(pdHelp);

    HBITMAP hBmpHelpMask =(HBITMAP)::LoadImage(_AtlModule.getInstance(),MAKEINTRESOURCE(IDB_HelpMask),IMAGE_BITMAP,0,0,LR_DEFAULTCOLOR);
    ::OpenClipboard(NULL);
    ::EmptyClipboard();
    ::SetClipboardData(CF_BITMAP, (HANDLE)hBmpHelpMask);
    ::CloseClipboard();
    ::DeleteObject(hBmpHelpMask);     

    PICTDESC pdescHelpMask;
    memset(&pdescHelpMask, 0, sizeof(pdescHelpMask));
    pdescHelpMask.cbSizeofstruct = sizeof(pdescHelpMask);
    pdescHelpMask.picType = PICTYPE_BITMAP;
    pdescHelpMask.bmp.hbitmap = hBmpHelpMask;
    CComPtr<IPictureDisp> pdHelpMask ;
    OleCreatePictureIndirect(&pdescHelpMask, IID_IPictureDisp, TRUE,(void**)&pdHelpMask);

    //spCmdButtonHelp->PutMask(pdHelpMask);
    spCmdButtonHelp->PutVisible(VARIANT_TRUE); 
    spCmdButtonHelp->PutCaption(OLESTR("Help")); 
    spCmdButtonHelp->PutEnabled(VARIANT_TRUE);
    spCmdButtonHelp->PutTooltipText(OLESTR("Help")); 
    spCmdButtonHelp->PutTag(OLESTR("Tag for Help")); 



	m_spToolbar->PutVisible(VARIANT_TRUE); //show the toolbar
	
	m_spButtonMakeColl = spCmdButtonMakeColl;
	m_spButtonGetUpd = spCmdButtonGetUpd;
    m_spButtonOptions = spCmdButtonOptions;
    m_spButtonHelp = spCmdButtonHelp;

	if(mspVersion!=MSP_14)
	{
		//////////////////////////////////////////////////////////////////////////
		//ADDING A MENU ITEM
		CComPtr < Office::CommandBarControls>	spCmdCtrls;
		CComPtr < Office::CommandBarControl>	spCmdCtrl;
		hr = m_spToolBarCtrl->get_ActiveMenuBar(&spCmdBar); // get CommandBar that is Project's main menu
		if (FAILED(hr))
			return hr;
		spCmdCtrls = spCmdBar->GetControls(); // get menu as CommandBarControls 
		CComVariant vItem(6); // we want to add a menu entry to Project's 6th(Tools) menu item
		spCmdCtrl= spCmdCtrls->GetItem(vItem);

		IDispatchPtr spDisp;
		spDisp = spCmdCtrl->GetControl(); 
		CComQIPtr < Office::CommandBarPopup> spCmdPopup(spDisp);

		spCmdCtrls = spCmdPopup->GetControls();

		CComVariant vMenuPos(6);  
		CComVariant vMenuEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);
		CComVariant vMenuShow(VARIANT_TRUE);
		CComVariant vMenuTemp(VARIANT_TRUE);
		CComVariant vMenuType(1/*Office::MsoControlType::msoControlButton*/);

		CComPtr < Office::CommandBarControl> spNewMenuSTU = spCmdCtrls->Add(_variant_t(10), vMenuEmpty, vMenuEmpty, vMenuEmpty, vMenuTemp);

		IDispatchPtr spDispSTU;
		spDispSTU = spNewMenuSTU->GetControl(); 
		CComQIPtr < Office::CommandBarPopup> spNewPopupSTU(spDispSTU);
		spCmdCtrls = spNewPopupSTU->GetControls();

		CComPtr < Office::CommandBarControl> spNewMenuMakeColl = spCmdCtrls->Add(vMenuType, vMenuEmpty, vMenuEmpty, vMenuEmpty, vMenuTemp);
		CComPtr < Office::CommandBarControl> spNewMenuGetUpd = spCmdCtrls->Add(vMenuType, vMenuEmpty, vMenuEmpty, vMenuEmpty, vMenuTemp);
        CComPtr < Office::CommandBarControl> spNewMenuOptions = spCmdCtrls->Add(vMenuType, vMenuEmpty, vMenuEmpty, vMenuEmpty, vMenuTemp);
        CComPtr < Office::CommandBarControl> spNewMenuHelp = spCmdCtrls->Add(vMenuType, vMenuEmpty, vMenuEmpty, vMenuEmpty, vMenuTemp);

		spNewMenuSTU->PutBeginGroup(VARIANT_TRUE);
		spNewMenuSTU->PutCaption(OLESTR("Seavus Task Update"));
		spNewMenuSTU->PutEnabled(VARIANT_TRUE);
		spNewMenuSTU->PutVisible(VARIANT_TRUE); 

		spNewMenuMakeColl->PutCaption(OLESTR("Make Collaborative"));
		spNewMenuMakeColl->PutEnabled(VARIANT_TRUE);
		spNewMenuMakeColl->PutVisible(VARIANT_TRUE); 

		spNewMenuGetUpd->PutCaption(OLESTR("Get Updates"));
		spNewMenuGetUpd->PutEnabled(VARIANT_TRUE);
		spNewMenuGetUpd->PutVisible(VARIANT_TRUE); 

        spNewMenuOptions->PutCaption(OLESTR("Options"));
        spNewMenuOptions->PutEnabled(VARIANT_TRUE);
        spNewMenuOptions->PutVisible(VARIANT_TRUE); 

        spNewMenuHelp->PutCaption(OLESTR("Help"));
        spNewMenuHelp->PutEnabled(VARIANT_TRUE);
        spNewMenuHelp->PutVisible(VARIANT_TRUE); 

		CComQIPtr < Office::_CommandBarButton> spCmdMenuMakeColl(spNewMenuMakeColl);
		spCmdMenuMakeColl->PutStyle(Office::msoButtonCaption);     
		spNewMenuMakeColl->PutVisible(VARIANT_TRUE); 

		CComQIPtr < Office::_CommandBarButton> spCmdMenuGetUpd(spNewMenuGetUpd);
		spCmdMenuGetUpd->PutStyle(Office::msoButtonCaption);     
		spNewMenuGetUpd->PutVisible(VARIANT_TRUE); 

        CComQIPtr < Office::_CommandBarButton> spCmdMenuOptions(spNewMenuOptions);
        spCmdMenuOptions->PutStyle(Office::msoButtonCaption);     
        spNewMenuOptions->PutVisible(VARIANT_TRUE); 

        CComQIPtr < Office::_CommandBarButton> spCmdMenuHelp(spNewMenuHelp);
        spCmdMenuHelp->PutStyle(Office::msoButtonCaption);     
        spNewMenuHelp->PutVisible(VARIANT_TRUE); 

		CComQIPtr < Office::_CommandBarButton> spCmdMenuButton3(spNewPopupSTU);   
		spNewPopupSTU->PutVisible(VARIANT_TRUE); 

		//set member variables
		m_spMenuMakeColl = spCmdMenuMakeColl;
		m_spMenuGetUpd = spCmdMenuGetUpd;
        m_spMenuOptions = spCmdMenuOptions;
        m_spMenuHelp = spCmdMenuHelp;
		m_spMenuCtrl = spNewPopupSTU;
		//END OF ADDING MENUITEM
		//////////////////////////////////////////////////////////////////////////
	}

	m_bConnected = true;
	
	Utils::initCustomFieldNames();
	
	//Events
	EventsAdvise();
	//////////////////////////////////////////////////////////////////////////
	return S_OK;
}
STDMETHODIMP CConnect::OnDisconnection(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY ** custom )
{
	if(m_bConnected)
	{
		EventsUnadvise();
		DeleteControls();

		m_bConnected = false;
	}

	return S_OK;
}
STDMETHODIMP CConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}
STDMETHODIMP CConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}
STDMETHODIMP CConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}



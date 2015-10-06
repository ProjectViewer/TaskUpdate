// GetTimeUpdatesDlg.cpp : implementation file
//

#include "stdafx.h"
#include "GetTaskUpdatesDlg.h"
#include "afxdialogex.h"



// CGetTaskUpdatesDlg dialog

IMPLEMENT_DYNAMIC(CGetTaskUpdatesDlg, CDialogEx)

CGetTaskUpdatesDlg::CGetTaskUpdatesDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CGetTaskUpdatesDlg::IDD, pParent),
    m_pUpdatesManager(new CUpdatesManager())
{
	bAscending = true;
}

CGetTaskUpdatesDlg::CGetTaskUpdatesDlg(CUpdatesManager* updatesManager /*= 0*/, CWnd* pParent /*=NULL*/)
	: CDialogEx(CGetTaskUpdatesDlg::IDD, pParent),
	m_pUpdatesManager(updatesManager != 0 ? updatesManager : (new CUpdatesManager()))
{
	bAscending = true;
}

CGetTaskUpdatesDlg::~CGetTaskUpdatesDlg()
{	
    delete m_pUpdatesManager;
}

void CGetTaskUpdatesDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_LIST1, m_listCtrl);
    DDX_Control(pDX, IDC_BUTTON1, m_AcceptButton);
    DDX_Control(pDX, IDC_BUTTON2, m_RejectButton);
    DDX_Control(pDX, IDC_BUTTON3, m_ResetButton);
    DDX_Control(pDX, IDOK, m_buttonOK);
    DDX_Control(pDX, IDCANCEL, m_buttonCancel);
    DDX_Control(pDX, IDC_COMBO1, m_PrjNameCtrl);
    DDX_Control(pDX, IDC_COMBO2, m_TaskNameCtrl);
    DDX_Control(pDX, IDC_COMBO4, m_StatusCtrl);
    DDX_Control(pDX, IDC_WHITE_BKG, m_whiteBkg);
    DDX_Control(pDX, IDC_WHITE_BKG2, m_whiteBkg2);
    DDX_Control(pDX, IDC_WHITE_BKG3, m_whiteBkg3);
    DDX_Control(pDX, IDC_WHITE_BKG4, m_whiteBkg4);
    DDX_Control(pDX, IDC_FRAME, m_frame);
    DDX_Control(pDX, IDC_COMMENTGROUP, m_GrBoxComment);
    DDX_Control(pDX, IDC_EDIT1, m_editTMComment);
    DDX_Control(pDX, IDC_EDIT2, m_editPMComment);

    DDX_Control(pDX, IDC_STATIC_PROJ_FILTER, m_staticProjectFilter);
    DDX_Control(pDX, IDC_STATIC_TASK_FILTER, m_staticTaskFilter);
    DDX_Control(pDX, IDC_STATIC_STATUS_FILTER, m_staticStatusFilter);

    DDX_Control(pDX, IDC_STATIC_TM, m_staticTMComment);
    DDX_Control(pDX, IDC_STATIC_PM, m_staticPMComment);
}


BEGIN_MESSAGE_MAP(CGetTaskUpdatesDlg, CDialogEx)
    ON_BN_CLICKED(IDOK, &CGetTaskUpdatesDlg::OnBnClickedOk)
    ON_BN_CLICKED(IDC_BUTTON1, &CGetTaskUpdatesDlg::OnBnClickedButtonAccept)
    ON_BN_CLICKED(IDC_BUTTON2, &CGetTaskUpdatesDlg::OnBnClickedButtonReject)
    ON_BN_CLICKED(IDC_BUTTON3, &CGetTaskUpdatesDlg::OnBnClickedButtonReset)
    ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST1, &CGetTaskUpdatesDlg::OnLvnItemchangedList1)
    ON_NOTIFY(HDN_ITEMCLICK, 0, &CGetTaskUpdatesDlg::OnHdnItemclickList1)
    ON_NOTIFY(NM_CLICK, IDC_LIST1, &CGetTaskUpdatesDlg::OnNMClickList1)
	ON_BN_CLICKED(IDCANCEL, &CGetTaskUpdatesDlg::OnBnClickedCancel)
	ON_CBN_CLOSEUP(IDC_COMBO1, &CGetTaskUpdatesDlg::OnCbnCloseupComboProjectName)
	ON_CBN_CLOSEUP(IDC_COMBO2, &CGetTaskUpdatesDlg::OnCbnCloseupComboTaskName)
    ON_CBN_CLOSEUP(IDC_COMBO4, &CGetTaskUpdatesDlg::OnCbnCloseupComboStatus)
	ON_WM_SIZE()
	ON_WM_GETMINMAXINFO()
/*	ON_WM_CTLCOLOR()*/
	ON_EN_CHANGE(IDC_EDIT2, &CGetTaskUpdatesDlg::OnEnChangeEditPMComment)
END_MESSAGE_MAP()


// CGetTaskUpdatesDlg message handlers


BOOL CGetTaskUpdatesDlg::OnInitDialog()
{
 	CDialogEx::OnInitDialog();
                                
	//fill Filters
	UpdateProjectFilter(m_pUpdatesManager->getProjReport());
	UpdateTaskFilter(m_pUpdatesManager->getProjReport());
    UpdateStatusFilter(m_pUpdatesManager->getProjReport());

	m_listCtrl.SetExtendedStyle(m_listCtrl.GetExtendedStyle()|LVS_EX_GRIDLINES|LVS_EX_FULLROWSELECT|LVS_EX_CHECKBOXES|LVS_EX_SINGLEROW);
	UpdateListCtrl();
	m_listCtrl.CreateHeaderCheckBox();
    m_listCtrl.SetFocus();
    m_listCtrl.SetItemState(0, LVIS_SELECTED, LVIS_SELECTED);

    m_AcceptButton.EnableWindow(false);
    m_RejectButton.EnableWindow(false);
    m_buttonOK.EnableWindow(false);       
    m_pUpdatesManager->lockSPVFiles(m_pUpdatesManager->getProjInfoList());

	return FALSE;  // return TRUE  unless you set the focus to a control
}
void CGetTaskUpdatesDlg::OnSize(UINT nType, int cx, int cy)
{
    CDialogEx::OnSize(nType, cx, cy);

    CRect rcDlg;
    GetWindowRect(&rcDlg);
    ScreenToClient(&rcDlg);

    const int commGrBottom = 95;
    const int commGrRight = 30;

    const int macGrBottom = 199;
    const int macGrRight = 30;

    CRect rc;
    if(m_whiteBkg)
    {
        m_whiteBkg.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.right = rcDlg.right - 18;
        rc.bottom = rcDlg.bottom - 43;
        rc.top = rcDlg.bottom - 303;
        m_whiteBkg.MoveWindow(rc);
    }
    if(m_whiteBkg2)
    {
        m_whiteBkg2.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.right = rcDlg.right - 18;
        m_whiteBkg2.MoveWindow(rc);
    }
    if(m_whiteBkg3)
    {
        m_whiteBkg3.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.bottom = rcDlg.bottom - 57;
        m_whiteBkg3.MoveWindow(rc);
    }
    if(m_whiteBkg4)
    {
        m_whiteBkg4.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.bottom = rcDlg.bottom - 57;
        rc.right = rcDlg.right - 18;
        rc.left = rcDlg.right - 30;
        m_whiteBkg4.MoveWindow(rc);
    }
    if(m_frame)
    {
        m_frame.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.right = rcDlg.right - 18;
        rc.bottom = rcDlg.bottom - 42;
        m_frame.MoveWindow(rc);
    }
    if(m_listCtrl)
    {
        m_listCtrl.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.right = rcDlg.right - 30;
        rc.bottom = rcDlg.bottom - 199;
        m_listCtrl.MoveWindow(rc);
    }
    if(m_AcceptButton)
    {
        m_AcceptButton.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.bottom = rcDlg.bottom - 59;
        rc.top = rc.bottom	- 23;
        m_AcceptButton.MoveWindow(rc);
    }
    if(m_RejectButton)
    {
        m_RejectButton.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.bottom = rcDlg.bottom - 59;
        rc.top = rc.bottom	- 23;
        m_RejectButton.MoveWindow(rc);
    }
    if(m_ResetButton)
    {
        m_ResetButton.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.bottom = rcDlg.bottom - 59;
        rc.top = rc.bottom	- 23;
        m_ResetButton.MoveWindow(rc);
    }
    if(m_buttonOK)
    {
        m_buttonOK.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.bottom = rcDlg.bottom - 12;
        rc.top = rc.bottom	- 23;
        rc.right = rcDlg.right - 100;
        rc.left = rc.right - 75;
        m_buttonOK.MoveWindow(rc);
    }
    if(m_buttonCancel)
    {
        m_buttonCancel.GetWindowRect(&rc);
        ScreenToClient(&rc);
        rc.bottom = rcDlg.bottom - 12;
        rc.top = rc.bottom	- 23;
        rc.right = rcDlg.right - 17;
        rc.left = rc.right - 75;
        m_buttonCancel.MoveWindow(rc);
    }

    if(m_GrBoxComment)//comment group
    {
        CRect rcGrBoxComment;
        m_GrBoxComment.GetWindowRect(&rcGrBoxComment);
        ScreenToClient(&rcGrBoxComment);
        rcGrBoxComment.bottom = rcDlg.bottom - commGrBottom;
        rcGrBoxComment.top = rcGrBoxComment.bottom	- 93;
        rcGrBoxComment.right = rcDlg.right - commGrRight;
        m_GrBoxComment.MoveWindow(rcGrBoxComment);

        if(m_editTMComment)
        {
            CRect rcEditTMComment(
                rcGrBoxComment.left + 9,
                rcGrBoxComment.top + 36,
                (rcGrBoxComment.left+rcGrBoxComment.right)/2-14,
                rcGrBoxComment.bottom - 8);
            m_editTMComment.MoveWindow(rcEditTMComment);
        }
        if(m_editPMComment)
        {
            CRect rcEditPMComment(
                (rcGrBoxComment.left+rcGrBoxComment.right)/2+14,
                rcGrBoxComment.top + 36,
                rcGrBoxComment.right-9,
                rcGrBoxComment.bottom - 8);
            m_editPMComment.MoveWindow(rcEditPMComment);
        }
        CStatic* pTMCommStatic = (CStatic*)GetDlgItem(IDC_STATIC_TM);
        if(pTMCommStatic)
        {
            CRect rcTMCommStatic(
                rcGrBoxComment.left+9,
                rcGrBoxComment.top+20,
                (rcGrBoxComment.left+rcGrBoxComment.right)/2-14,
                rcGrBoxComment.top+33);
            pTMCommStatic->MoveWindow(rcTMCommStatic);
        }
        CStatic* pPMCommStatic = (CStatic*)GetDlgItem(IDC_STATIC_PM);
        if(pPMCommStatic)
        {
            CRect rcPMCommStatic(
                (rcGrBoxComment.left+rcGrBoxComment.right)/2+14,
                rcGrBoxComment.top+20,
                rcGrBoxComment.right-9,
                rcGrBoxComment.top+33);
            pPMCommStatic->MoveWindow(rcPMCommStatic);
        }
    }

    Invalidate();
}
void CGetTaskUpdatesDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
    lpMMI->ptMinTrackSize.x = 796;
    lpMMI->ptMinTrackSize.y = 400;

    CDialogEx::OnGetMinMaxInfo(lpMMI);
}
HBRUSH CGetTaskUpdatesDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
    HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);
    return hbr;
}
void CGetTaskUpdatesDlg::UpdateProjectFilter(CTUProjectReport* pProjectReport)
{
	m_PrjNameCtrl.InsertString(0, _T("All"));
	m_PrjNameCtrl.SetItemGUID(0, _T(""));
	m_PrjNameCtrl.SetItemDescription(0, _T("All"));

	int iItem=1;

    TUTaskReportsList::iterator itTaskReports;
    for(itTaskReports = pProjectReport->GetTaskReportsList()->begin();
        itTaskReports != pProjectReport->GetTaskReportsList()->end();
        ++itTaskReports)
    {
        if(TUCMB_NOTVALID == m_PrjNameCtrl.GeItemWithGUID((*itTaskReports)->GetProjectGUID()))
        {
            m_PrjNameCtrl.InsertString(iItem, (*itTaskReports)->GetProjectName());
            m_PrjNameCtrl.SetItemGUID(iItem, (*itTaskReports)->GetProjectGUID());
            m_PrjNameCtrl.SetItemDescription(iItem, (*itTaskReports)->GetProjectName());	
            iItem++;
        }
    }

	m_PrjNameCtrl.SelectItemWithGUID(TUCMB_ALL_PROJECTS);
}
void CGetTaskUpdatesDlg::UpdateTaskFilter(CTUProjectReport* pProjectReport)
{
    m_TaskNameCtrl.InsertString(0, _T("All"));
	m_TaskNameCtrl.SetItemKey(0, TUCMB_ALL_TASKS);
	m_TaskNameCtrl.SetItemDescription(0, _T("All"));	

	int iItem=1;

    TUTaskReportsList::iterator itTaskReports;
    for(itTaskReports = pProjectReport->GetTaskReportsList()->begin();
        itTaskReports != pProjectReport->GetTaskReportsList()->end();
        ++itTaskReports)
    {
        if(TUCMB_NOTVALID == m_TaskNameCtrl.GeItemWithKey((int)(*itTaskReports)->GetTaskUID()))
        {
            m_TaskNameCtrl.InsertString(iItem, (*itTaskReports)->GetTaskName());
            m_TaskNameCtrl.SetItemKey(iItem, (int)(*itTaskReports)->GetTaskUID());
            m_TaskNameCtrl.SetItemDescription(iItem, (*itTaskReports)->GetTaskName());	
            iItem++;
        }
    }

	m_TaskNameCtrl.SelectItemWithKey(TUCMB_ALL_TASKS);
}
void CGetTaskUpdatesDlg::UpdateStatusFilter(CTUProjectReport* pProjectReport)
{
    m_StatusCtrl.InsertString(0, _T("All"));
    m_StatusCtrl.SetItemKey(0, TUCMB_ALL_STATUSES);
    m_StatusCtrl.SetItemDescription(0, _T("All"));	

    m_StatusCtrl.InsertString(1, _T("Pending"));
    m_StatusCtrl.SetItemKey(1, TUTaskOpen);
    m_StatusCtrl.SetItemDescription(1, _T("Pending"));	

    m_StatusCtrl.InsertString(2, _T("To Be Accepted"));
    m_StatusCtrl.SetItemKey(2, TUTaskToBeAccepted);
    m_StatusCtrl.SetItemDescription(2, _T("To Be Accepted"));	

    m_StatusCtrl.InsertString(3, _T("To Be Rejected"));
    m_StatusCtrl.SetItemKey(3, TUTaskToBeRejected);
    m_StatusCtrl.SetItemDescription(3, _T("To Be Rejected"));	

    m_StatusCtrl.InsertString(4, _T("Accepted"));
    m_StatusCtrl.SetItemKey(4, TUTaskAccepted);
    m_StatusCtrl.SetItemDescription(4, _T("Accepted"));	

    m_StatusCtrl.InsertString(5, _T("Rejected"));
    m_StatusCtrl.SetItemKey(5, TUTaskRejected);
    m_StatusCtrl.SetItemDescription(5, _T("Rejected"));	

    m_StatusCtrl.SelectItemWithKey(TUTaskOpen);
}
void CGetTaskUpdatesDlg::UpdateListCtrl()
{
	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	
    CString GUID = m_listCtrl.GetSelectedItemGUID();

    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    MSProject::_IProjectDocPtr project = app->GetActiveProject();  
    MSProject::PjUnit tempDefaultWorkUnits=project->GetDefaultWorkUnits();
    enumTUUnit tempWorkUnit;
    switch(tempDefaultWorkUnits)
    {
        case MSProject::pjMinute    : tempWorkUnit = TUUnitMin;     break;   
        case MSProject::pjHour      : tempWorkUnit = TUUnitHour;    break;
        case MSProject::pjDay       : tempWorkUnit = TUUnitDay;     break;
        case MSProject::pjWeek      : tempWorkUnit = TUUnitWeek;    break;  
        case MSProject::pjMonthUnit : tempWorkUnit = TUUnitMonth;   break; 

        default:
            tempWorkUnit = TUUnitNone;
    }
    double tempGetHoursPerDay=project->GetHoursPerDay();
    double tempGetHoursPerWeek=project->GetHoursPerWeek();
    double tempGetDaysPerMonth=project->GetDaysPerMonth();

	CString prjFilter = m_PrjNameCtrl.GetSelectedItemsGUID();
	int taskFilter = m_TaskNameCtrl.GetSelectedItemsKey();
    int statusFilter = m_StatusCtrl.GetSelectedItemsKey();
   
	int columns = COLUMN_CHECKBOX | COLUMN_NAME;

	if (pProjReport->HasReportWithWorkInsertMode(TUTaskPercentCompleteReport))
	{
		columns = columns | COLUMN_CURR_PERC_COMPL | COLUMN_REP_PERC_COMPL;
	}
	if (pProjReport->HasReportWithWorkInsertMode(TUTaskPercentWorkCompleteReport))
	{
		columns = columns | COLUMN_CURR_PERC_WORK_COMPL | COLUMN_REP_PERC_WORK_COMPL;
	}
	if (pProjReport->HasReportWithWorkInsertMode(TUTaskActualWorkReport))
	{
		columns = columns | COLUMN_CURR_ACT_WORK | COLUMN_REP_ACT_WORK;
	}
	if (pProjReport->HasReportWithWorkInsertMode(TUTaskActualStartReport))
	{
		columns = columns | COLUMN_CURR_ACT_START | COLUMN_REP_ACT_START;
	}
	if (pProjReport->HasReportWithWorkInsertMode(TUTaskActualFinishReport))
	{
		columns = columns | COLUMN_CURR_ACT_FINISH | COLUMN_REP_ACT_FINISH;
	}
    if (pProjReport->HasReportWithWorkInsertMode(TUTaskActualDurationReport))
    {
        columns = columns | COLUMN_CURR_ACT_DURATION | COLUMN_REP_ACT_DURATION;
    }

    m_listCtrl.FillListCtrl_Task(pProjReport,  tempWorkUnit, tempGetHoursPerDay, tempGetHoursPerWeek, (int)tempGetDaysPerMonth,
		columns, prjFilter, taskFilter, statusFilter);
    
    m_listCtrl.SelectItem(GUID);
}
void CGetTaskUpdatesDlg::OnBnClickedOk()
{
	if(IDOK == AfxMessageBox(_T("The selected updates will be saved in the project plan. Do you want to continue?"),MB_OKCANCEL))
	{
		//unlock SPV files before writing
		m_pUpdatesManager->unLockSPVFiles(m_pUpdatesManager->getProjInfoList()); 

		//write accepted updates to MPP files
        bool bCannotWriteManual = false;
		m_pUpdatesManager->processWritingToFiles(m_pUpdatesManager->getProjInfoList(), bCannotWriteManual, TUTaskReportingMode); 

        m_pUpdatesManager->getProjReport()->ChangeAllReportsStatus(TUTaskToBeAccepted, TUTaskAccepted);
        m_pUpdatesManager->getProjReport()->ChangeAllReportsStatus(TUTaskToBeRejected, TUTaskRejected);
        m_pUpdatesManager->getProjReport()->UpdateCurrentValues();//this is to change current values for accepted reports 

        //save file
        MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
        if  (0 != app->FileSave())
        {
            //if manual cannot be updated (MSP 2010 without Service Pack)
            if (bCannotWriteManual)
            {
                AfxMessageBox(_T("To update Manually Scheduled Tasks, please install Service Pack 1 for Microsoft Project 2010"));
            }
        }

		//write reports to all SPV files
		bool bSuccess = m_pUpdatesManager->saveSPVFiles(m_pUpdatesManager->getProjInfoList()); 

		//activate main project(if there are sub projects, activate master project)
		m_pUpdatesManager->reopenProject();
		CDialogEx::OnCancel();
	}
}
void CGetTaskUpdatesDlg::OnBnClickedButtonAccept() 
{
	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
		return;

	int itemCount = m_listCtrl.GetItemCount();
	for(int iItem =0; iItem<itemCount; ++iItem)
	{
		if(m_listCtrl.GetCheck(iItem))
		{
			CString strGuid = m_listCtrl.GetItemGUID(iItem);

            CTUTaskReport* pTaskreport = pProjReport->GetTaskReport(strGuid);
            if (pTaskreport)
            {
                pTaskreport->SetStatus(TUTaskToBeAccepted);
            }
		}
	}

	UpdateListCtrl();
	m_buttonOK.EnableWindow(true);     
}
void CGetTaskUpdatesDlg::OnBnClickedButtonReject()  
{
	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
		return;

	int itemCount = m_listCtrl.GetItemCount();
	for(int iItem =0; iItem<itemCount; ++iItem)
	{
		if(m_listCtrl.GetCheck(iItem))
		{
			CString strGuid = m_listCtrl.GetItemGUID(iItem);

            CTUTaskReport* pTaskreport = pProjReport->GetTaskReport(strGuid);
            if (pTaskreport)
            {
                pTaskreport->SetStatus(TUTaskToBeRejected);
            }
		}
	}

	UpdateListCtrl();
	m_buttonOK.EnableWindow(true); 
}
void CGetTaskUpdatesDlg::OnBnClickedButtonReset()
{
    m_listCtrl.UncheckAllItems();

	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(pProjReport)
	{
        pProjReport->ChangeAllReportsStatus(TUTaskToBeAccepted, TUTaskOpen);
        pProjReport->ChangeAllReportsStatus(TUTaskToBeRejected, TUTaskOpen);
	}

	UpdateListCtrl();
	m_buttonOK.EnableWindow(false); 
}
void CGetTaskUpdatesDlg::OnLvnItemchangedList1(NMHDR *pNMHDR, LRESULT *pResult)
{
    LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
   	m_listCtrl.Check_Uncheck_HeaderCheckBox();

    bool bHasCheckItem = m_listCtrl.HasCheckedItem();

    m_AcceptButton.EnableWindow(bHasCheckItem);
    m_RejectButton.EnableWindow(bHasCheckItem);

    m_editTMComment.SetWindowText(_T(""));
    m_editPMComment.SetWindowText(_T(""));

    CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
    if(pProjReport)
    {
        CString strGUID = m_listCtrl.GetSelectedItemGUID();
        if(!strGUID.IsEmpty())
        {
            CTUTaskReport* pTaskReport = pProjReport->GetTaskReport(strGUID);
            if(pTaskReport)
            {
                m_editTMComment.SetWindowText(pTaskReport->GetTMNote());
                m_editPMComment.SetWindowText(pTaskReport->GetPMNote());
            }
        }
    }
    UpdateData(FALSE);

    *pResult = 0;
}
void CGetTaskUpdatesDlg::OnHdnItemclickList1(NMHDR *pNMHDR, LRESULT *pResult)
{
	NMLISTVIEW *pLV = (NMLISTVIEW *) pNMHDR;

	int colType;
	if(pLV->iItem == 2)
		colType=SORT_DAT;//Report Date column
	else
		colType=SORT_STR;//All other columns

	if(pLV->iItem>0)
	{
		m_listCtrl.SortColumn(pLV->iItem,!bAscending,colType);
		bAscending=!bAscending;
	}
 
	if(pLV->iItem==0)
		m_listCtrl.Check_Uncheck_AllItems();

    *pResult = 0;
}
void CGetTaskUpdatesDlg::OnNMClickList1(NMHDR *pNMHDR, LRESULT *pResult)
{
	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
		return;

    bool bHasCheckedItem = m_listCtrl.HasCheckedItem();
	bool bHasSaveableReport = pProjReport->HasAcceptedReport() || pProjReport->HasRejectedReport();

    m_AcceptButton.EnableWindow(bHasCheckedItem);
    m_RejectButton.EnableWindow(bHasCheckedItem);
    m_buttonOK.EnableWindow(bHasSaveableReport);       

	m_editTMComment.SetWindowText(_T(""));
	m_editPMComment.SetWindowText(_T(""));


    CString strGUID = m_listCtrl.GetSelectedItemGUID();
    if(!strGUID.IsEmpty())
    {
        CTUTaskReport* pTaskReport = pProjReport->GetTaskReport(strGUID);
        if(pTaskReport)
        {
            m_editTMComment.SetWindowText(pTaskReport->GetTMNote());
            m_editPMComment.SetWindowText(pTaskReport->GetPMNote());
        }
    }
    UpdateData(FALSE);
    *pResult = 0;
}
void CGetTaskUpdatesDlg::OnBnClickedCancel()
{
	//unlock SPV files before writing
	m_pUpdatesManager->unLockSPVFiles(m_pUpdatesManager->getProjInfoList()); 

    CDialogEx::OnCancel();
}
void CGetTaskUpdatesDlg::OnCbnCloseupComboProjectName()
{
	UpdateListCtrl();

	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
		return;

	bool bHasSaveableReport = pProjReport->HasAcceptedReport() || pProjReport->HasRejectedReport();
	m_buttonOK.EnableWindow(bHasSaveableReport);   
}
void CGetTaskUpdatesDlg::OnCbnCloseupComboTaskName()
{
	UpdateListCtrl();

	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
		return;

	bool bHasSaveableReport = pProjReport->HasAcceptedReport() || pProjReport->HasRejectedReport();
	m_buttonOK.EnableWindow(bHasSaveableReport);   
}
void CGetTaskUpdatesDlg::OnCbnCloseupComboStatus()
{
    UpdateListCtrl();

    CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
    if(!pProjReport)
        return;

    bool bHasSaveableReport = pProjReport->HasAcceptedReport() || pProjReport->HasRejectedReport();
    m_buttonOK.EnableWindow(bHasSaveableReport);   
}
void CGetTaskUpdatesDlg::OnEnChangeEditPMComment()
{
	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
	{		
		return;
	}

	CString strGUID = m_listCtrl.GetSelectedItemGUID();
	if(!strGUID.IsEmpty())
	{
        CTUTaskReport* pTaskReport = pProjReport->GetTaskReport(strGUID);
        if (pTaskReport)
        {
            CString str;
            m_editPMComment.GetWindowText(str);

            pTaskReport->SetPMNote(str);
        }
	}

    m_buttonOK.EnableWindow(true);  
}

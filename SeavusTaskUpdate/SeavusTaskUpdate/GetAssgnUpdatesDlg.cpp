// GetTimeUpdatesDlg.cpp : implementation file
//

#include "stdafx.h"
#include "GetAssgnUpdatesDlg.h"
#include "afxdialogex.h"



// CGetAssgnUpdatesDlg dialog

IMPLEMENT_DYNAMIC(CGetAssgnUpdatesDlg, CDialogEx)

CGetAssgnUpdatesDlg::CGetAssgnUpdatesDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CGetAssgnUpdatesDlg::IDD, pParent),
    m_pUpdatesManager(new CUpdatesManager())
    , m_IsAssigmMarkedCompleted(FALSE)
    , m_iMarkedAsComplete(0)
{
	bAscending = true;
}
CGetAssgnUpdatesDlg::CGetAssgnUpdatesDlg(CUpdatesManager* updatesManager /*= 0*/, CWnd* pParent /*=NULL*/)
	: CDialogEx(CGetAssgnUpdatesDlg::IDD, pParent),
	m_pUpdatesManager(updatesManager != 0 ? updatesManager : (new CUpdatesManager()))
	, m_IsAssigmMarkedCompleted(FALSE)
	, m_iMarkedAsComplete(0)
{
	bAscending = true;
}

CGetAssgnUpdatesDlg::~CGetAssgnUpdatesDlg()
{	
    delete m_pUpdatesManager;
}

void CGetAssgnUpdatesDlg::DoDataExchange(CDataExchange* pDX)
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
    DDX_Control(pDX, IDC_COMBO3, m_ResNameCtrl);
    DDX_Control(pDX, IDC_COMBO5, m_StatusCtrl);
    DDX_Control(pDX, IDC_WHITE_BKG, m_whiteBkg);
    DDX_Control(pDX, IDC_WHITE_BKG2, m_whiteBkg2);
    DDX_Control(pDX, IDC_WHITE_BKG3, m_whiteBkg3);
    DDX_Control(pDX, IDC_WHITE_BKG4, m_whiteBkg4);
    DDX_Control(pDX, IDC_FRAME, m_frame);
    DDX_Control(pDX, IDC_COMMENTGROUP, m_GrBoxComment);
    DDX_Control(pDX, IDC_EDIT1, m_editTMComment);
    DDX_Control(pDX, IDC_EDIT2, m_editPMComment);
    DDX_Radio(pDX, IDC_RADIO_NONE, m_iMarkedAsComplete);
    DDX_Control(pDX, IDC_MACGROUP, m_GrBoxMarkAsComplete);
    DDX_Control(pDX, IDC_RADIO_NONE, m_radioNone);
    DDX_Control(pDX, IDC_RADIO_FINISH, m_radioFinish);
    DDX_Control(pDX, IDC_RADIO_PERCENT_WORK_COMPLETED, m_radioPercent);

    DDX_Control(pDX, IDC_STATIC_PROJ_FILTER, m_staticProjectFilter);
    DDX_Control(pDX, IDC_STATIC_TASK_FILTER, m_staticTaskFilter);
    DDX_Control(pDX, IDC_STATIC_RES_FILTER, m_staticResourceFilter);
    DDX_Control(pDX, IDC_STATIC_ASSGN_STATUS_FILTER, m_staticStatusFilter);

    DDX_Control(pDX, IDC_STATIC_TM, m_staticTMComment);
    DDX_Control(pDX, IDC_STATIC_PM, m_staticPMComment);
    
    DDX_Control(pDX, IDC_STATIC_ASSIG_MARKED_AS_COMPLETED, m_staticMACNotificication);
    DDX_Control(pDX, IDC_STATIC_SELECT_MOST_APPROPRIATE, m_staticMACSelectMethod);
}


BEGIN_MESSAGE_MAP(CGetAssgnUpdatesDlg, CDialogEx)
    ON_BN_CLICKED(IDOK, &CGetAssgnUpdatesDlg::OnBnClickedOk)
    ON_BN_CLICKED(IDC_BUTTON1, &CGetAssgnUpdatesDlg::OnBnClickedButtonAccept)
    ON_BN_CLICKED(IDC_BUTTON2, &CGetAssgnUpdatesDlg::OnBnClickedButtonReject)
    ON_BN_CLICKED(IDC_BUTTON3, &CGetAssgnUpdatesDlg::OnBnClickedButtonReset)
    ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST1, &CGetAssgnUpdatesDlg::OnLvnItemchangedList1)
    ON_NOTIFY(HDN_ITEMCLICK, 0, &CGetAssgnUpdatesDlg::OnHdnItemclickList1)
    ON_NOTIFY(NM_CLICK, IDC_LIST1, &CGetAssgnUpdatesDlg::OnNMClickList1)
	ON_BN_CLICKED(IDCANCEL, &CGetAssgnUpdatesDlg::OnBnClickedCancel)
	ON_CBN_CLOSEUP(IDC_COMBO1, &CGetAssgnUpdatesDlg::OnCbnCloseupComboProjectName)
	ON_CBN_CLOSEUP(IDC_COMBO2, &CGetAssgnUpdatesDlg::OnCbnCloseupComboTaskName)
	ON_CBN_CLOSEUP(IDC_COMBO3, &CGetAssgnUpdatesDlg::OnCbnCloseupComboResourceName)
    ON_CBN_CLOSEUP(IDC_COMBO5, &CGetAssgnUpdatesDlg::OnCbnCloseupComboStatus)
	ON_WM_SIZE()
	ON_WM_GETMINMAXINFO()
/*	ON_WM_CTLCOLOR()*/
	ON_EN_CHANGE(IDC_EDIT2, &CGetAssgnUpdatesDlg::OnEnChangeEditPMComment)
    ON_BN_CLICKED(IDC_RADIO_NONE, &CGetAssgnUpdatesDlg::UpdateReportsMarkAsCompleteType)
    ON_BN_CLICKED(IDC_RADIO_PERCENT_WORK_COMPLETED, &CGetAssgnUpdatesDlg::UpdateReportsMarkAsCompleteType)
    ON_BN_CLICKED(IDC_RADIO_FINISH, &CGetAssgnUpdatesDlg::UpdateReportsMarkAsCompleteType)
END_MESSAGE_MAP()


// CGetAssgnUpdatesDlg message handlers


BOOL CGetAssgnUpdatesDlg::OnInitDialog()
{
 	CDialogEx::OnInitDialog();
                                
	//fill Filters
	UpdateProjectFilter(m_pUpdatesManager->getProjReport());
	UpdateTaskFilter(m_pUpdatesManager->getProjReport());
	UpdateResourceFilter(m_pUpdatesManager->getProjReport());
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

    UpdateMarkAsCompleteRadioBtns();

	return FALSE;  // return TRUE  unless you set the focus to a control
}
void CGetAssgnUpdatesDlg::OnSize(UINT nType, int cx, int cy)
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
        rc.bottom = rcDlg.bottom - 303;
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

    if (m_GrBoxMarkAsComplete)
    {
        CRect rcGrBoxMarkAsComplete;
        m_GrBoxMarkAsComplete.GetWindowRect(&rcGrBoxMarkAsComplete);
        ScreenToClient(&rcGrBoxMarkAsComplete);
        rcGrBoxMarkAsComplete.bottom = rcDlg.bottom - macGrBottom;
        rcGrBoxMarkAsComplete.top = rcGrBoxMarkAsComplete.bottom - 93;
        rcGrBoxMarkAsComplete.right = rcDlg.right - macGrRight;
        m_GrBoxMarkAsComplete.MoveWindow(rcGrBoxMarkAsComplete);

        if(m_radioNone)
        {
            CRect rcRadioNone;
            m_radioNone.GetWindowRect(&rcRadioNone);
            int w = rcRadioNone.Width();

            rcRadioNone.SetRect(
                (rcGrBoxMarkAsComplete.left+rcGrBoxMarkAsComplete.right)/2+14,
                rcGrBoxMarkAsComplete.top + 34,
                (rcGrBoxMarkAsComplete.left+rcGrBoxMarkAsComplete.right)/2+14 + w,
                rcGrBoxMarkAsComplete.top + 50);
            m_radioNone.MoveWindow(rcRadioNone);
        }
        if(m_radioPercent)
        {
            CRect rcRadioPercent;
            m_radioPercent.GetWindowRect(&rcRadioPercent);
            int w = rcRadioPercent.Width();

            rcRadioPercent.SetRect(
                (rcGrBoxMarkAsComplete.left+rcGrBoxMarkAsComplete.right)/2+14,
                rcGrBoxMarkAsComplete.top + 52,
                (rcGrBoxMarkAsComplete.left+rcGrBoxMarkAsComplete.right)/2+14 + w,
                rcGrBoxMarkAsComplete.top + 68);
            m_radioPercent.MoveWindow(rcRadioPercent);
        }
        if(m_radioFinish)
        {
            CRect rcRadioFinish;
            m_radioFinish.GetWindowRect(&rcRadioFinish);
            int w = rcRadioFinish.Width();

            rcRadioFinish.SetRect(
                (rcGrBoxMarkAsComplete.left+rcGrBoxMarkAsComplete.right)/2+14,
                rcGrBoxMarkAsComplete.top + 70,
                (rcGrBoxMarkAsComplete.left+rcGrBoxMarkAsComplete.right)/2+14 + w,
                rcGrBoxMarkAsComplete.top + 86);
            m_radioFinish.MoveWindow(rcRadioFinish);
        }

        if(m_staticMACSelectMethod)
        {
            CRect rcGrMACSelect;
            m_staticMACSelectMethod.GetWindowRect(&rcGrMACSelect);
            int w = rcGrMACSelect.Width();
            rcGrMACSelect.SetRect(
                (rcGrBoxMarkAsComplete.left+rcGrBoxMarkAsComplete.right)/2+14,
                rcGrBoxMarkAsComplete.top + 15,
                (rcGrBoxMarkAsComplete.left+rcGrBoxMarkAsComplete.right)/2+14 + w,
                rcGrBoxMarkAsComplete.top + 28);
            m_staticMACSelectMethod.MoveWindow(rcGrMACSelect);
        }

        if(m_staticMACNotificication)
        {
            CRect rcGrMACNote;
            m_staticMACNotificication.GetWindowRect(&rcGrMACNote);
            int w = rcGrMACNote.Width();
            rcGrMACNote.SetRect(
                rcGrBoxMarkAsComplete.left + 9,
                rcGrBoxMarkAsComplete.top + 29,
                (rcGrBoxMarkAsComplete.left+rcGrBoxMarkAsComplete.right)/2-14,
                rcGrBoxMarkAsComplete.top + 73);
            m_staticMACNotificication.MoveWindow(rcGrMACNote);
        }

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
void CGetAssgnUpdatesDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
    lpMMI->ptMinTrackSize.x = 796;
    lpMMI->ptMinTrackSize.y = 500;

    CDialogEx::OnGetMinMaxInfo(lpMMI);
}
HBRUSH CGetAssgnUpdatesDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
    HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);
    return hbr;
}
void CGetAssgnUpdatesDlg::UpdateProjectFilter(CTUProjectReport* pProjectReport)
{
	TUAssignmentReportsList::iterator itAssgnReports;
	TUDateReportsList::iterator itDateReports;

	m_PrjNameCtrl.InsertString(0, _T("All"));
	m_PrjNameCtrl.SetItemGUID(0, _T(""));
	m_PrjNameCtrl.SetItemDescription(0, _T("All"));

	int iItem=1;
	for(itAssgnReports = pProjectReport->GetAssignmentReportsList()->begin();
		itAssgnReports != pProjectReport->GetAssignmentReportsList()->end();
		++itAssgnReports)
	{
		if(TUCMB_NOTVALID == m_PrjNameCtrl.GeItemWithGUID((*itAssgnReports)->GetProjectGUID()))
		{
			m_PrjNameCtrl.InsertString(iItem, (*itAssgnReports)->GetProjectName());
			m_PrjNameCtrl.SetItemGUID(iItem, (*itAssgnReports)->GetProjectGUID());
			m_PrjNameCtrl.SetItemDescription(iItem, (*itAssgnReports)->GetProjectName());	
			iItem++;
		}
	}

	m_PrjNameCtrl.SelectItemWithGUID(TUCMB_ALL_PROJECTS);
}
void CGetAssgnUpdatesDlg::UpdateTaskFilter(CTUProjectReport* pProjectReport)
{
	TUAssignmentReportsList::iterator itAssgnReports;
	TUDateReportsList::iterator itDateReports;

	m_TaskNameCtrl.InsertString(0, _T("All"));
	m_TaskNameCtrl.SetItemKey(0, TUCMB_ALL_TASKS);
	m_TaskNameCtrl.SetItemDescription(0, _T("All"));	

	int iItem=1;
	for(itAssgnReports = pProjectReport->GetAssignmentReportsList()->begin();
		itAssgnReports != pProjectReport->GetAssignmentReportsList()->end();
		++itAssgnReports)
	{
		if(TUCMB_NOTVALID == m_TaskNameCtrl.GeItemWithKey((int)(*itAssgnReports)->GetTaskUID()))
		{
			m_TaskNameCtrl.InsertString(iItem, (*itAssgnReports)->GetTaskName());
			m_TaskNameCtrl.SetItemKey(iItem, (int)(*itAssgnReports)->GetTaskUID());
			m_TaskNameCtrl.SetItemDescription(iItem, (*itAssgnReports)->GetTaskName());	
			iItem++;
		}
	}

	m_TaskNameCtrl.SelectItemWithKey(TUCMB_ALL_TASKS);
}
void CGetAssgnUpdatesDlg::UpdateResourceFilter(CTUProjectReport* pProjectReport)
{
	TUAssignmentReportsList::iterator itAssgnReports;
	TUDateReportsList::iterator itDateReports;

	m_ResNameCtrl.InsertString(0, _T("All"));
	m_ResNameCtrl.SetItemKey(0, TUCMB_ALL_RESOURCES);
	m_ResNameCtrl.SetItemDescription(0, _T("All"));	

	int iItem=1;
	for(itAssgnReports = pProjectReport->GetAssignmentReportsList()->begin();
		itAssgnReports != pProjectReport->GetAssignmentReportsList()->end();
		++itAssgnReports)
	{
		if(TUCMB_NOTVALID == m_ResNameCtrl.GeItemWithKey((int)(*itAssgnReports)->GetResourceUID()))
		{
			m_ResNameCtrl.InsertString(iItem, (*itAssgnReports)->GetResourceName());
			m_ResNameCtrl.SetItemKey(iItem, (int)(*itAssgnReports)->GetResourceUID());
			m_ResNameCtrl.SetItemDescription(iItem, (*itAssgnReports)->GetResourceName());	
			iItem++;
		}
	}

	m_ResNameCtrl.SelectItemWithKey(TUCMB_ALL_RESOURCES);
}
void CGetAssgnUpdatesDlg::UpdateStatusFilter(CTUProjectReport* pProjectReport)
{
    m_StatusCtrl.InsertString(0, _T("All"));
    m_StatusCtrl.SetItemKey(0, TUCMB_ALL_STATUSES);
    m_StatusCtrl.SetItemDescription(0, _T("All"));	

    m_StatusCtrl.InsertString(1, _T("Pending"));
    m_StatusCtrl.SetItemKey(1, TUOpen);
    m_StatusCtrl.SetItemDescription(1, _T("Pending"));	

    m_StatusCtrl.InsertString(2, _T("To Be Accepted"));
    m_StatusCtrl.SetItemKey(2, TUToBeAccepted);
    m_StatusCtrl.SetItemDescription(2, _T("To Be Accepted"));	

    m_StatusCtrl.InsertString(3, _T("To Be Rejected"));
    m_StatusCtrl.SetItemKey(3, TUToBeRejected);
    m_StatusCtrl.SetItemDescription(3, _T("To Be Rejected"));	

    m_StatusCtrl.InsertString(4, _T("Accepted"));
    m_StatusCtrl.SetItemKey(4, TUAccepted);
    m_StatusCtrl.SetItemDescription(4, _T("Accepted"));	

    m_StatusCtrl.InsertString(5, _T("Rejected"));
    m_StatusCtrl.SetItemKey(5, TURejected);
    m_StatusCtrl.SetItemDescription(5, _T("Rejected"));	

    m_StatusCtrl.SelectItemWithKey(TUTaskOpen);
}
void CGetAssgnUpdatesDlg::UpdateListCtrl()
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
	int ResFilter = m_ResNameCtrl.GetSelectedItemsKey();
    int statusFilter = m_StatusCtrl.GetSelectedItemsKey();
   

    m_listCtrl.FillListCtrl_Assgn(pProjReport,  tempWorkUnit, tempGetHoursPerDay, tempGetHoursPerWeek, (int)tempGetDaysPerMonth,
		true, COLUMN_ALL, prjFilter, taskFilter, ResFilter, statusFilter);
    
    m_listCtrl.SelectItem(GUID);

    UpdateMarkAsCompleteRadioBtns();
}
void CGetAssgnUpdatesDlg::OnBnClickedOk()
{
	if(IDOK == AfxMessageBox(_T("The selected updates will be saved in the project plan. Do you want to continue?"),MB_OKCANCEL))
	{
		//unlock SPV files before writing
		m_pUpdatesManager->unLockSPVFiles(m_pUpdatesManager->getProjInfoList()); 

		//write accepted updates to MPP files
        bool bCannotWriteManual = false;
		m_pUpdatesManager->processWritingToFiles(m_pUpdatesManager->getProjInfoList(), bCannotWriteManual, TUAssignmentReportingMode); 

        m_pUpdatesManager->getProjReport()->ChangeAllReportsStatus(TUTaskToBeAccepted, TUTaskAccepted);
        m_pUpdatesManager->getProjReport()->ChangeAllReportsStatus(TUTaskToBeRejected, TUTaskRejected);


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
void CGetAssgnUpdatesDlg::OnBnClickedButtonAccept() 
{
	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
		return;

	bool bMsgShowed = false;
	int itemCount = m_listCtrl.GetItemCount();
	for(int iItem =0; iItem<itemCount; ++iItem)
	{
		if(m_listCtrl.GetCheck(iItem))
		{
			CString strGuid= m_listCtrl.GetItemGUID(iItem);
			CTUDateReport* pDateReport = pProjReport->GetDateReport(strGuid);
			if(pDateReport)
            {
				pDateReport->SetStatus(TUToBeAccepted);
            }
		}
	}

	UpdateListCtrl();
	m_buttonOK.EnableWindow(true);     
}
void CGetAssgnUpdatesDlg::OnBnClickedButtonReject()  
{
	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
		return;

	bool bMsgShowed = false;
	int itemCount = m_listCtrl.GetItemCount();
	for(int iItem =0; iItem<itemCount; ++iItem)
	{
		if(m_listCtrl.GetCheck(iItem))
		{
			CString strGuid= m_listCtrl.GetItemGUID(iItem);
			CTUDateReport* pDateReport = pProjReport->GetDateReport(strGuid);
			if(pDateReport)
            {
				pDateReport->SetStatus(TUToBeRejected);
            }
		}
	}

	UpdateListCtrl();
	m_buttonOK.EnableWindow(true); 
}
void CGetAssgnUpdatesDlg::OnBnClickedButtonReset()
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
void CGetAssgnUpdatesDlg::OnLvnItemchangedList1(NMHDR *pNMHDR, LRESULT *pResult)
{
    LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
   	m_listCtrl.Check_Uncheck_HeaderCheckBox();

    bool bHasCheckItem = m_listCtrl.HasCheckedItem();

    m_AcceptButton.EnableWindow(bHasCheckItem);
    m_RejectButton.EnableWindow(bHasCheckItem);

    CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
    if(pProjReport)
    {
        CString strGUID = m_listCtrl.GetSelectedItemGUID();
        if(!strGUID.IsEmpty())
        {
            CTUDateReport* pReport = pProjReport->GetDateReport(strGUID);
            if(pReport)
            {
                m_editTMComment.SetWindowText(pReport->GetTMNote());
                m_editPMComment.SetWindowText(pReport->GetPMNote());

                CTUAssignmentReport* pAssgnReport = pProjReport->GetDatesAssignmentReport(pReport);
                if (pAssgnReport)
                {
                    EnableMarkAsCompleteGroup(pAssgnReport->IsMarkedAsCompleted());
                    m_iMarkedAsComplete = pAssgnReport->GetMarkAsCompleteType();
                }
            }
        }
    }
    UpdateData(FALSE);
    *pResult = 0;
}
void CGetAssgnUpdatesDlg::OnHdnItemclickList1(NMHDR *pNMHDR, LRESULT *pResult)
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
void CGetAssgnUpdatesDlg::OnNMClickList1(NMHDR *pNMHDR, LRESULT *pResult)
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

    EnableMarkAsCompleteGroup(false);

	CString strGUID = m_listCtrl.GetSelectedItemGUID();
	if(!strGUID.IsEmpty())
	{
		CTUDateReport* pReport = pProjReport->GetDateReport(strGUID);
		if(pReport)
		{
			m_editTMComment.SetWindowText(pReport->GetTMNote());
			m_editPMComment.SetWindowText(pReport->GetPMNote());

            CTUAssignmentReport* pAssgnReport = pProjReport->GetDatesAssignmentReport(pReport);
            if (pAssgnReport)
            {
                EnableMarkAsCompleteGroup(pAssgnReport->IsMarkedAsCompleted());
                m_iMarkedAsComplete = pAssgnReport->GetMarkAsCompleteType();
            }
		}
	}
    UpdateData(FALSE);
    *pResult = 0;
}
void CGetAssgnUpdatesDlg::OnBnClickedCancel()
{
	//unlock SPV files before writing
	m_pUpdatesManager->unLockSPVFiles(m_pUpdatesManager->getProjInfoList()); 

    CDialogEx::OnCancel();
}
void CGetAssgnUpdatesDlg::OnCbnCloseupComboProjectName()
{
	UpdateListCtrl();

	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
		return;

	bool bHasSaveableReport = pProjReport->HasAcceptedReport() || pProjReport->HasRejectedReport();
	m_buttonOK.EnableWindow(bHasSaveableReport);   
}
void CGetAssgnUpdatesDlg::OnCbnCloseupComboTaskName()
{
	UpdateListCtrl();

	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
		return;

	bool bHasSaveableReport = pProjReport->HasAcceptedReport() || pProjReport->HasRejectedReport();
	m_buttonOK.EnableWindow(bHasSaveableReport);   
}
void CGetAssgnUpdatesDlg::OnCbnCloseupComboResourceName()
{
	UpdateListCtrl();

	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
		return;

	bool bHasSaveableReport = pProjReport->HasAcceptedReport() || pProjReport->HasRejectedReport();
	m_buttonOK.EnableWindow(bHasSaveableReport);   
}
void CGetAssgnUpdatesDlg::OnCbnCloseupComboStatus()
{
    UpdateListCtrl();

    CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
    if(!pProjReport)
        return;

    bool bHasSaveableReport = pProjReport->HasAcceptedReport() || pProjReport->HasRejectedReport();
    m_buttonOK.EnableWindow(bHasSaveableReport);   
}
void CGetAssgnUpdatesDlg::OnEnChangeEditPMComment()
{
	CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
	if(!pProjReport)
	{		
		return;
	}

	CString strGUID = m_listCtrl.GetSelectedItemGUID();
	if(!strGUID.IsEmpty())
	{
		CTUDateReport* pReport = pProjReport->GetDateReport(strGUID);
		if(pReport)
		{
			CString str;
			m_editPMComment.GetWindowText(str);

			pReport->SetPMNote(str);
		}
	}

    m_buttonOK.EnableWindow(true);  
}
void CGetAssgnUpdatesDlg::UpdateReportsMarkAsCompleteType()
{
    CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
    if(!pProjReport)
        return;

    CString strGUID = m_listCtrl.GetSelectedItemGUID();
    if(!strGUID.IsEmpty())
    {
        CTUDateReport* pReport = pProjReport->GetDateReport(strGUID);
        if(pReport)
        {
            CTUAssignmentReport* pAssgnReport = pProjReport->GetDatesAssignmentReport(pReport);
            if (pAssgnReport)
            {
                UpdateData(TRUE);
                pAssgnReport->SetMarkAsCompleteType((enumMarkAsCompleteType)m_iMarkedAsComplete); 
                pAssgnReport->GetDateReportsList();
                if (m_iMarkedAsComplete)
                {
                    if (m_iMarkedAsComplete == TU_MAC_Percent || m_iMarkedAsComplete == TU_MAC_Finish)
                    {
                        m_buttonOK.EnableWindow(TRUE);
                    }
                    else
                    {
                        m_buttonOK.EnableWindow(FALSE);
                    }

                    TUDateReportsList::iterator itDateReport;
                    for (itDateReport = pAssgnReport->GetDateReportsList()->begin();
                        itDateReport != pAssgnReport->GetDateReportsList()->end();
                        itDateReport++)
                    {
                        (*itDateReport)->SetStatus(TUToBeAccepted);
                    }
                    UpdateListCtrl();
                }
            }
        }
    }
}
void CGetAssgnUpdatesDlg::EnableMarkAsCompleteGroup(bool MarkAsCompleted)
{
    m_radioNone.EnableWindow(MarkAsCompleted == TRUE);
    m_radioPercent.EnableWindow(MarkAsCompleted == TRUE);
    m_radioFinish.EnableWindow(MarkAsCompleted == TRUE);
    m_staticMACSelectMethod.EnableWindow(MarkAsCompleted == TRUE);
    GetDlgItem(IDC_MACGROUP)->EnableWindow(MarkAsCompleted == TRUE);
    m_staticMACNotificication.ShowWindow(MarkAsCompleted ? SW_SHOW : SW_HIDE);

    if (!MarkAsCompleted)
        m_iMarkedAsComplete = 0;
}
void CGetAssgnUpdatesDlg::UpdateMarkAsCompleteRadioBtns()
{
    CTUProjectReport* pProjReport = m_pUpdatesManager->getProjReport();
    if(!pProjReport)
        return;

    m_iMarkedAsComplete = 0;
    EnableMarkAsCompleteGroup(false);

    CString strGUID = m_listCtrl.GetSelectedItemGUID();
    if(!strGUID.IsEmpty())
    {
        CTUDateReport* pReport = pProjReport->GetDateReport(strGUID);
        if(pReport)
        {
            CTUAssignmentReport* pAssgnReport = pProjReport->GetDatesAssignmentReport(pReport);
            if (pAssgnReport)
            {
                m_iMarkedAsComplete = pAssgnReport->GetMarkAsCompleteType();
                
                EnableMarkAsCompleteGroup(pAssgnReport->IsMarkedAsCompleted());
            }
        }
    }
    UpdateData(FALSE);
}

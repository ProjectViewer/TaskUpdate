#include "TaskUpdate_ListCtrl.h"
#include "Utils.h"

#define HDF_SORTUP              0x0400
#define HDF_SORTDOWN            0x0200

IMPLEMENT_DYNAMIC(CTaskUpdate_ListCtrl, CListCtrl)

CTaskUpdate_ListCtrl::CTaskUpdate_ListCtrl(void)
{
	m_pchTip=NULL;
	m_pwchTip=NULL;
	itemPropsPtrList.clear();
    m_clickState = 0;
}

CTaskUpdate_ListCtrl::~CTaskUpdate_ListCtrl(void)
{
	if(m_pchTip != NULL)
		delete m_pchTip;

	if(m_pwchTip != NULL)
		delete m_pwchTip;

	std::vector<itemPropsStruct*>::iterator it;
	for(it = itemPropsPtrList.begin(); it != itemPropsPtrList.end();)
	{
		delete *it;
		*it = NULL;
		it = itemPropsPtrList.erase(it);
	}
}

BEGIN_MESSAGE_MAP(CTaskUpdate_ListCtrl, CListCtrl)
	ON_NOTIFY_EX(TTN_NEEDTEXTA, 0, OnToolNeedText)
	ON_NOTIFY_EX(TTN_NEEDTEXTW, 0, OnToolNeedText)
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, &CTaskUpdate_ListCtrl::OnNMCustomdraw)
    ON_NOTIFY_REFLECT(NM_CLICK, &CTaskUpdate_ListCtrl::OnNMClick)
END_MESSAGE_MAP()

struct dummy 
{
	CListCtrl *listCtrl;
	bool m_Ascending;
	int iWhich;
	int m_ColumnType;
};

static int CALLBACK Compare(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	dummy* pListCtrl = (dummy*) lParamSort;

	int iWhich = pListCtrl->iWhich;
	bool bAsc = pListCtrl->m_Ascending;
	int m_colType = pListCtrl->m_ColumnType;
	CListCtrl *listCtrl = pListCtrl->listCtrl;

	switch(pListCtrl->m_ColumnType)
	{
	case SORT_STR:
		{
			TCHAR left[256] = _T(""), right[256] = _T("");
			ListView_GetItemText(listCtrl->GetSafeHwnd(), lParam1, pListCtrl->iWhich, left, sizeof(left));
			ListView_GetItemText(listCtrl->GetSafeHwnd(), lParam2, pListCtrl->iWhich, right, sizeof(right));	

			return (pListCtrl->m_Ascending) ? _tcscmp( right, left ) : _tcscmp( left, right );	
		}
		break;

	case SORT_DAT:
		{
			TCHAR left[256] = _T(""), right[256] = _T("");
			ListView_GetItemText(listCtrl->GetSafeHwnd(), lParam1, pListCtrl->iWhich, left, sizeof(left));
			ListView_GetItemText(listCtrl->GetSafeHwnd(), lParam2, pListCtrl->iWhich, right, sizeof(right));	

			COleDateTime odt_left, odt_right;
			odt_left.ParseDateTime(left, LOCALE_NOUSEROVERRIDE,LANG_USER_DEFAULT); 
			odt_right.ParseDateTime(right, LOCALE_NOUSEROVERRIDE,LANG_USER_DEFAULT);
			int i_left = odt_left.m_status == 0 ? static_cast<int>(odt_left) : 0;
			int i_right = odt_right.m_status == 0 ? static_cast<int>(odt_right) : 0;
			int dif  = i_left - i_right;

			if (dif==0) return 0;

			if (pListCtrl->m_Ascending)
			{
				if (dif>0) return -1;
				if (dif<0) return 1;
			}
			else
			{
				if (dif>0) return 1;
				if (dif<0) return -1;
			}
		}
		break;

	case SORT_NUM:
		{
			TCHAR left[256] = _T(""), right[256] = _T("");
			ListView_GetItemText(listCtrl->GetSafeHwnd(), lParam1, pListCtrl->iWhich, left, sizeof(left));
			ListView_GetItemText(listCtrl->GetSafeHwnd(), lParam2, pListCtrl->iWhich, right, sizeof(right));	

			double dif  = static_cast<double>(_wtof(left)) - static_cast<double>(_wtof(right));

			if (dif==0) return 0;
			if (pListCtrl->m_Ascending)
			{
				if (dif>0) return -1;
				if (dif<0) return 1;
			}
			else
			{
				if (dif>0) return 1;
				if (dif<0) return -1;
			}	
		}
		break;
	}

	return 0;

}

namespace {
	LRESULT EnableWindowTheme(HWND hwnd, LPCWSTR classList, LPCWSTR subApp, LPCWSTR idlist)
	{
		LRESULT lResult = S_FALSE;
		HMODULE hinstDll;
		BOOL (WINAPI *pIsThemeActive)();
		HRESULT (WINAPI *pSetWindowTheme)(HWND hwnd, LPCWSTR pszSubAppName, LPCWSTR pszSubIdList);
		HANDLE (WINAPI *pOpenThemeData)(HWND hwnd, LPCWSTR pszClassList);
		HRESULT (WINAPI *pCloseThemeData)(HANDLE hTheme);

		// Check if running on Windows XP or newer
		hinstDll = ::LoadLibrary(_T("UxTheme.dll"));
		if (hinstDll)
		{
			// Check if theme service is running
			(FARPROC&)pIsThemeActive = ::GetProcAddress( hinstDll, "IsThemeActive" );
			if( pIsThemeActive && pIsThemeActive() )
			{
				(FARPROC&)pOpenThemeData = ::GetProcAddress(hinstDll, "OpenThemeData");
				(FARPROC&)pCloseThemeData = ::GetProcAddress(hinstDll, "CloseThemeData");
				(FARPROC&)pSetWindowTheme = ::GetProcAddress(hinstDll, "SetWindowTheme");
				if (pSetWindowTheme && pOpenThemeData && pCloseThemeData)			
				{
					// Check is themes is available for the application
					HANDLE hTheme = pOpenThemeData(hwnd,classList);
					if (hTheme!=NULL)
					{
						VERIFY(pCloseThemeData(hTheme)==S_OK);
						// Enable Windows Theme Style
						lResult = pSetWindowTheme(hwnd, subApp, idlist);
					}
				}
			}
			::FreeLibrary(hinstDll);
		}
		return lResult;
	}
}
void				CTaskUpdate_ListCtrl::PreSubclassWindow()
{
	CListCtrl::PreSubclassWindow();

		// Focus retangle is not painted properly without double-buffering
#if (_WIN32_WINNT >= 0x501)
	SetExtendedStyle(LVS_EX_DOUBLEBUFFER | GetExtendedStyle());
#endif
	SetExtendedStyle(GetExtendedStyle() | LVS_EX_FULLROWSELECT);
	SetExtendedStyle(GetExtendedStyle() | LVS_EX_HEADERDRAGDROP);
	SetExtendedStyle(GetExtendedStyle() | LVS_EX_GRIDLINES);
	
	// Enable Vista-look if possible
	EnableWindowTheme(m_hWnd, L"ListView", L"Explorer", NULL);

	// Disable the CToolTipCtrl of CListCtrl so it won't disturb the CWnd tooltip
	GetToolTips()->Activate(FALSE);

	// Activates the standard CWnd tooltip functionality
	VERIFY( EnableToolTips(TRUE) );
}
//////////////////////////////////////////////////////////////////////////
bool				CTaskUpdate_ListCtrl::SortColumn(int columnIndex, bool ascending,int columnType)
{
	dummy temp;
	temp.m_Ascending = ascending;
	temp.iWhich = columnIndex;
	temp.m_ColumnType = columnType;
	temp.listCtrl=this;

	ListView_SortItemsEx(this->GetSafeHwnd(), Compare, (LPARAM) &temp);

	SetSortArrow(columnIndex,ascending);
	return true;
}  
void				CTaskUpdate_ListCtrl::SetSortArrow(int colIndex, bool ascending)
{
	for(int i = 0; i < GetHeaderCtrl()->GetItemCount(); ++i)
	{
		HDITEM hditem = {0};
		hditem.mask = HDI_FORMAT;
		VERIFY( GetHeaderCtrl()->GetItem( i, &hditem ) );
		hditem.fmt &= ~(HDF_SORTDOWN|HDF_SORTUP);
		if (i == colIndex)
		{
			hditem.fmt |= ascending ? HDF_SORTDOWN : HDF_SORTUP;
		}
		VERIFY( GetHeaderCtrl()->SetItem( i, &hditem ) );
	}
}
bool				CTaskUpdate_ListCtrl::AllItemsAreChecked()
{
	int nCount = GetItemCount();
	for(int nItem = 0; nItem < nCount; nItem++)
	{
		if ( !GetCheck(nItem))
		{
			return false;
		}
	}
	return true;
}
bool				CTaskUpdate_ListCtrl::HasCheckedItem()
{
	int nCount = GetItemCount();
	for(int nItem = 0; nItem < nCount; nItem++)
	{
		if (GetCheck(nItem))
		{
			return true;
		}
	}
	return false;
}
void				CTaskUpdate_ListCtrl::CheckAllItems()
{
	int nCount = GetItemCount();
	for(int nItem = 0; nItem < nCount; nItem++)
	{
		SetCheck(nItem,true);
	}
}
void				CTaskUpdate_ListCtrl::UncheckAllItems()
{
	int nCount = GetItemCount();
	while (HasCheckedItem())
	{
		for(int nItem = 0; nItem < nCount; nItem++)
			SetCheck(nItem,false);
		for(int nItem = nCount-1; nItem >=0; nItem--)
			SetCheck(nItem,false);
	}
}
DWORD   			CTaskUpdate_ListCtrl::GetOSMajorVersion()
{
    OSVERSIONINFO osvi;
    ZeroMemory(&osvi, sizeof(OSVERSIONINFO));
    osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
    GetVersionEx(&osvi);
    return osvi.dwMajorVersion;
}
void				CTaskUpdate_ListCtrl::CreateHeaderCheckBox()
{
	bmp_check.LoadBitmap(IDB_CT_CHECKBOXES);
	ImgHeaders.Create(13,13,ILC_COLOR24,4,4);
	ImgHeaders.Add(&bmp_check,RGB(255,0,255));
	GetHeaderCtrl()->SetImageList(&ImgHeaders);

	HDITEM hdi;
	hdi.mask = HDI_IMAGE | HDI_FORMAT;
	hdi.fmt =  HDF_IMAGE;
	hdi.iImage = GetOSMajorVersion()>=6 ? 0 : 2;
	GetHeaderCtrl()->SetItem( 0, &hdi);
}
void				CTaskUpdate_ListCtrl::Check_Uncheck_HeaderCheckBox()
{
	HDITEM hdi;
	hdi.mask = HDI_IMAGE;
	hdi.iImage = (AllItemsAreChecked() && !IsEmptyList()) ? (GetOSMajorVersion()>=6 ? 1 : 3) : (GetOSMajorVersion()>=6 ? 0 : 2);
	GetHeaderCtrl()->SetItem(0, &hdi);
}
void				CTaskUpdate_ListCtrl::Check_Uncheck_AllItems()
{
	HDITEM hdi;
	hdi.mask = HDI_IMAGE | HDI_FORMAT;
	hdi.fmt =  HDF_IMAGE;
	int nCount = GetItemCount();
	if (AllItemsAreChecked())
	{	
		hdi.iImage =GetOSMajorVersion()>=6 ? 0 : 2;
		UncheckAllItems();
	}
	else
	{		
		hdi.iImage =GetOSMajorVersion()>=6 ? 1 : 3;
		CheckAllItems();
	}
	GetHeaderCtrl()->SetItem( 0, &hdi);
}
bool				CTaskUpdate_ListCtrl::IsEmptyList()
{
	return GetItemCount()==0 ? true : false;
}
int					CTaskUpdate_ListCtrl::GetCheckedItemCount()
{
	UINT i=0;
	int nCount = GetItemCount();
	for(int nItem = 0; nItem < nCount; nItem++)
	{
		if (GetCheck(nItem))
			i++;
	}
	return i;
}
std::vector<int>	CTaskUpdate_ListCtrl::GetCheckedItemsList()
{
	std::vector<int> theList;
	for(int i=0; i<GetItemCount(); ++i)
	{
		if(GetCheck(i))
			theList.push_back(i);
	}

	return theList;
}
//////////////////////////////////////////////////////////////////////////
void				CTaskUpdate_ListCtrl::FillListCtrl_Task(	
    CTUProjectReport* pProjReport, 
    enumTUUnit appUnit,  
    double hoursPerDay, 
    double hoursPerWeek, 
    int daysPerMonth, 
    int iInsertColumns, 
    CString strProjectFilter,
    int	iTaskFilter,
    int iStatusFilter)
{
    int colNum = 0;
    CString str;

    ClearListCtrl();

    InsertColumn(colNum,_T(""),LVCFMT_CENTER,((iInsertColumns & COLUMN_CHECKBOX) || (iInsertColumns & COLUMN_ALL)) ? 30 : 0,-1);
    colNum++;

    if((iInsertColumns & COLUMN_NAME) || (iInsertColumns & COLUMN_ALL))
    {
        str.LoadString(IDS_Name);
        InsertColumn(colNum,str,LVCFMT_LEFT,150,-1);
        colNum++;
    }
    if((iInsertColumns & COLUMN_CURR_PERC_COMPL) || (iInsertColumns & COLUMN_ALL))
    {
        str.LoadString(IDS_CurrentPercentComplete);
        InsertColumn(colNum,str,LVCFMT_LEFT,150,-1);
        colNum++;
    }
    if((iInsertColumns & COLUMN_REP_PERC_COMPL) || (iInsertColumns & COLUMN_ALL))
    {
        str.LoadString(IDS_ReportedPercentComplete);
        InsertColumn(colNum,str,LVCFMT_LEFT,150,-1);
        colNum++;
    }
    if((iInsertColumns & COLUMN_CURR_PERC_WORK_COMPL) || (iInsertColumns & COLUMN_ALL))
    {
        str.LoadString(IDS_CurrentPercentWorkComplete);
        InsertColumn(colNum,str,LVCFMT_LEFT,175,-1);
        colNum++;
    }
    if((iInsertColumns & COLUMN_REP_PERC_WORK_COMPL) || (iInsertColumns & COLUMN_ALL))
    {
        str.LoadString(IDS_ReportedPercentWorkComplete);
        InsertColumn(colNum,str,LVCFMT_LEFT,175,-1);
        colNum++;
    }
    if((iInsertColumns & COLUMN_CURR_ACT_WORK) || (iInsertColumns & COLUMN_ALL))
    {
        str.LoadString(IDS_CurrentActualWork);
        InsertColumn(colNum,str,LVCFMT_LEFT,120,-1);
        colNum++;
    }
    if((iInsertColumns & COLUMN_REP_ACT_WORK) || (iInsertColumns & COLUMN_ALL))
    {
        str.LoadString(IDS_ReportedActualWork);
        InsertColumn(colNum,str,LVCFMT_LEFT,120,-1);
        colNum++;
    }
	if((iInsertColumns & COLUMN_CURR_ACT_START) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_CurrentActualStart);
		InsertColumn(colNum,str,LVCFMT_LEFT,130,-1);
		colNum++;
	}
	if((iInsertColumns & COLUMN_REP_ACT_START) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_ReportedActualStart);
		InsertColumn(colNum,str,LVCFMT_LEFT,130,-1);
		colNum++;
	}
	if((iInsertColumns & COLUMN_CURR_ACT_FINISH) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_CurrentActualFinish);
		InsertColumn(colNum,str,LVCFMT_LEFT,130,-1);
		colNum++;
	}
	if((iInsertColumns & COLUMN_REP_ACT_FINISH) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_ReportedActualFinish);
		InsertColumn(colNum,str,LVCFMT_LEFT,130,-1);
		colNum++;
	}
    if((iInsertColumns & COLUMN_CURR_ACT_DURATION) || (iInsertColumns & COLUMN_ALL))
    {
        str.LoadString(IDS_CurrentActualDuration);
        InsertColumn(colNum,str,LVCFMT_LEFT,130,-1);
        colNum++;
    }
    if((iInsertColumns & COLUMN_REP_ACT_DURATION) || (iInsertColumns & COLUMN_ALL))
    {
        str.LoadString(IDS_ReportedActualDuration);
        InsertColumn(colNum,str,LVCFMT_LEFT,140,-1);
        colNum++;
    }

    int i=0;
    TUTaskReportsList::iterator itTaskReports;
    for(itTaskReports = pProjReport->GetTaskReportsList()->begin();
        itTaskReports != pProjReport->GetTaskReportsList()->end();
        ++itTaskReports)
    {
        if ((*itTaskReports)->IsUndefined())
            continue;

        if(strProjectFilter != ((*itTaskReports)->GetProjectGUID()) && strProjectFilter != ALL_PROJECTS)
            continue;

        if(iTaskFilter != ((*itTaskReports)->GetTaskUID()) && iTaskFilter != ALL_TASKS)
            continue;

        if(iStatusFilter != ((*itTaskReports)->GetStatus()) && iStatusFilter != ALL_STATUSES)
        {
            if (!(((*itTaskReports)->GetStatus() == TUTaskToBeAccepted || (*itTaskReports)->GetStatus() == TUTaskToBeRejected) && iStatusFilter == TUTaskOpen))
                continue;
        }

        int colNum = 1;

        CString strGUID;
        CString strTaskName;
        CString strCurrPercCompl;
        CString strRepPercCompl;
        CString strCurrPercWorkCompl;
        CString strRepPercWorkCompl; 
        CString strCurrActWork;
        CString strRepActWork;
		CString strCurrActStart;
		CString strRepActStart;
		CString strCurrActFinish;
		CString strRepActFinish;
        CString strCurrActDuration;
        CString strRepActDuration;
        int iVal;

        strGUID = (*itTaskReports)->GetReportGUID();

        strTaskName = (*itTaskReports)->GetTaskName();

        iVal = (*itTaskReports)->GetCurrentPercentComplete();
        strCurrPercCompl.Format(_T("%d%%"), iVal);

        iVal = (*itTaskReports)->GetReportedPercentComplete();
        strRepPercCompl.Format(_T("%d%%"), iVal);

        iVal = (*itTaskReports)->GetCurrentPercentWorkComplete();
        strCurrPercWorkCompl.Format(_T("%d%%"), iVal);

        iVal = (*itTaskReports)->GetReportedPercentWorkComplete();
        strRepPercWorkCompl.Format(_T("%d%%"), iVal);

        strCurrActWork = Utils::FormatWork((*itTaskReports)->GetCurrentActualWork());
        strRepActWork = Utils::FormatWork((*itTaskReports)->GetReportedActualWork());

		CTime tmCurrActStart = /*Utils::toGMT(*/(*itTaskReports)->GetCurrentActualStart()/*)*/;
		CTime tmRepActStart = Utils::toGMT((*itTaskReports)->GetReportedActualStart());
		CTime tmCurrActFinish = /*Utils::toGMT(*/(*itTaskReports)->GetCurrentActualFinish()/*)*/;
		CTime tmRepActFinish = Utils::toGMT((*itTaskReports)->GetReportedActualFinish());

		strCurrActStart = tmCurrActStart != CTime(0)
			? Utils::formatMSTime(tmCurrActStart)
			: _T("NA");

		strRepActStart = tmRepActStart != CTime(0)
			? Utils::formatMSTime(tmRepActStart)
			: _T("NA");

		strCurrActFinish = tmCurrActFinish != CTime(0)
			? Utils::formatMSTime(tmCurrActFinish)
			: _T("NA");

		strRepActFinish = tmRepActFinish != CTime(0)
			? Utils::formatMSTime(tmRepActFinish)
			: _T("NA");

        strCurrActDuration = Utils::FormatWork((*itTaskReports)->GetCurrentActualDuration());
        strRepActDuration = Utils::FormatWork((*itTaskReports)->GetReportedActualDuration());

        if((iInsertColumns & COLUMN_NAME) || (iInsertColumns & COLUMN_ALL))
        {
            InsertItem(i,_T(""));
            SetItemText(i,colNum,strTaskName);
            colNum++;
        }
        if((iInsertColumns & COLUMN_CURR_PERC_COMPL) || (iInsertColumns & COLUMN_ALL))
        {
            SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskPercentCompleteReport ? strCurrPercCompl : _T(""));
            colNum++;
        }
        if((iInsertColumns & COLUMN_REP_PERC_COMPL) || (iInsertColumns & COLUMN_ALL))
        {
            SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskPercentCompleteReport ? strRepPercCompl : _T(""));
            colNum++;
        }
        if((iInsertColumns & COLUMN_CURR_PERC_WORK_COMPL) || (iInsertColumns & COLUMN_ALL))
        {
            SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskPercentWorkCompleteReport ? strCurrPercWorkCompl : _T(""));
            colNum++;
        }
        if((iInsertColumns & COLUMN_REP_PERC_WORK_COMPL) || (iInsertColumns & COLUMN_ALL))
        {
            SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskPercentWorkCompleteReport ? strRepPercWorkCompl : _T(""));
            colNum++;
        }
        if((iInsertColumns & COLUMN_CURR_ACT_WORK) || (iInsertColumns & COLUMN_ALL))
        {
            SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskActualWorkReport ? strCurrActWork : _T(""));
            colNum++;
        }
        if((iInsertColumns & COLUMN_REP_ACT_WORK) || (iInsertColumns & COLUMN_ALL))
        {
            SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskActualWorkReport ? strRepActWork : _T(""));
            colNum++;
        }
		if((iInsertColumns & COLUMN_CURR_ACT_START) || (iInsertColumns & COLUMN_ALL))
		{
			SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskActualStartReport ? strCurrActStart : _T(""));
			colNum++;
		}
		if((iInsertColumns & COLUMN_CURR_ACT_START) || (iInsertColumns & COLUMN_ALL))
		{
			SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskActualStartReport ? strRepActStart : _T(""));
			colNum++;
		}
		if((iInsertColumns & COLUMN_CURR_ACT_FINISH) || (iInsertColumns & COLUMN_ALL))
		{
			SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskActualFinishReport ? strCurrActFinish : _T(""));
			colNum++;
		}
		if((iInsertColumns & COLUMN_CURR_ACT_FINISH) || (iInsertColumns & COLUMN_ALL))
		{
			SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskActualFinishReport ? strRepActFinish : _T(""));
			colNum++;
		}
        if((iInsertColumns & COLUMN_CURR_ACT_DURATION) || (iInsertColumns & COLUMN_ALL))
        {
            SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskActualDurationReport ? strCurrActDuration : _T(""));
            colNum++;
        }
        if((iInsertColumns & COLUMN_REP_ACT_DURATION) || (iInsertColumns & COLUMN_ALL))
        {
            SetItemText(i,colNum,(*itTaskReports)->GetWorkInsertMode() == TUTaskActualDurationReport ? strRepActDuration : _T(""));
            colNum++;
        }

        COLORREF bkClr = CLR_DEFAULT;
        COLORREF txtClr = CLR_DEFAULT;

        if((*itTaskReports)->GetStatus()==TUTaskAccepted)
            bkClr = RGB(150,255,200);
        else if((*itTaskReports)->GetStatus()==TUTaskToBeAccepted)
            bkClr = RGB(200,255,230);
        else if((*itTaskReports)->GetStatus()==TUTaskRejected)
            bkClr = RGB(255,150,200);
        else if((*itTaskReports)->GetStatus()==TUTaskToBeRejected)
            bkClr = RGB(255,200,230);

        SetItemProps(	
            i, 
            (*itTaskReports)->GetWorkInsertMode(),
            bkClr, 
            txtClr,
            strGUID,
            strTaskName,
            strCurrPercCompl,
            strRepPercCompl,
            strCurrPercWorkCompl,
            strRepPercWorkCompl,
            strCurrActWork,
            strRepActWork,
			strCurrActStart,
			strRepActStart,
			strCurrActFinish,
			strRepActFinish,
            strCurrActDuration,
            strRepActDuration);
        i++;
    }
}
void				CTaskUpdate_ListCtrl::FillListCtrl_Assgn(	
    CTUProjectReport* pProjReport, 
	enumTUUnit appUnit,  
	double hoursPerDay, 
	double hoursPerWeek, 
	int daysPerMonth, 
	bool bGrouping, 
	int iInsertColumns, 
	CString strProjectFilter,
	int	iTaskFilter,
	int iResourceFilter,
    int iStatusFilter)
{
	int colNum = 0;
	CString str;

	ClearListCtrl();

	InsertColumn(colNum,_T(""),LVCFMT_CENTER,((iInsertColumns & COLUMN_CHECKBOX) || (iInsertColumns & COLUMN_ALL)) ? 30 : 0,-1);
	colNum++;

	if((iInsertColumns & COLUMN_NAME) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_Name);
		InsertColumn(colNum,str,LVCFMT_LEFT,110,-1);
		colNum++;
	}

	if((iInsertColumns & COLUMN_DATE) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_Date);
		InsertColumn(colNum,str,LVCFMT_LEFT,70,-1);
		colNum++;
	}

	if((iInsertColumns & COLUMN_WORK) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_Work);
		InsertColumn(colNum,str,LVCFMT_CENTER,50,-1);
		colNum++;
	}

	if((iInsertColumns & COLUMN_ACT_WORK) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_Act_Work);
		InsertColumn(colNum,str,LVCFMT_CENTER,60,-1);
		colNum++;
	}

	if((iInsertColumns & COLUMN_OVT_WORK) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_Ovr_Work);
		InsertColumn(colNum,str,LVCFMT_CENTER,65,-1);
		colNum++;
	}

	if((iInsertColumns & COLUMN_ACT_OVT_WORK) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_REPORTEDOVTWORK);
		InsertColumn(colNum,str,LVCFMT_CENTER,85,-1);
		colNum++;
	}

	if((iInsertColumns & COLUMN_LAST_PERC) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_LASTPERCCOMPL);
		InsertColumn(colNum,str,LVCFMT_CENTER,110,-1);
		colNum++;
	}

	if((iInsertColumns & COLUMN_REP_PERC) || (iInsertColumns & COLUMN_ALL))
	{
		str.LoadString(IDS_REPPERCCOMPL);
		InsertColumn(colNum,str,LVCFMT_CENTER,115,-1);
		colNum++;
	}

	std::map<int,CString> taskMap;
	std::map<int,CString>::iterator itTaskMap;
	TUAssignmentReportsList::iterator itAssgnReports;
	TUDateReportsList::iterator itDateReports;
	for(itAssgnReports = pProjReport->GetAssignmentReportsList()->begin();
		itAssgnReports != pProjReport->GetAssignmentReportsList()->end();
		++itAssgnReports)
	{
		itTaskMap = taskMap.find((int)(*itAssgnReports)->GetTaskUID());
		if(itTaskMap == taskMap.end())
		{
			std::pair<int,CString> temp;
			temp.first = (int)(*itAssgnReports)->GetTaskUID();
			temp.second = (*itAssgnReports)->GetTaskName();
			taskMap.insert(temp);
		}
	}

    if(bGrouping)
    {
		EnableGroupView(true);
 		for (itTaskMap = taskMap.begin(); itTaskMap != taskMap.end(); ++itTaskMap)
 		{
 			CString strGroupName = itTaskMap->second;;

 			LVGROUP lvg;

 			ZeroMemory(&lvg, sizeof(lvg));
 			lvg.cbSize = sizeof(lvg);
 			lvg.mask = LVGF_HEADER |LVGF_GROUPID | LVGF_ALIGN | LVGF_STATE;
 			lvg.pszHeader = strGroupName.GetBuffer();
 			lvg.cchHeader = (int)wcslen(lvg.pszHeader);
 			lvg.iGroupId = itTaskMap->first;
 			lvg.uAlign = LVGA_HEADER_LEFT;
			lvg.state = GetOSMajorVersion()>=6 ? LVGS_COLLAPSIBLE : LVGS_NORMAL;
			
 			InsertGroup(itTaskMap->first, &lvg);
 		}
    }
	int i=0;

	for(itAssgnReports = pProjReport->GetAssignmentReportsList()->begin();
		itAssgnReports != pProjReport->GetAssignmentReportsList()->end();
		++itAssgnReports)
	{
		for(itDateReports = (*itAssgnReports)->GetDateReportsList()->begin();
			itDateReports != (*itAssgnReports)->GetDateReportsList()->end();
			++itDateReports)
		{
			if(strProjectFilter != ((*itAssgnReports)->GetProjectGUID()) && strProjectFilter != ALL_PROJECTS)
				continue;

			if(iTaskFilter != ((*itAssgnReports)->GetTaskUID()) && iTaskFilter != ALL_TASKS)
				continue;

			if(iResourceFilter != ((*itAssgnReports)->GetResourceUID()) && iResourceFilter != ALL_RESOURCES)
				continue;

            if(iStatusFilter != ((*itDateReports)->GetStatus()) && iStatusFilter != ALL_STATUSES)
            {
                if (!(((*itDateReports)->GetStatus() == TUTaskToBeAccepted || (*itDateReports)->GetStatus() == TUTaskToBeRejected) && iStatusFilter == TUTaskOpen))
                    continue;
            }

			int colNum = 1;

			CString strGUID, strTaskName, strResName, strDate, strWork, strActWork, strOvtWork, strActOvtWork,strCurrPerc, strRepPerc, strMarkAsCompl;
			int iVal;
			bool bVal;

            bool hasRepWork = -1 != (*itDateReports)->GetReportedWork();
            bool hasRepOvtWork = -1 != (*itDateReports)->GetReportedOvertimeWork();

			strGUID = (*itDateReports)->GetReportGUID();

			strTaskName = (*itAssgnReports)->GetTaskName();

			strResName = (*itAssgnReports)->GetResourceName();

			CTime tm = (*itDateReports)->GetReportDate();
			CTime tm2 = CTime(Utils::GetYear(tm), Utils::GetMonth(tm), Utils::GetDay(tm), 0, 0, 0);

			strDate = tm2 != CTime(0)
				? Utils::formatMSTime(tm2)
				: _T("NA");

			strWork = hasRepWork ? Utils::FormatWork((*itDateReports)->GetPlannedWork()) : _T("");

			strActWork = hasRepWork ? Utils::FormatWork((*itDateReports)->GetReportedWork()) : _T("");

			strOvtWork = hasRepOvtWork ? Utils::FormatWork((*itDateReports)->GetPlannedOvertimeWork()) : _T("");

			strActOvtWork = hasRepOvtWork ? Utils::FormatWork((*itDateReports)->GetReportedOvertimeWork()) : _T("");

			iVal = (*itDateReports)->GetCurrentPercentComplete();
			strCurrPerc.Format(_T("%d%%"), iVal);

			iVal = (*itDateReports)->GetReportedPercentComplete();
			strRepPerc.Format(_T("%d%%"), iVal);

			bVal = (*itAssgnReports)->GetWorkInsertMode()==TUAssgnTimeReports
					? (*itAssgnReports)->GetMarkedAsCompletedForDateReport(*itDateReports)
					: (*itAssgnReports)->IsMarkedAsCompleted();

			CString strYes, strNo;
			strYes.LoadString(IDS_Yes);
			strNo.LoadString(IDS_NO);
			strMarkAsCompl = bVal ? strYes : strNo;

			if((iInsertColumns & COLUMN_NAME) || (iInsertColumns & COLUMN_ALL))
			{
				InsertItem(i,_T(""));
				SetItemText(i,colNum,strResName);
				colNum++;
			}

			if((iInsertColumns & COLUMN_DATE) || (iInsertColumns & COLUMN_ALL))
			{
				SetItemText(i,colNum,(*itAssgnReports)->GetWorkInsertMode()==TUAssgnTimeReports ? strDate : _T(""));
				colNum++;
			}

			if((iInsertColumns & COLUMN_WORK) || (iInsertColumns & COLUMN_ALL))
			{
				SetItemText(i,colNum,(*itAssgnReports)->GetWorkInsertMode()==TUAssgnTimeReports ? strWork : _T(""));
				colNum++;
			}

			if((iInsertColumns & COLUMN_ACT_WORK) || (iInsertColumns & COLUMN_ALL))
			{
				SetItemText(i,colNum,(*itAssgnReports)->GetWorkInsertMode()==TUAssgnTimeReports ? strActWork : _T(""));
				colNum++;
			}

			if((iInsertColumns & COLUMN_OVT_WORK) || (iInsertColumns & COLUMN_ALL))
			{
				SetItemText(i,colNum,(*itAssgnReports)->GetWorkInsertMode()==TUAssgnTimeReports ? strOvtWork : _T(""));
				colNum++;
			}

			if((iInsertColumns & COLUMN_ACT_OVT_WORK) || (iInsertColumns & COLUMN_ALL))
			{
				SetItemText(i,colNum,(*itAssgnReports)->GetWorkInsertMode()==TUAssgnTimeReports ? strActOvtWork : _T(""));
				colNum++;
			}

			if((iInsertColumns & COLUMN_LAST_PERC) || (iInsertColumns & COLUMN_ALL))
			{
				SetItemText(i,colNum,(*itAssgnReports)->GetWorkInsertMode()==TUAssgnPercentReports ? strCurrPerc : _T(""));
				colNum++;
			}

			if((iInsertColumns & COLUMN_REP_PERC) || (iInsertColumns & COLUMN_ALL))
			{
				SetItemText(i,colNum,(*itAssgnReports)->GetWorkInsertMode()==TUAssgnPercentReports ? strRepPerc : _T(""));
				colNum++;
			}

			COLORREF bkClr = CLR_DEFAULT;
			COLORREF txtClr = CLR_DEFAULT;

            if((*itDateReports)->GetStatus()==TUAccepted)
                bkClr = RGB(150,255,200);
            else if((*itDateReports)->GetStatus()==TUToBeAccepted)
                bkClr = RGB(200,255,230);
            else if((*itDateReports)->GetStatus()==TURejected)
                bkClr = RGB(255,150,200);
            else if((*itDateReports)->GetStatus()==TUToBeRejected)
                bkClr = RGB(255,200,230);

			SetItemProps(	i,
							(*itAssgnReports)->GetWorkInsertMode(),
							bkClr,
							txtClr,
							strGUID,
							strTaskName,
							strResName,
							strDate,
							strWork,
							strActWork,
							strOvtWork,
							strActOvtWork,
							strCurrPerc,
							strRepPerc,
							strMarkAsCompl);

			if(bGrouping)// Set item group ID
			{
				LVITEM lvi;
				ZeroMemory(&lvi, sizeof(lvi));
				lvi.iItem = i;
				lvi.mask = LVIF_GROUPID;
				lvi.iGroupId = (int)(*itAssgnReports)->GetTaskUID();
				SetItem(&lvi);
			}

			i++;
		}
	}
}

void				CTaskUpdate_ListCtrl::ClearListCtrl()
{
	DeleteAllItems();

	// Delete all of the columns.
	int nColumnCount = GetHeaderCtrl()->GetItemCount();
	for (int i=0;i < nColumnCount;i++)
		DeleteColumn(0);
}
//////////////////////////////////////////////////////////////////////////
BOOL				CTaskUpdate_ListCtrl::OnToolNeedText( UINT id, NMHDR * pNMHDR, LRESULT * pResult )
{
	if(pNMHDR->idFrom == -1)
		return FALSE;

	CToolTipCtrl* pToolTip = AfxGetModuleThreadState()->m_pToolTip;
	if (!pToolTip)
		return FALSE;

	pToolTip->SetMaxTipWidth(SHRT_MAX);

	TOOLTIPTEXTA* pTTTA = (TOOLTIPTEXTA*)pNMHDR;
	TOOLTIPTEXTW* pTTTW = (TOOLTIPTEXTW*)pNMHDR;

	CPoint pt(GetMessagePos());
	ScreenToClient(&pt);
 
	int nRow, nCol;
	CellHitTest(pt, nRow, nCol);
	CString strTip,strTitle;
 	if(!GetToolTipText(nRow,strTitle,strTip))
		return FALSE;

	pToolTip->SetTitle(NULL, strTitle);

   #ifndef _UNICODE
   	if(pNMHDR->code == TTN_NEEDTEXTA)
   	{
   		if(m_pchTip != NULL)
   			delete m_pchTip;
   
   		m_pchTip = new TCHAR[strTip.GetLength()+1];
   		lstrcpyn(m_pchTip, strTip, strTip.GetLength()+1);
   		m_pchTip[strTip.GetLength()] = 0;
   		pTTTW->lpszText = (WCHAR*)m_pchTip;
   	}
   	else
   	{
   		if(m_pwchTip != NULL)
   			delete m_pwchTip;
   
   		m_pwchTip = new WCHAR[strTip.GetLength()+1];
   		_mbstowcsz(m_pwchTip, strTip, strTip.GetLength()+1);
   		m_pwchTip[strTip.GetLength()] = 0; // end of text
   		pTTTW->lpszText = (WCHAR*)m_pwchTip;
   	}
	#else
   	if(pNMHDR->code == TTN_NEEDTEXTA)
   	{
   		if(m_pchTip != NULL)
   			delete m_pchTip;
   
   		m_pchTip = new CHAR[strTip.GetLength()+1];
   		_wcstombsz(m_pchTip, strTip, strTip.GetLength()+1);
   		m_pchTip[strTip.GetLength()] = 0; // end of text
   		pTTTA->lpszText = (LPSTR)m_pchTip;
   	}
   	else
   	{
   		if(m_pwchTip != NULL)
   			delete m_pwchTip;
   
   		m_pwchTip = new WCHAR[strTip.GetLength()+1];
   		lstrcpyn(m_pwchTip, strTip, strTip.GetLength()+1);
   		m_pwchTip[strTip.GetLength()] = 0;
   		pTTTA->lpszText = (LPSTR) m_pwchTip;
   	}
	#endif
	*pResult = 0;
  
	return TRUE;    // message was handled
}
bool				CTaskUpdate_ListCtrl::GetToolTipText(int nRow, CString& strTitle, CString& strTip)
 {
	if (nRow!=-1)
	{
		itemPropsStruct* itemProps = GetItemProps(nRow);
		if(itemProps)
		{
			if(itemProps->eAssgnType == TUAssgnTimeReports)
			{
				strTitle =	_T("   Task Name: ") + itemProps->strTask;
								
				strTip =	_T("   Resource Name: ") + itemProps->strResource + _T("\n\n")+
							_T("   Report Date: ") + itemProps->strDate + _T("\n")+ 
							_T("   Work: ") + itemProps->strWork + _T("\n")+
							_T("   Actual Work: ") + itemProps->strActWork + _T("\n")+
							_T("   Ovtertime Work: ") + itemProps->strOvtWork + _T("\n")+
							_T("   Actual Overtime Work: ") + itemProps->strActOvtWork + _T("\n\n")+
							_T("   Marked As Completed: ") + itemProps->strMarkAsCompl;

				return true;
			}
			else if(itemProps->eAssgnType == TUAssgnPercentReports)
			{
				strTitle =	_T("   Task Name: ") + itemProps->strTask;

				strTip =	_T("   Resource Name: ") + itemProps->strResource + _T("\n\n")+
							_T("   Last % Work Complete: ") + itemProps->strCurrPerc + _T("\n")+
							_T("   Reported % Work Complete: ") + itemProps->strRepPerc + _T("\n");

				return true;
			}
            else if(itemProps->eAssgnType == TUTaskPercentCompleteReport)
			{
				strTitle =	_T("   Task Name: ") + itemProps->strTask;

				strTip =	_T("   Current % Complete: ") + itemProps->strCurrPercCompl + _T("\n")+
							_T("   Reported % Complete: ") + itemProps->strRepPercCompl + _T("\n");

				return true;
			}
            else if(itemProps->eAssgnType == TUTaskPercentWorkCompleteReport)
			{
				strTitle =	_T("   Task Name: ") + itemProps->strTask;

				strTip =	_T("   Current % Work Complete: ") + itemProps->strCurrPercWorkCompl + _T("\n")+
							_T("   Reported % Work Complete: ") + itemProps->strRepPercWorkCompl + _T("\n");

				return true;
			}
            else if(itemProps->eAssgnType == TUTaskActualWorkReport)
			{
				strTitle =	_T("   Task Name: ") + itemProps->strTask;

				strTip =	_T("   Current Actual Work: ") + itemProps->strCurrActWork + _T("\n")+
							_T("   Reported Actual Work: ") + itemProps->strRepActWork + _T("\n");

				return true;
			}
            else if(itemProps->eAssgnType == TUTaskActualStartReport)
            {
                strTitle =	_T("   Task Name: ") + itemProps->strTask;

                strTip =	_T("   Current Actual Start: ") + itemProps->strCurrActStart + _T("\n")+
                    _T("   Reported Actual Start: ") + itemProps->strRepActStart + _T("\n");

                return true;
            }
            else if(itemProps->eAssgnType == TUTaskActualFinishReport)
            {
                strTitle =	_T("   Task Name: ") + itemProps->strTask;

                strTip =	_T("   Current Actual Finish: ") + itemProps->strCurrActFinish + _T("\n")+
                    _T("   Reported Actual Finish: ") + itemProps->strRepActFinish + _T("\n");

                return true;
            }
            else if(itemProps->eAssgnType == TUTaskActualDurationReport)
            {
                strTitle =	_T("   Task Name: ") + itemProps->strTask;

                strTip =	_T("   Current Actual Duration: ") + itemProps->strCurrActDuration + _T("\n")+
                    _T("   Reported Actual Duration: ") + itemProps->strRepActDuration + _T("\n");

                return true;
            }
			else
				return false;
		}
		else
			return false;
	}

	return false;
  }
void				CTaskUpdate_ListCtrl::CellHitTest(const CPoint& pt, int& nRow, int& nCol) const
  {
 	 nRow = -1;
 	 nCol = -1;
 
 	 LVHITTESTINFO lvhti = {0};
 	 lvhti.pt = pt;
 	 nRow = ListView_SubItemHitTest(m_hWnd, &lvhti);	// SubItemHitTest is non-const
 	 nCol = lvhti.iSubItem;
 	 if (!(lvhti.flags & LVHT_ONITEMLABEL))
 		 nRow = -1;
  }
INT_PTR				CTaskUpdate_ListCtrl::OnToolHitTest(CPoint point, TOOLINFO * pTI) const
{
	CPoint pt(GetMessagePos());
	ScreenToClient(&pt);

	int nRow, nCol;
	CellHitTest(pt, nRow, nCol);

	//Get the client (area occupied by this control
	RECT rcClient;
	GetClientRect( &rcClient );

	//Fill in the TOOLINFO structure
	pTI->hwnd = m_hWnd;
	pTI->uId = (UINT) (nRow * 1000 + nCol);
	pTI->lpszText = LPSTR_TEXTCALLBACK;	// Send TTN_NEEDTEXT when tooltip should be shown
	pTI->rect = rcClient;

	return pTI->uId; // Must return a unique value for each cell (Marks a new tooltip)
}
 //////////////////////////////////////////////////////////////////////////
void				CTaskUpdate_ListCtrl::OnNMCustomdraw(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVCUSTOMDRAW lpLVCustomDraw = reinterpret_cast<LPNMLVCUSTOMDRAW>(pNMHDR);

	switch(lpLVCustomDraw->nmcd.dwDrawStage)
	{
	case CDDS_ITEMPREPAINT:
	case CDDS_ITEMPREPAINT | CDDS_SUBITEM:
		{
			itemPropsStruct* itemProps = (itemPropsStruct*)lpLVCustomDraw->nmcd.lItemlParam;
			if(itemProps)
			{
				lpLVCustomDraw->clrText = itemProps->clrText; 
				lpLVCustomDraw->clrTextBk = itemProps->clrBackground;
			}
			else 
			{
				lpLVCustomDraw->clrText = CLR_DEFAULT;
				lpLVCustomDraw->clrTextBk = CLR_DEFAULT;
			}
		}
		break;

	default: break;    
	}

	*pResult = 0;
	*pResult |= CDRF_NOTIFYPOSTPAINT;
	*pResult |= CDRF_NOTIFYITEMDRAW;
	*pResult |= CDRF_NOTIFYSUBITEMDRAW;
}
//////////////////////////////////////////////////////////////////////////
    void			CTaskUpdate_ListCtrl::SetItemProps(	
            int iItem, 
            enumTUTaskWorkInsertMode type, 
            DWORD bkClr, 
            DWORD textClr,
            CString	guid,
            CString	task,
            CString	currPercCompl,
            CString	repPercCompl,
            CString	currPercWorkCompl,
            CString	repPercWorkCompl,
            CString currActWork,
            CString repActWork,
			CString currActStart,
			CString repActStart,
			CString currActFinish,
			CString repActFinish,
            CString currActDuration,
            CString repActDuration)
{
	DWORD_PTR ptr = GetItemData(iItem);
	itemPropsStruct* itemProps = (itemPropsStruct*)ptr;
	if(itemProps)
	{
		itemProps->eTaskType = type;
		itemProps->clrBackground = bkClr;
		itemProps->clrText = textClr;
		itemProps->strGUID = guid;
		itemProps->strTask = task;
		itemProps->strCurrPercCompl = currPercCompl;
        itemProps->strRepPercCompl = repPercCompl;
		itemProps->strCurrPercWorkCompl = currPercWorkCompl;
        itemProps->strRepPercWorkCompl = repPercWorkCompl;
        itemProps->strCurrActWork = currActWork;
        itemProps->strRepActWork = repActWork;
        itemProps->strCurrActStart = currActStart;
        itemProps->strRepActStart = repActStart;
        itemProps->strCurrActFinish = currActFinish;
        itemProps->strRepActFinish = repActFinish;
        itemProps->strCurrActDuration = currActDuration;
        itemProps->strRepActDuration = repActDuration;
	}
	else
	{
		itemProps = new itemPropsStruct();
		itemProps->eTaskType = type;
		itemProps->clrBackground = bkClr;
		itemProps->clrText = textClr;
		itemProps->strGUID = guid;
		itemProps->strTask = task;
		itemProps->strCurrPercCompl = currPercCompl;
        itemProps->strRepPercCompl = repPercCompl;
		itemProps->strCurrPercWorkCompl = currPercWorkCompl;
        itemProps->strRepPercWorkCompl = repPercWorkCompl;
        itemProps->strCurrActWork = currActWork;
        itemProps->strRepActWork = repActWork;
        itemProps->strCurrActStart = currActStart;
        itemProps->strRepActStart = repActStart;
        itemProps->strCurrActFinish = currActFinish;
        itemProps->strRepActFinish = repActFinish;
        itemProps->strCurrActDuration = currActDuration;
        itemProps->strRepActDuration = repActDuration;

		itemPropsPtrList.push_back(itemProps);
		SetItemData(iItem, (LPARAM)itemProps);
	}
}

void				CTaskUpdate_ListCtrl::SetItemProps(	
            int iItem,
			enumTUAssgnWorkInsertMode type,
			DWORD bkClr,
			DWORD textClr,
			CString guid,
			CString task,
			CString	resource,
			CString	date,
			CString work,
			CString actWork,
			CString ovtWork,
			CString actOvtWork,
			CString currPerc,
			CString repPerc,
			CString markAsCompl)
{
	DWORD_PTR ptr = GetItemData(iItem);
	itemPropsStruct* itemProps = (itemPropsStruct*)ptr;
	if(itemProps)
	{
		itemProps->eAssgnType = type;
		itemProps->clrBackground = bkClr;
		itemProps->clrText = textClr;
		itemProps->strGUID = guid;
		itemProps->strTask = task;
		itemProps->strResource = resource;
		itemProps->strDate = date;
		itemProps->strWork = work;
		itemProps->strActWork = actWork;
		itemProps->strOvtWork = ovtWork;
		itemProps->strActOvtWork = actOvtWork;
		itemProps->strCurrPerc = currPerc;
		itemProps->strRepPerc = repPerc;
		itemProps->strMarkAsCompl = markAsCompl;
	}
	else
	{
		itemProps = new itemPropsStruct();
		itemProps->eAssgnType = type;
		itemProps->clrBackground = bkClr;
		itemProps->clrText = textClr;
		itemProps->strGUID = guid;
		itemProps->strTask = task;
		itemProps->strResource = resource;
		itemProps->strDate = date;
		itemProps->strWork = work;
		itemProps->strActWork = actWork;
		itemProps->strOvtWork = ovtWork;
		itemProps->strActOvtWork = actOvtWork;
		itemProps->strCurrPerc = currPerc;
		itemProps->strRepPerc = repPerc;
		itemProps->strMarkAsCompl = markAsCompl;

		itemPropsPtrList.push_back(itemProps);
		SetItemData(iItem, (LPARAM)itemProps);
	}
}

itemPropsStruct*	CTaskUpdate_ListCtrl::GetItemProps(int iItem)
{
	DWORD_PTR ptr = GetItemData(iItem);
	itemPropsStruct* itemProps = (itemPropsStruct*)ptr;
	return itemProps;
}
itemPropsStruct*	CTaskUpdate_ListCtrl::GetSelectedItemProps()
{
	POSITION n_pos = GetFirstSelectedItemPosition();
	int sel_pos = n_pos ? GetNextSelectedItem(n_pos) : -1;
	if(sel_pos==-1)
		return NULL;

	return sel_pos>-1 ? GetItemProps(sel_pos) : NULL;
}
int					CTaskUpdate_ListCtrl::GetSelectedItemIndex()
{
	POSITION n_pos = GetFirstSelectedItemPosition();
	return n_pos ? GetNextSelectedItem(n_pos) : -1; 
}
CString				CTaskUpdate_ListCtrl::GetItemGUID(int iItem)
{
	DWORD_PTR ptr = GetItemData(iItem);
	itemPropsStruct* itemProps = (itemPropsStruct*)ptr;
	return itemProps ? itemProps->strGUID : _T("");
}
CString				CTaskUpdate_ListCtrl::GetSelectedItemGUID()
{
	int iSel = GetSelectedItemIndex();
	if (iSel != -1)
	{
		return GetItemGUID(iSel);
	}
	return _T("");
}

void                CTaskUpdate_ListCtrl::SelectItem(CString GUID)
{
    std::vector<int> theList;
    for(int i=0; i<GetItemCount(); ++i)
    {
        CString str = GetItemProps(i)->strGUID;
        if(!str.IsEmpty())
        {
            if(str == GUID)
                SetItemState(i, LVIS_SELECTED, LVIS_SELECTED);
        }
    }
}

void                CTaskUpdate_ListCtrl::OnNMClick(NMHDR *pNMHDR, LRESULT *pResult)
{
    LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);

    CRect rcItem;
    GetSubItemRect(pNMItemActivate->iItem, pNMItemActivate->iSubItem, LVIR_ICON, rcItem);
    GetColumnWidth(pNMItemActivate->iSubItem);

    if ((m_clickState & CLICK_YESNO) && pNMItemActivate->iSubItem == 1)
    {
        CString itemLCText = GetItemText(pNMItemActivate->iItem, pNMItemActivate->iSubItem);
        SetItemText(pNMItemActivate->iItem, pNMItemActivate->iSubItem, itemLCText == _T("Yes") ? _T("No") : _T("Yes"));
    }
    else if ((m_clickState & CLICK_REPORTMODE) && pNMItemActivate->iSubItem == 2)
    {
        CString itemLCText = GetItemText(pNMItemActivate->iItem, pNMItemActivate->iSubItem);
        SetItemText(pNMItemActivate->iItem, pNMItemActivate->iSubItem, itemLCText == _T("Task Reports") ? _T("Assignment Reports") : _T("Task Reports"));
    }

    *pResult = 0;
}

void CTaskUpdate_ListCtrl::SetClickState(int state)
{
    m_clickState = state;
}
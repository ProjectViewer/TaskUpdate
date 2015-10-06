#pragma once
#include "TaskUpdate_Utils.h"
#include <vector>
#define SORT_NUM 1
#define SORT_STR 2
#define SORT_DAT 3

#define FILTER_ACCEPTED			    0x000001
#define FILTER_REJECTED			    0x000002
#define FILTER_PENDING			    0x000004
#define FILTER_OPEN				    0x000008
#define FILTER_TOBE_ACCEPTED		0x000010
#define FILTER_TOBE_REJECTED		0x000020
#define FILTER_ALL				    0x000040

#define COLUMN_CHECKBOX			    0x000001
#define COLUMN_NAME				    0x000002
#define COLUMN_DATE				    0x000004
#define COLUMN_WORK				    0x000008
#define COLUMN_ACT_WORK			    0x000010
#define COLUMN_OVT_WORK			    0x000020
#define COLUMN_ACT_OVT_WORK		    0x000040
#define COLUMN_LAST_PERC		    0x000080
#define COLUMN_REP_PERC			    0x000100
#define COLUMN_CURR_PERC_COMPL		0x000200
#define COLUMN_CURR_PERC_WORK_COMPL	0x000400
#define COLUMN_CURR_ACT_WORK		0x000800
#define COLUMN_REP_PERC_COMPL		0x001000
#define COLUMN_REP_PERC_WORK_COMPL	0x002000
#define COLUMN_REP_ACT_WORK		    0x004000
#define COLUMN_CURR_ACT_START       0x008000
#define COLUMN_REP_ACT_START        0x010000
#define COLUMN_CURR_ACT_FINISH      0x020000
#define COLUMN_REP_ACT_FINISH       0x040000
#define COLUMN_CURR_ACT_DURATION    0x080000
#define COLUMN_REP_ACT_DURATION     0x100000
#define COLUMN_ALL				    0x200000


#define ALL_PROJECTS		_T("")
#define ALL_ASSIGNMENTS		-1
#define ALL_TASKS			-1
#define ALL_RESOURCES		-1
#define ALL_STATUSES        -1

#define CLICK_YESNO                 0x000001
#define CLICK_REPORTMODE            0x000002

struct itemPropsStruct 
{
	enumTUTaskWorkInsertMode				eTaskType;
    enumTUAssgnWorkInsertMode				eAssgnType;

	DWORD			clrBackground;
	DWORD			clrText;

	CString			strGUID;
	CString			strTask;
	CString			strResource;
	CString			strDate;
	CString			strWork;
	CString			strActWork;
	CString			strOvtWork;
	CString			strActOvtWork;
	CString			strCurrPerc;
	CString			strRepPerc;
	CString			strMarkAsCompl;

    CString         strCurrPercCompl;
    CString         strRepPercCompl;
    CString         strCurrPercWorkCompl;
    CString         strRepPercWorkCompl;
    CString         strCurrActWork;
    CString         strRepActWork;
    CString         strCurrActStart;
    CString         strRepActStart;
    CString         strCurrActFinish;
    CString         strRepActFinish;
    CString         strCurrActDuration;
    CString         strRepActDuration;
};

class CTaskUpdate_ListCtrl : public CListCtrl
{
DECLARE_DYNAMIC(CTaskUpdate_ListCtrl)
public:
	CTaskUpdate_ListCtrl(void);
	~CTaskUpdate_ListCtrl(void);

protected:
	DECLARE_MESSAGE_MAP()

	WCHAR*			m_pwchTip;//tool tips
	CHAR*			m_pchTip;//tool tips
	CBitmap			bmp_check;//header check box
	CImageList		ImgHeaders;
    int             m_clickState;
	std::vector<itemPropsStruct*> itemPropsPtrList;//keeps the item color pointers passed to items LPARAM, for deleting them on destructor

	//custom draw
	afx_msg void		OnNMCustomdraw(NMHDR *pNMHDR, LRESULT *pResult);
	//tooltips
	afx_msg BOOL		OnToolNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
	virtual INT_PTR		OnToolHitTest(CPoint point, TOOLINFO * pTI) const;
 	void				CellHitTest(const CPoint& pt, int& nRow, int& nCol) const;
 	bool				GetToolTipText(int nRow, CString& strTitle, CString& strTip);
	virtual void		PreSubclassWindow();

    afx_msg void        OnNMClick(NMHDR *pNMHDR, LRESULT *pResult);
public:
	//sorting
	bool				SortColumn(int columnIndex, bool ascending,int columnType);
	void				SetSortArrow(int colIndex, bool ascending);
	bool				AllItemsAreChecked();
	bool				HasCheckedItem();
	void				CheckAllItems();
	void				UncheckAllItems();
	DWORD   			GetOSMajorVersion();
	void				CreateHeaderCheckBox();
	void				Check_Uncheck_HeaderCheckBox();
	void				Check_Uncheck_AllItems();
	bool				IsEmptyList();
	int					GetCheckedItemCount();
	std::vector<int>	GetCheckedItemsList();  
    void                SetClickState(int state);

	//task update 
        void            FillListCtrl_Task(	
            CTUProjectReport* pProjReport, 
            enumTUUnit appUnit,  
            double hoursPerDay, 
            double hoursPerWeek, 
            int daysPerMonth, 
            int iInsertColumns=COLUMN_ALL, 
            CString strProjectFilter=ALL_PROJECTS,
            int	iTaskFilter=ALL_TASKS,
            int iStatusFilter=ALL_STATUSES);
	    void            FillListCtrl_Assgn(	
            CTUProjectReport* pProjReport, 
            enumTUUnit appUnit,  
            double hoursPerDay, 
            double hoursPerWeek, 
            int daysPerMonth, 
            bool bGrouping = true, 
            int iInsertColumns=COLUMN_ALL, 
            CString strProjectFilter=ALL_PROJECTS,
            int	iTaskFilter=ALL_TASKS,
            int iResourceFilter=ALL_RESOURCES,
            int iStatusFilter=ALL_STATUSES);
	void				ClearListCtrl();

    void				SetItemProps(	
            int iItem, 
            enumTUTaskWorkInsertMode eType, 
            DWORD bkClr, 
            DWORD textClr,
            CString	strGUID,
            CString	strTask,
            CString	strCurrPercCompl,
            CString	strRepPercCompl,
            CString	strCurrPercWorkCompl,
            CString	strRepPercWorkCompl,
            CString strCurrActWork,
			CString strRepActWork,
			CString currActStart,
			CString repActStart,
			CString currActFinish,
			CString repActFinish,
            CString currActDuration,
            CString repActDuration);
	void				SetItemProps(	
            int iItem, 
			enumTUAssgnWorkInsertMode eType, 
			DWORD bkClr, 
			DWORD textClr,
			CString	strGUID,
			CString	strTask,
			CString	strResource,
			CString	strDate,
			CString	strWork,
			CString	strActWork,
			CString	strOvtWork,
			CString	strActOvtWork,
			CString	strCurrPerc,
			CString	strRepPerc,
			CString	strMarkAsCompl);

	itemPropsStruct*	GetItemProps(int iItem);
	itemPropsStruct*	GetSelectedItemProps();
	int					GetSelectedItemIndex();
	CString				GetItemGUID(int iItem);
	CString				GetSelectedItemGUID();
    void                SelectItem(CString GUID);
};


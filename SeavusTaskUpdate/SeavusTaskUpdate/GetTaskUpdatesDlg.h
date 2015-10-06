#pragma once
#include "afxwin.h"
#include "afxcmn.h"
#include "TaskUpdate_ProjectInfo.h"
#include "UpdatesManager.h"
#include "TaskUpdate_ComboBox.h"
#include "TaskUpdate_ListCtrl.h"
#include "TaskUpdate_RadioButton.h"
#include "TaskUpdate_StaticCtrl.h"

// CGetTaskUpdatesDlg dialog

class CGetTaskUpdatesDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CGetTaskUpdatesDlg)


private:

public:
	CGetTaskUpdatesDlg(CWnd* pParent = NULL);   // standard constructor
	CGetTaskUpdatesDlg(CUpdatesManager* updatesManager, CWnd* pParent = NULL);
	virtual ~CGetTaskUpdatesDlg();

// Dialog Data
	enum { IDD = IDD_DLG_TASK_GETUPDATES };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
    CTaskUpdate_ListCtrl m_listCtrl;
    CTUProjectReport m_TUProjRep;
   
    virtual BOOL OnInitDialog();
   
    afx_msg void OnBnClickedOk();
    afx_msg void OnBnClickedButtonAccept();
    afx_msg void OnBnClickedButtonReject();
    afx_msg void OnBnClickedButtonReset();
    afx_msg void OnLvnItemchangedList1(NMHDR *pNMHDR, LRESULT *pResult);
    afx_msg void OnLvnItemchangingList1(NMHDR *pNMHDR, LRESULT *pResult);
    CEdit m_TMNotesEdit;
    CEdit m_PMNotesEdit;
    CButton m_AssgMarkedAsCompleted;
    afx_msg void OnHdnItemclickList1(NMHDR *pNMHDR, LRESULT *pResult);

private:
	CUpdatesManager* m_pUpdatesManager;
	CTaskUpdate_ComboBox m_PrjNameCtrl;
	CTaskUpdate_ComboBox m_TaskNameCtrl;
    CTaskUpdate_ComboBox m_StatusCtrl;

    CTaskUpdate_StaticCtrl m_staticProjectFilter;
    CTaskUpdate_StaticCtrl m_staticTaskFilter;
    CTaskUpdate_StaticCtrl m_staticStatusFilter;

	bool	bAscending;

    CButton m_AcceptButton;
    CButton m_RejectButton;
	CButton m_ResetButton;
    CButton m_buttonOK;	
	CButton m_buttonCancel;
	CStatic m_whiteBkg;
	CStatic m_whiteBkg2;
	CStatic m_whiteBkg3;
	CStatic m_whiteBkg4;
	CStatic m_frame;
	CTaskUpdate_StaticCtrl m_GrBoxComment;
	CEdit	m_editTMComment;
    CEdit	m_editPMComment;
    CTaskUpdate_StaticCtrl m_staticTMComment;
    CTaskUpdate_StaticCtrl m_staticPMComment;

private:
	void	UpdateProjectFilter(CTUProjectReport* pProjectReport);
	void	UpdateTaskFilter(CTUProjectReport* pProjectReport);
    void    UpdateStatusFilter(CTUProjectReport* pProjectReport);
	void	UpdateListCtrl();

    afx_msg void	OnNMClickList1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void	OnCbnCloseupComboProjectName();
	afx_msg void	OnCbnCloseupComboTaskName();
    afx_msg void	OnCbnCloseupComboStatus();
    afx_msg void	OnBnClickedCancel();
	afx_msg void	OnSize(UINT nType, int cx, int cy);
	afx_msg void	OnGetMinMaxInfo(MINMAXINFO* lpMMI);
	afx_msg HBRUSH	OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void	OnEnChangeEditPMComment();
};


#pragma once
#include "TaskUpdate_TreeCtrl.h"
#include "TaskUpdate_ListCtrl.h"
#include "TaskUpdate_ProjectInfo.h"
#include "afxcmn.h"
#include "afxwin.h"

class CProjectProps 
{
public:
    CString name;
    CString path;
    bool collaborative;
    int reportingMode;
    int level;
    int row;
    std::vector<CProjectProps*> subProjects;
};
typedef std::vector<CProjectProps*> PropectPropsList;

class CCollaborationManager;

class CCollDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CCollDlg)

public:
	CCollDlg(CWnd* pParent = NULL);   // standard constructor
	CCollDlg(void* ob, CWnd* pParent = NULL);
	virtual ~CCollDlg();
	
private:

// Dialog Data
	enum { IDD = IDD_DLG_MAKE_COLLABORATIVE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	virtual void OnOK();

private:
    CProjectProps m_projectProps;
	boost::shared_ptr<CCollaborationManager> m_CCollaborationManager;
    CButton m_CollaborativeButtonOK;
    CTaskUpdate_ListCtrl m_list;

    void fillProjectsList();
    void fillSubProjectsList(CProjectProps* prjProps, TUProjectsInfoList* pProjInfoList, int level);
    void fillList();
    void updatePropsFromList();
    void insertSubProjects(PropectPropsList* pProjPropsList);
    
    void writeFiles(PropectPropsList* pProjPropsList);
    void writeFile(CProjectProps* pPrjProp);

    CProjectProps* getPrjProps(PropectPropsList* pProjPropsList, int row);
};

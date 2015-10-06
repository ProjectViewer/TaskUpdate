#include "stdafx.h"
#include "CollDlg.h"
#include "afxdialogex.h"
#include "CollaborationManager.h"
#include "Utils.h"


IMPLEMENT_DYNAMIC(CCollDlg, CDialogEx)

CCollDlg::CCollDlg(CWnd* pParent /*=NULL*/)
: CDialogEx(CCollDlg::IDD, pParent),
m_CCollaborationManager(new CCollaborationManager())
{
}
CCollDlg::CCollDlg(void* obj, CWnd* pParent /*=NULL*/)
	: CDialogEx(CCollDlg::IDD, pParent),
	m_CCollaborationManager(new CCollaborationManager())
{
}

CCollDlg::~CCollDlg()
{
}

void CCollDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Control(pDX, IDOK, m_CollaborativeButtonOK);
    DDX_Control(pDX, IDC_LIST1, m_list);
}


BEGIN_MESSAGE_MAP(CCollDlg, CDialogEx)
END_MESSAGE_MAP()


BOOL CCollDlg::OnInitDialog()
{
    CDialogEx::OnInitDialog();

	EnableToolTips(TRUE);

    fillProjectsList();
    
    fillList();

    return TRUE;
}
void CCollDlg::OnOK()
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
    MSProject::_IProjectDocPtr project = app->GetActiveProject();  
	
	//get active project path
    _bstr_t projectPath = project->FullName;

    //get values from list control
    updatePropsFromList();

    //close the active file
    app->FileClose( MSProject::pjSave);

    //close all projects that need to be modified   
    Utils::closeCollaborativeProjects(app, m_CCollaborationManager->getProjectInfo());

    //update selected files
    writeFile(&m_projectProps);
    writeFiles(&m_projectProps.subProjects);

	//re-open active project
    app->FileOpen(projectPath, 
        vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
        vtMissing,vtMissing,vtMissing, MSProject::pjDoNotOpenPool, vtMissing, vtMissing, vtMissing, vtMissing );

    CDialogEx::OnOK();
}

void CCollDlg::fillProjectsList()
{
    m_projectProps.name = m_CCollaborationManager->getProjectInfo()->GetProjectName();
    m_projectProps.path = m_CCollaborationManager->getProjectInfo()->GetProjectPath();
    m_projectProps.collaborative = m_CCollaborationManager->getProjectInfo()->IsFileCollaborational(); 
    m_projectProps.reportingMode = m_CCollaborationManager->getProjectInfo()->GetProjectReportingMode(); 
    m_projectProps.level = 0;
    fillSubProjectsList(&m_projectProps, m_CCollaborationManager->getProjectInfo()->GetSubProjectsList(), 0);
}
void CCollDlg::fillSubProjectsList(CProjectProps* prjProps, TUProjectsInfoList* pProjInfoList, int level)
{
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
    {       
        CProjectProps* subPrjProps = new CProjectProps();       
        subPrjProps->name = (*itPrjInfo)->GetProjectName();
        subPrjProps->path = (*itPrjInfo)->GetProjectPath();
        subPrjProps->collaborative = (*itPrjInfo)->IsFileCollaborational(); 
        subPrjProps->reportingMode = (*itPrjInfo)->GetProjectReportingMode();
        subPrjProps->level = level + 1;
        prjProps->subProjects.push_back(subPrjProps);

        if ((*itPrjInfo)->HasChilds())
        {
            fillSubProjectsList(subPrjProps, (*itPrjInfo)->GetSubProjectsList(), subPrjProps->level);
        }
    }
}
void CCollDlg::fillList()
{
    CString projectName = m_projectProps.name;
    bool bCollaborative = m_projectProps.collaborative; 
    int reportingMode = m_projectProps.reportingMode; 
    m_projectProps.row = 0;

    m_list.SetClickState(CLICK_YESNO | CLICK_REPORTMODE);
    m_list.ClearListCtrl();

    m_list.InsertColumn(0, _T("Project Name"), LVCFMT_LEFT, 240, -1);
    m_list.InsertColumn(1, _T("Collaborative"), LVCFMT_CENTER, 80, -1);
    m_list.InsertColumn(2, _T("Reporting Mode"), LVCFMT_LEFT, 120, -1);

    m_list.InsertItem(0,_T(""));
    m_list.SetItemText(0, 0, projectName);
    m_list.SetItemText(0, 1, bCollaborative == true ? _T("Yes") : _T("No"));
    m_list.SetItemText(0, 2, reportingMode == 0 ?  _T("Task Reports") : _T("Assignment Reports"));

    insertSubProjects(&m_projectProps.subProjects);

}
void CCollDlg::insertSubProjects(PropectPropsList* pProjPropsList)
{
    PropectPropsList::iterator itPrjProp;
    for (itPrjProp = pProjPropsList->begin(); itPrjProp != pProjPropsList->end(); ++itPrjProp)
    {       
        int row = m_list.GetItemCount();
        (*itPrjProp)->row = row;
        CString projectName = (*itPrjProp)->name;
        bool bCollaborative = (*itPrjProp)->collaborative; 
        int reportingMode = (*itPrjProp)->reportingMode; 
        int level = (*itPrjProp)->level;
        CString indent = _T("");
        while (level)
        {
            indent += _T("        ");
            level--;
        }
      
        m_list.InsertItem(row, _T(""));
        m_list.SetItemText(row, 0, indent + projectName);
        m_list.SetItemText(row, 1, bCollaborative == true ? _T("Yes") : _T("No"));
        m_list.SetItemText(row, 2, reportingMode == 0 ?  _T("Task Reports") : _T("Assignment Reports"));

        if ((*itPrjProp)->subProjects.size())
        {
            insertSubProjects(&(*itPrjProp)->subProjects);
        }
    }
}

void CCollDlg::writeFiles(PropectPropsList* pProjPropsList)
{
    PropectPropsList::iterator itPrjProp;
    for (itPrjProp = pProjPropsList->begin(); itPrjProp != pProjPropsList->end(); ++itPrjProp)
    {
        writeFile(*itPrjProp);
        if ((*itPrjProp)->subProjects.size())
        {
            writeFiles(&(*itPrjProp)->subProjects);
        }
    }
}
void CCollDlg::writeFile(CProjectProps* pPrjProp)
{
    if (CTU_Utils::IsFileWriteble(pPrjProp->path))
    {
        CTUProjectInfo* pProjInfo = m_CCollaborationManager->getProjectInfo(m_CCollaborationManager->GetProjInfoList(), pPrjProp->path);
        if (pProjInfo)
            m_CCollaborationManager->setFileCollaboration(pProjInfo, pPrjProp->collaborative, pPrjProp->reportingMode);
    }
    else
    {
        CString mInfo;
        mInfo.Format(_T("You don’t have sufficient rights on \"%s\" to perform this operation"), pPrjProp->path);
        AfxMessageBox(mInfo);
    }
}

void CCollDlg::updatePropsFromList()
{
    int itemCount = m_list.GetItemCount();
    for (int row = 0; row < itemCount; row++)
    {
        CString collaborative = m_list.GetItemText(row, 1);
        CString reportingMode = m_list.GetItemText(row, 2);

        if (row == 0)
        {
            m_projectProps.collaborative = collaborative == _T("Yes") ? 1 : 0;
            m_projectProps.reportingMode = reportingMode == _T("Task Reports") ? 0 : 1;
        }
        else
        {
            CProjectProps* pPrjProps = getPrjProps(&m_projectProps.subProjects, row);
            if (pPrjProps)
            {
                pPrjProps->collaborative = collaborative == _T("Yes") ? 1 : 0;
                pPrjProps->reportingMode = reportingMode == _T("Task Reports") ? 0 : 1;
            }
        }
    }
}

CProjectProps* CCollDlg::getPrjProps(PropectPropsList* pProjPropsList, int row)
{
    CProjectProps* prjProps = NULL;
    PropectPropsList::iterator itPrjProp;
    for (itPrjProp = pProjPropsList->begin(); itPrjProp != pProjPropsList->end(); ++itPrjProp)
    {
        if ((*itPrjProp)->row == row)
            return *itPrjProp;

        if ((*itPrjProp)->subProjects.size())
        {
            prjProps = getPrjProps(&(*itPrjProp)->subProjects, row);
            if (prjProps)
                return prjProps;
        }
    }

    return prjProps;
}



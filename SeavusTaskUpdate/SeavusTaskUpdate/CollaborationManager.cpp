#include "StdAfx.h"
#include "CollaborationManager.h"
#include "TaskUpdate_ProjectInfo.h"
#include "TaskUpdate_Utils.h"

CCollaborationManager::CCollaborationManager(void)
	: m_pProjectInfo(new CTUProjectInfo())
{
    loadProjectInfo();
    FillProjectInfoList();
}

CCollaborationManager::~CCollaborationManager(void)
{
    if (m_pProjectInfo)
    {
        m_pProjectInfo->Clear();

        delete m_pProjectInfo;
        m_pProjectInfo = NULL;
    }
}

CTUProjectInfo* CCollaborationManager::getProjectInfo() const
{
    return m_pProjectInfo;
}
TUProjectsInfoList*	CCollaborationManager::GetProjInfoList()
{
    return &m_ProjInfoList;
}
void CCollaborationManager::updateProjectInfo(CTUProjectInfo* info)
{
}

void CCollaborationManager::FillProjectInfoList()
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 

    MSProject::_IProjectDocPtr project = app->GetActiveProject();  
    if ((BSTR)project->FullName)
    {
        CTUProjectInfo* pRootPrjInfo = new CTUProjectInfo();

        CString updatesDirPath = _T("");
        CString pathSPV = _T("");
        pRootPrjInfo->SetProjectName((BSTR)project->Name);
        pRootPrjInfo->SetProjectPath((BSTR)project->FullName);
        pRootPrjInfo->SetProjectGUID(CTU_Utils::GetCollIDFromFile((BSTR)project->FullName));
        pRootPrjInfo->SetProjectBaseUID(0);
        CTU_Utils::GetFilePaths((BSTR)project->FullName, pathSPV, updatesDirPath);
        pRootPrjInfo->SetProjectSPVPath(pathSPV);
        pRootPrjInfo->SetProjectReportingMode(CTU_Utils::GetReportingModeFromFile((BSTR)project->FullName));

        m_ProjInfoList.push_back(pRootPrjInfo);

        FillSubProjectsInfoList(app, project, pRootPrjInfo, 0x00400000);
    }
}
void CCollaborationManager::FillSubProjectsInfoList(MSProject::_MSProjectPtr app, MSProject::_IProjectDocPtr project, CTUProjectInfo* pProjectInfo, int parentBaseUID)
{
    MSProject::SubprojectsPtr subProjectsList = project->Subprojects;

    for (int i = 1; i <= subProjectsList->GetCount(); i++)
    {
        if (MSProject::SubprojectPtr subProject = subProjectsList->Item[i])
        {         
            if (MSProject::_IProjectDocPtr subProjectDoc = subProject->GetSourceProject())
            {
                int baseUID = parentBaseUID + (subProject->Index * 0x00400000);

                CTUProjectInfo* pSubProjectInfo = pProjectInfo->AddSubProject(CTU_Utils::GetCollIDFromFile((BSTR)subProjectDoc->FullName), 
                    (BSTR)subProjectDoc->Name, 
                    baseUID,    
                    (BSTR)subProjectDoc->FullName, 
                    0); 

                FillSubProjectsInfoList(app, subProjectDoc, pSubProjectInfo, baseUID);
            }
        }
    }
}


void CCollaborationManager::setFileCollaboration(CTUProjectInfo * pProj, bool isCollaborative, int reportingMode)
{
    pProj->SetFileCollaborational(pProj->GetProjectPath(), isCollaborative, reportingMode);
}

void CCollaborationManager::loadProjectInfo()
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    m_app=app;
    MSProject::_IProjectDocPtr project = app->GetActiveProject();  
    if((LPCTSTR)project->FullName)
    {
        m_pProjectInfo->SetProjectName((LPCTSTR)project->Name);
        m_pProjectInfo->SetProjectPath((LPCTSTR)project->FullName);
        m_pProjectInfo->SetProjectGUID(m_pProjectInfo->GetCollIDFromFile());
        m_pProjectInfo->SetProjectReportingMode(CTU_Utils::GetReportingModeFromFile(m_pProjectInfo->GetProjectPath()));
   
        LoadSubProjects(app, project, m_pProjectInfo);
    }
}

void CCollaborationManager::LoadSubProjects(MSProject::_MSProjectPtr app, 
    MSProject::_IProjectDocPtr project, CTUProjectInfo* projectInfo)
{
    MSProject::SubprojectsPtr subProjectsList = project->Subprojects;

    for(int i = 1; i <= subProjectsList->GetCount(); i++)
    {
        CString path;
        if (MSProject::SubprojectPtr subProject = subProjectsList->Item[i])
        {   
            MSProject::_IProjectDocPtr subProjectDoc = subProject->GetSourceProject();
			CString name = (LPCTSTR)subProjectDoc->Name;
            CString path = (LPCTSTR)subProjectDoc->FullName;
            CString tempCollIDFromFile = projectInfo->GetCollIDFromFile();
            m_pProjectInfo->SetProjectGUID(m_pProjectInfo->GetCollIDFromFile());
            CTUProjectInfo* subProjectInfo = projectInfo->AddSubProject(projectInfo->GetCollIDFromFile(), name, 0, path, 0);  
            LoadSubProjects(app, subProjectDoc, subProjectInfo);
        }
    }
}

int CCollaborationManager::getProjectUpdateMethod()
{
    return CTU_Utils::GetReportingModeFromFile(m_pProjectInfo->GetProjectPath());
}

CTUProjectInfo* CCollaborationManager::getProjectInfo(TUProjectsInfoList* pProjInfoList, CString path)
{
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
    {       
        if ((*itPrjInfo)->GetProjectPath() == path)
            return (*itPrjInfo);

        if ((*itPrjInfo)->HasChilds())
        {
            CTUProjectInfo* pSubProjInfo = getProjectInfo((*itPrjInfo)->GetSubProjectsList(), path);
            if (pSubProjInfo)
                return pSubProjInfo;
        }
    }
    return NULL;
}



#pragma once
#include "TaskUpdate_ProjectInfo.h"
#include <vector>

class CTUProjectInfo;

class CCollaborationManager
{
public:
	CCollaborationManager(void);
	~CCollaborationManager(void);

	CTUProjectInfo* getProjectInfo() const;
	void updateProjectInfo(CTUProjectInfo* info);
	void setFileCollaboration(CTUProjectInfo * pProj , bool isCollaborative, int reportingMode);
    MSProject::_MSProjectPtr getApp(){return m_app;}
    int getProjectUpdateMethod();
    TUProjectsInfoList* GetProjInfoList();
    CTUProjectInfo* getProjectInfo(TUProjectsInfoList* pProjInfoList, CString path);
    
private:
	void loadProjectInfo();
    void FillProjectInfoList();
    void FillSubProjectsInfoList(MSProject::_MSProjectPtr app, MSProject::_IProjectDocPtr project, CTUProjectInfo* pProjectInfo, int parentBaseUID);

    std::vector<CString> mFilePathList;    
    std::vector<CString>::iterator iterFilePathList;
    void LoadSubProjects(MSProject::_MSProjectPtr app, MSProject::_IProjectDocPtr project, CTUProjectInfo* pPrjInfo);
    MSProject::_MSProjectPtr m_app;
	CTUProjectInfo* m_pProjectInfo;
    TUProjectsInfoList			m_ProjInfoList;

};


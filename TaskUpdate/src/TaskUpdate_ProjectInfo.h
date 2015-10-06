#pragma once
#include "TaskUpdate_Utils.h"
#include "DocCustomProperties.h"

class CTUProjectReport;

class CTUProjectInfo
{
public:
    CTUProjectInfo(void);
    ~CTUProjectInfo(void);

private:
    CString m_ProjectGUID;
    CString m_ProjectName;
    int m_ProjectReportingMode;
    long long m_ProjectBaseUID;
    CString m_ProjectPath;
    CString m_ProjecSPVPath;
    CString m_ProjectUpdatesDirPath;
    int m_ProjectRevisionNumber;
    unsigned int m_CustomText1FieldID;
    unsigned int m_CustomText2FieldID;
    unsigned int m_CustomText3FieldID;
    CTUProjectInfo* m_pParentProject;
    std::vector<CTUProjectInfo*> m_SubProjectsList;

    static CTUProjectInfo* pNullPrjInfo;
public:

    int GetProjectReportingMode() const;
    void SetProjectReportingMode(int mode);
    CString GetProjectGUID() const;
    void SetProjectGUID(CString strPrjGUID);
    long long GetProjectBaseUID() const;
    void SetProjectBaseUID(long long iPrjBaseUID);
    CString GetProjectPath() const;
    void SetProjectPath(CString strPrjPath);
    CString GetProjectSPVPath() const;
    void SetProjectSPVPath(CString strPrjSPVPath);
    CString GetProjectUpdatesDirPath() const;
    void SetProjectUpdatesDirPath(CString strPrjUpdDirPath);
    int GetProjectRevisionNumber() const;
    void SetProjectRevisionNumber(int iPrjRevNum);
    CString GetProjectName() const;
    void SetProjectName(CString strPrjname);
    CTUProjectInfo*                     GetParentProject();
    void SetParentProject(CTUProjectInfo* pParentPrj);
    std::vector<CTUProjectInfo*>*       GetSubProjectsList();

    CTUProjectInfo*                     GetRootProject(); //the master project or the active if there are no sub projects
    CTUProjectInfo*                     AddSubProject(CString strPrjGUID, CString strName, int iPrjBaseUID, CString strPrjPath, int iPrjRevNum); //returns pointer to the added subproject object
    void AddSubProject(CTUProjectInfo* pPrjInfo);
    bool RemoveSubProject(CTUProjectInfo* pPrjInfo);                               //returns is sub project(and all it's childs) deleted. for master only sub projects are deleted
    bool RemoveSubProject(CString prjGUID);                               //returns is sub project(and all it's childs) deleted. for master only sub projects are deleted
    void RemoveAllSubProjects();
    void Clear();
    CTUProjectInfo*                     GetProjectInfo(CString prjGUID); //get project with GUID which is sub project to this (or the same project)
    CTUProjectInfo*                     GetProjectInfoByPath(CString prjPath); //get project with path which is sub project to this (or the same project)   
    CTUProjectInfo*                     GetProjectInfoByBaseUID(long long baseUID); //get project with BaseUID which is sub project to this (or the same project)
    bool MPPFileExists();
    bool SPVFileExists();
    bool DeleteSPVFile();
    bool UpdatesDirExists();
    bool HasChilds();
    bool CreateEmptySPVFile();

    CString GetCollIDFromFile();
    bool IsFileCollaborational();
    void SetFileCollaborational(const CString path, bool isCollaborative, int reportingMode);
    void SetMappingValues(std::vector<std::string>& valuesList);
    void SetCustomDocProperty(const std::wstring strKey, std::string strVal);
    bool HasSaveableReport(CTUProjectReport* pProjectReport);                               //for SPV

    void GetAllCustomTextFields(std::vector<unsigned int>& fieldIDs);

    unsigned int GetProjectCustomText1FieldID() const;
    void SetProjectCustomText1FieldID(unsigned int filedID);
    unsigned int GetProjectCustomText2FieldID() const;
    void SetProjectCustomText2FieldID(unsigned int filedID);
    unsigned int GetProjectCustomText3FieldID() const;
    void SetProjectCustomText3FieldID(unsigned int filedID);
};

typedef std::vector<CTUProjectInfo*> TUProjectsInfoList;
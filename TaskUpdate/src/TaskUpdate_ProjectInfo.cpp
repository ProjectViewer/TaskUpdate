#include "TaskUpdate_ProjectInfo.h"

CTUProjectInfo* CTUProjectInfo::pNullPrjInfo = NULL;

CTUProjectInfo::CTUProjectInfo(void)
{
    m_ProjectGUID           = "";
    m_ProjectBaseUID        = 0;
    m_ProjectPath           = "";
    m_ProjecSPVPath         = "";
    m_ProjectRevisionNumber = 0;
    m_ProjectName           = "";
    m_pParentProject        = NULL;
    m_CustomText1FieldID    = 0;
    m_CustomText2FieldID    = 0;
    m_CustomText3FieldID    = 0;
}

CTUProjectInfo::~CTUProjectInfo(void)
{
    Clear();
}

int CTUProjectInfo::GetProjectReportingMode() const
{
    return m_ProjectReportingMode;
}
void CTUProjectInfo::SetProjectReportingMode(int mode)
{
    m_ProjectReportingMode = mode;
}

CString CTUProjectInfo::GetProjectGUID() const
{
    return m_ProjectGUID;
}
void CTUProjectInfo::SetProjectGUID(CString strPrjGUID)
{
    m_ProjectGUID = strPrjGUID;
}
long long CTUProjectInfo::GetProjectBaseUID() const
{
    return m_ProjectBaseUID;
}
void CTUProjectInfo::SetProjectBaseUID(long long iPrjBaseUID)
{
    m_ProjectBaseUID = iPrjBaseUID;
}
CString CTUProjectInfo::GetProjectPath() const
{
    return m_ProjectPath;
}
void CTUProjectInfo::SetProjectPath(CString strPrjPath)
{
    m_ProjectPath = strPrjPath;
}
CString CTUProjectInfo::GetProjectSPVPath() const
{
    return m_ProjecSPVPath;
}
void CTUProjectInfo::SetProjectSPVPath(CString strPrjSPVPath)
{
    m_ProjecSPVPath = strPrjSPVPath;
}
CString CTUProjectInfo::GetProjectUpdatesDirPath() const
{
    return m_ProjectUpdatesDirPath;
}
void CTUProjectInfo::SetProjectUpdatesDirPath(CString strPrjUpdDirPath)
{
    m_ProjectUpdatesDirPath = strPrjUpdDirPath;
}
int CTUProjectInfo::GetProjectRevisionNumber() const
{
    return m_ProjectRevisionNumber;
}
void CTUProjectInfo::SetProjectRevisionNumber(int iPrjRevNum)
{
    m_ProjectRevisionNumber = iPrjRevNum;
}
CString CTUProjectInfo::GetProjectName() const
{
    return m_ProjectName;
}
void CTUProjectInfo::SetProjectName(CString strPrjname)
{
    m_ProjectName = strPrjname;
}
CTUProjectInfo*                     CTUProjectInfo::GetParentProject()
{
    return m_pParentProject;
}
void CTUProjectInfo::SetParentProject(CTUProjectInfo* pParentPrj)
{
    m_pParentProject = pParentPrj;
}
std::vector<CTUProjectInfo*>*       CTUProjectInfo::GetSubProjectsList()
{
    return &m_SubProjectsList;
}

CTUProjectInfo*                     CTUProjectInfo::GetRootProject()
{
    CTUProjectInfo* pParent = GetParentProject();
    while (pParent)
    {
        if (pParent->GetParentProject())
            pParent = pParent->GetParentProject();
        else
            return pParent;
    }
    return this;
}
void CTUProjectInfo::RemoveAllSubProjects()
{
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = m_SubProjectsList.begin(); itPrjInfo != m_SubProjectsList.end();)
    {
        if ((*itPrjInfo)->HasChilds())
        {
            (*itPrjInfo)->RemoveAllSubProjects();
        }
        else
        {
            delete *itPrjInfo;
            *itPrjInfo = NULL;
            itPrjInfo = m_SubProjectsList.erase(itPrjInfo);
        }
    }
}
CTUProjectInfo*                     CTUProjectInfo::AddSubProject(CString strPrjGUID, CString strName, int iPrjBaseUID, CString strPrjPath, int iPrjRevNum)
{
    CString strPath, prjUpdPath;
    CTU_Utils::GetFilePaths(strPrjPath, strPath, prjUpdPath);

    CTUProjectInfo* pSub = new CTUProjectInfo();
    pSub->SetProjectGUID(strPrjGUID);
    pSub->SetProjectName(strName);
    pSub->SetProjectBaseUID(iPrjBaseUID);
    pSub->SetProjectPath(strPrjPath);
    pSub->SetProjectSPVPath(strPath);
    pSub->SetProjectUpdatesDirPath(prjUpdPath);
    pSub->SetProjectRevisionNumber(iPrjRevNum);
    pSub->SetParentProject(this);
    pSub->SetProjectReportingMode(CTU_Utils::GetReportingModeFromFile(strPrjPath));
    pSub->SetProjectCustomText1FieldID(CTU_Utils::GetCustomText1FieldID(strPrjPath));
    pSub->SetProjectCustomText2FieldID(CTU_Utils::GetCustomText2FieldID(strPrjPath));
    pSub->SetProjectCustomText3FieldID(CTU_Utils::GetCustomText3FieldID(strPrjPath));

    m_SubProjectsList.push_back(pSub);

    return pSub;
}
void CTUProjectInfo::AddSubProject(CTUProjectInfo* pPrjInfo)
{
    m_SubProjectsList.push_back(pPrjInfo);
    pPrjInfo->SetParentProject(this);
}
bool CTUProjectInfo::RemoveSubProject(CTUProjectInfo* pPrjInfo)
{
    if (pPrjInfo->GetParentProject() == this)
    {
        pPrjInfo->RemoveAllSubProjects();
        return true;
    }

    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = m_SubProjectsList.begin(); itPrjInfo != m_SubProjectsList.end(); ++itPrjInfo)
    {
        if (*itPrjInfo == pPrjInfo)
        {
            (*itPrjInfo)->RemoveAllSubProjects();

            delete *itPrjInfo;
            *itPrjInfo = NULL;
            m_SubProjectsList.erase(itPrjInfo);
            return true;
        }
    }
    return false;
}
bool CTUProjectInfo::RemoveSubProject(CString prjGUID)
{
    CTUProjectInfo* pPrjInfo = GetProjectInfo(prjGUID);
    if (pPrjInfo)
        return RemoveSubProject(pPrjInfo);
    return false;
}
void CTUProjectInfo::Clear()
{
    RemoveAllSubProjects();
}

CTUProjectInfo*                     CTUProjectInfo::GetProjectInfo(CString prjGUID)
{
    if (m_ProjectGUID == prjGUID)
        return this;

    CTUProjectInfo* pPrjInfo = NULL;
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = m_SubProjectsList.begin(); itPrjInfo != m_SubProjectsList.end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->m_ProjectGUID == prjGUID)
        {
            pPrjInfo =  *itPrjInfo;
            break;
        }
        else
        {
            pPrjInfo = (*itPrjInfo)->GetProjectInfo(prjGUID);
            if (pPrjInfo)
                break;
        }
    }
    return pPrjInfo;
}
CTUProjectInfo*                     CTUProjectInfo::GetProjectInfoByPath(CString prjPath)
{
    if (m_ProjectPath == prjPath)
        return this;

    CTUProjectInfo* pPrjInfo = NULL;
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = m_SubProjectsList.begin(); itPrjInfo != m_SubProjectsList.end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->m_ProjectPath == prjPath)
        {
            pPrjInfo =  *itPrjInfo;
            break;
        }
        else
        {
            pPrjInfo = (*itPrjInfo)->GetProjectInfoByPath(prjPath);
            if (pPrjInfo)
                break;
        }
    }
    return pPrjInfo;
}
CTUProjectInfo*                     CTUProjectInfo::GetProjectInfoByBaseUID(long long baseUID)
{
    if (m_ProjectBaseUID == baseUID)
        return this;

    CTUProjectInfo* pPrjInfo = NULL;
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = m_SubProjectsList.begin(); itPrjInfo != m_SubProjectsList.end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->m_ProjectBaseUID == baseUID)
        {
            pPrjInfo =  *itPrjInfo;
            break;
        }
        else
        {
            pPrjInfo = (*itPrjInfo)->GetProjectInfoByBaseUID(baseUID);
            if (pPrjInfo)
                break;
        }
    }
    return pPrjInfo;
}
bool CTUProjectInfo::MPPFileExists()
{
    return CTU_Utils::FileExists(m_ProjectPath);
}
bool CTUProjectInfo::SPVFileExists()
{
    return CTU_Utils::FileExists(m_ProjecSPVPath);
}
bool CTUProjectInfo::DeleteSPVFile()
{
    return ::DeleteFile(m_ProjecSPVPath) == TRUE;
}
bool CTUProjectInfo::UpdatesDirExists()
{
    return CTU_Utils::DirExists(m_ProjectUpdatesDirPath);
}
bool CTUProjectInfo::HasChilds()
{
    return m_SubProjectsList.size() > 0;
}

bool CTUProjectInfo::CreateEmptySPVFile()
{
    CTUProjectReport theEmptyProjectReports;
    theEmptyProjectReports.SetProjectGUID(m_ProjectGUID);
    theEmptyProjectReports.SetProjectName(m_ProjectName);

    return CTU_Utils::SaveXML(m_ProjecSPVPath, &theEmptyProjectReports);
}

CString CTUProjectInfo::GetCollIDFromFile()
{
    CDocCustomProperties docprop;
    docprop.LoadCustomProperties(m_ProjectPath);
    return docprop.GetCustomProperertyValueFor(_T("Collaboration ID"));
}

bool CTUProjectInfo::IsFileCollaborational()
{
    CDocCustomProperties docprop;
    docprop.LoadCustomProperties(m_ProjectPath);
    CString strGUID =  docprop.GetCustomProperertyValueFor(_T("Collaboration ID"));
    CString strActive = docprop.GetCustomProperertyValueFor(_T("IsTUActive"));

    return (CTU_Utils::IsValidGUID(strGUID) && strActive == _T("1"));
}

void CTUProjectInfo::SetFileCollaborational(const CString path, bool isCollaborative, int reportingMode)
{
    CDocCustomProperties docprop;
    CString strOldGUID = GetCollIDFromFile();
    if (strOldGUID.IsEmpty() || !CTU_Utils::IsValidGUID(strOldGUID))
        strOldGUID = CTU_Utils::GenerateGUID();
    docprop.SetCustomDocProperties(path, isCollaborative, strOldGUID, reportingMode);
}
void CTUProjectInfo::SetMappingValues(std::vector<std::string>& valuesList)
{
    CDocCustomProperties docprop;
    docprop.SetCustomDocPropertiesMapping(GetProjectPath(), valuesList);
}
void CTUProjectInfo::SetCustomDocProperty(const std::wstring strKey, std::string strVal)
{
    CDocCustomProperties docprop;
    docprop.SetCustomDocProperty(GetProjectPath(), strKey, strVal);
}
bool CTUProjectInfo::HasSaveableReport(CTUProjectReport* pProjectReport)
{
    TUTaskReportsList::iterator itTaskReport;

    for (itTaskReport = pProjectReport->GetTaskReportsList()->begin();
        itTaskReport != pProjectReport->GetTaskReportsList()->end();
        itTaskReport++)
    {
        if ((*itTaskReport)->GetTaskBaseUID() == m_ProjectBaseUID)
        {
            return (*itTaskReport)->IsSaveable();
        }
    }
    return false;

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = pProjectReport->GetAssignmentReportsList()->begin();
         itAssgnReport != pProjectReport->GetAssignmentReportsList()->end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->GetAssignmentBaseUID() == m_ProjectBaseUID)
        {
            if ((*itAssgnReport)->HasSaveableReport())
                return true;
        }
    }
    return false;
}

void CTUProjectInfo::GetAllCustomTextFields(std::vector<unsigned int>& fieldIDs)
{
    CTUProjectInfo* pPrjInfo = NULL;
    TUProjectsInfoList::iterator itPrjInfo;

    fieldIDs.push_back(m_CustomText1FieldID);
    fieldIDs.push_back(m_CustomText2FieldID);
    fieldIDs.push_back(m_CustomText3FieldID);

    for (itPrjInfo = m_SubProjectsList.begin(); itPrjInfo != m_SubProjectsList.end(); ++itPrjInfo)
    {
        (*itPrjInfo)->GetAllCustomTextFields(fieldIDs);
    }
}

unsigned int CTUProjectInfo::GetProjectCustomText1FieldID() const
{
    return m_CustomText1FieldID;
}
void CTUProjectInfo::SetProjectCustomText1FieldID(unsigned int filedID)
{
    m_CustomText1FieldID = filedID;
}
unsigned int CTUProjectInfo::GetProjectCustomText2FieldID() const
{
    return m_CustomText2FieldID;
}
void CTUProjectInfo::SetProjectCustomText2FieldID(unsigned int filedID)
{
    m_CustomText2FieldID = filedID;
}
unsigned int CTUProjectInfo::GetProjectCustomText3FieldID() const
{
    return m_CustomText3FieldID;
}
void CTUProjectInfo::SetProjectCustomText3FieldID(unsigned int filedID)
{
    m_CustomText3FieldID = filedID;
}
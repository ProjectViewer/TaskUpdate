#include "TaskUpdate_ProjectReport.h"

CTUTaskReport* CTUProjectReport::pNullTaskReport = NULL;
CTUAssignmentReport* CTUProjectReport::pNullAssgnReport = NULL;
CTUDateReport* CTUProjectReport::pNullDateReport = NULL;

CTUProjectReport::CTUProjectReport(void)
{
}

CTUProjectReport::~CTUProjectReport(void)
{
    Clear();
}
void                        CTUProjectReport::Clear()
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin(); itTaskReport != m_TaskReportsList.end();)
    {
        delete *itTaskReport;
        *itTaskReport = NULL;
        itTaskReport = m_TaskReportsList.erase(itTaskReport);
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin(); itAssgnReport != m_AssignmentReportsList.end();)
    {
        (*itAssgnReport)->Clear();
        delete *itAssgnReport;
        *itAssgnReport = NULL;
        itAssgnReport = m_AssignmentReportsList.erase(itAssgnReport);
    }
}
CString                     CTUProjectReport::GetProjectGUID() const
{
    return m_ProjectGUID;
}
void                        CTUProjectReport::SetProjectGUID(CString strPrjGUID)
{
    m_ProjectGUID = strPrjGUID;
}
CString                     CTUProjectReport::GetProjectName() const
{
    return m_ProjectName;
}
void                        CTUProjectReport::SetProjectName(CString strPrjName)
{
    m_ProjectName  = strPrjName;
}
enumTUProjectReportingMode CTUProjectReport::GetWorkInsertMode() const
{
    return m_workInsertMode;
}
void                        CTUProjectReport::SetWorkInsertMode (enumTUProjectReportingMode mode)
{
    m_workInsertMode = mode;
}

TUTaskReportsList*          CTUProjectReport::GetTaskReportsList()
{
    return &m_TaskReportsList;
}
CTUTaskReport*              CTUProjectReport::GetTaskReport(long long iTaskUID)
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->GetTaskUID() == iTaskUID)
        {
            return *itTaskReport;
        }
    }
    return pNullTaskReport;
}
CTUTaskReport*              CTUProjectReport::GetTaskReport(CString GUID)
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->GetReportGUID() == GUID)
        {
            return *itTaskReport;
        }
    }
    return pNullTaskReport;
}
CTUTaskReport*              CTUProjectReport::GetTaskReportByUINT_PTR(UINT_PTR uint_ptr)
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->GetUINT_PTR() == uint_ptr)
        {
            return *itTaskReport;
        }
    }
    return pNullTaskReport;
}
void                        CTUProjectReport::AddTaskReport(CTUTaskReport* pTaskReport)
{
    m_TaskReportsList.push_back(pTaskReport);
}
void                        CTUProjectReport::RemoveTaskReport(CTUTaskReport* pTaskReport)
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->GetTaskUID() == pTaskReport->GetTaskUID())
        {
            m_TaskReportsList.erase(itTaskReport);
            delete pTaskReport;
            pTaskReport = NULL;
            return;
        }
    }
}


TUAssignmentReportsList*    CTUProjectReport::GetAssignmentReportsList()
{
    return &m_AssignmentReportsList;
}
void                        CTUProjectReport::GetAssignmentReportsListForTask(long long iTaskUID, TUAssignmentReportsList& theList)
{
    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->GetTaskUID() == iTaskUID)
            theList.push_back(*itAssgnReport);
    }
}
void                        CTUProjectReport::GetAssignmentReportsListForResource(long long iResUID, TUAssignmentReportsList& theList)
{
    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->GetResourceUID() == iResUID)
            theList.push_back(*itAssgnReport);
    }
}
CTUAssignmentReport*        CTUProjectReport::GetAssgnReport(long long assgnUID)
{
    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->GetAssignmentUID() == assgnUID)
        {
            return *itAssgnReport;
        }
    }
    return pNullAssgnReport;
}
CTUAssignmentReport*        CTUProjectReport::GetAssgnReport(long long iTaskUID, long long iResUID)
{
    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->GetTaskUID() == iTaskUID && (*itAssgnReport)->GetResourceUID() == iResUID)
        {
            return *itAssgnReport;
        }
    }
    return pNullAssgnReport;
}
void                        CTUProjectReport::AddAssgnReport(CTUAssignmentReport* pAssgnReport)
{
    m_AssignmentReportsList.push_back(pAssgnReport);
}
void                        CTUProjectReport::RemoveAssgnReport(CTUAssignmentReport* pAssgnReport)
{
    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->GetAssignmentUID() == pAssgnReport->GetAssignmentUID())
        {
            m_AssignmentReportsList.erase(itAssgnReport);
            delete pAssgnReport;
            pAssgnReport = NULL;
            return;
        }
    }
}

CTUDateReport*              CTUProjectReport::GetPercentReport(long long assgnUID)
{
    CTUAssignmentReport* pAssgnReport = GetAssgnReport(assgnUID);
    if (pAssgnReport)
    {
        return pAssgnReport->GetPercentReport();
    }
    return pNullDateReport;
}
CTUDateReport*              CTUProjectReport::GetPercentReport(long long assgnUID, bool bActive)
{
    CTUAssignmentReport* pAssgnReport = GetAssgnReport(assgnUID);
    if (pAssgnReport)
    {
        return pAssgnReport->GetPercentReport(bActive);
    }
    return pNullDateReport;
}
CTUDateReport*              CTUProjectReport::GetTimeReport(long long assgnUID, CTime reportTime)
{
    CTUAssignmentReport* pAssgnReport = GetAssgnReport(assgnUID);
    if (pAssgnReport)
    {
        return pAssgnReport->GetTimeReport(reportTime);
    }
    return pNullDateReport;
}
CTUDateReport*              CTUProjectReport::GetTimeReport(long long assgnUID, CTime reportTime, bool bActive)
{
    CTUAssignmentReport* pAssgnReport = GetAssgnReport(assgnUID);
    if (pAssgnReport)
    {
        return pAssgnReport->GetTimeReport(reportTime, bActive);
    }
    return pNullDateReport;
}
CTUDateReport*              CTUProjectReport::GetDateReport(CString GUID)
{
    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        CTUDateReport* pReport = (*itAssgnReport)->GetDateReport(GUID);
        if (pReport)
        {
            return pReport;
        }
    }
    return pNullDateReport;
}
CTUDateReport*              CTUProjectReport::GetDateReportByUINT_PTR(UINT_PTR uint_ptr)
{
    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        CTUDateReport* pReport = (*itAssgnReport)->GetReportByUINT_PTR(uint_ptr);
        if (pReport)
        {
            return pReport;
        }
    }
    return pNullDateReport;
}
bool                        CTUProjectReport::HasSaveableReport()
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->IsSaveable())
        {
            return true;
        }
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->HasSaveableReport())
        {
            return true;
        }
    }
    return false;
}
bool                        CTUProjectReport::HasAcceptedReport()
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->GetStatus() == TUTaskAccepted)
        {
            return true;
        }
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->HasAcceptedReport())
        {
            return true;
        }
    }
    return false;
}
bool                        CTUProjectReport::HasRejectedReport()
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->GetStatus() == TUTaskRejected)
        {
            return true;
        }
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->HasRejectedReport())
        {
            return true;
        }
    }
    return false;
}
bool                        CTUProjectReport::HasEditedReport()
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->IsEdited())
        {
            return true;
        }
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->HasEditedReport())
        {
            return true;
        }
    }
    return false;
}
bool                        CTUProjectReport::HasTMNote()
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->HasTMNote())
        {
            return true;
        }
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
        itAssgnReport != m_AssignmentReportsList.end();
        itAssgnReport++)
    {
        if ((*itAssgnReport)->HasTMNote())
        {
            return true;
        }
    }
    return false;
}

void                        CTUProjectReport::SetReportsCreationDate()
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        (*itTaskReport)->SetCreationDate(CTime::GetCurrentTime());
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    TUDateReportsList::iterator itDateReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        for (itDateReport = (*itAssgnReport)->GetDateReportsList()->begin();
             itDateReport != (*itAssgnReport)->GetDateReportsList()->end();
             itDateReport++)
        {
            (*itDateReport)->SetCreationDate(CTime::GetCurrentTime());
        }
    }
}
void                        CTUProjectReport::MarkEditedReportsAsSaved()
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->IsEdited())
        {
            (*itTaskReport)->SetPercentCompleteEdited(false);
            (*itTaskReport)->SetPercentWorkCompleteEdited(false);
            (*itTaskReport)->SetActualWorkEdited(false);
			(*itTaskReport)->SetActualStartEdited(false);
			(*itTaskReport)->SetActualFinishEdited(false);
            (*itTaskReport)->SetActualDurationEdited(false);
            (*itTaskReport)->SetNotesEdited(false);
            (*itTaskReport)->SetCustomText1Edited(false);
            (*itTaskReport)->SetCustomText2Edited(false);
            (*itTaskReport)->SetCustomText3Edited(false);
            (*itTaskReport)->SetSaved();
            (*itTaskReport)->SetCustomTextSaved();
            (*itTaskReport)->SetNotesSaved();
        }
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        (*itAssgnReport)->MarkEditedReportsAsSaved();
    }
}

void                        CTUProjectReport::EmptyInactiveReports()
{
    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        (*itAssgnReport)->EmptyInactiveReports();
    }
}

void                        CTUProjectReport::ClearTaskReportRects()
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        (*itTaskReport)->ClearReportRects();
    }
}

void                        CTUProjectReport::ClearReportRects(enumTUAssgnWorkInsertMode eWorkInsertMode)
{
    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        (*itAssgnReport)->ClearReportRects(eWorkInsertMode);
    }
}

void                        CTUProjectReport::AddBaseUID(long long iBaseUID)
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        (*itTaskReport)->SetTaskBaseUID(iBaseUID);
        (*itTaskReport)->SetTaskUID((*itTaskReport)->GetTaskUID() + iBaseUID);
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        (*itAssgnReport)->SetAssignmentBaseUID(iBaseUID);
        (*itAssgnReport)->SetAssignmentUID((*itAssgnReport)->GetAssignmentUID() + iBaseUID);
        (*itAssgnReport)->SetResourceUID((*itAssgnReport)->GetResourceUID() + iBaseUID);
        (*itAssgnReport)->SetTaskUID((*itAssgnReport)->GetTaskUID() + iBaseUID);
    }
}
void                        CTUProjectReport::RemoveBaseUID(long long iBaseUID)
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        (*itTaskReport)->SetTaskBaseUID(iBaseUID);
        (*itTaskReport)->SetTaskUID((*itTaskReport)->GetTaskUID() - iBaseUID);
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        (*itAssgnReport)->SetAssignmentBaseUID(iBaseUID);
        (*itAssgnReport)->SetAssignmentUID((*itAssgnReport)->GetAssignmentUID() - iBaseUID);
        (*itAssgnReport)->SetResourceUID((*itAssgnReport)->GetResourceUID() - iBaseUID);
        (*itAssgnReport)->SetTaskUID((*itAssgnReport)->GetTaskUID() - iBaseUID);
    }
}

CTUAssignmentReport*        CTUProjectReport::GetDatesAssignmentReport(CTUDateReport* pDateReport)
{
    TUAssignmentReportsList::iterator itAssgnReport;
    TUDateReportsList::iterator itDateReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        CTUDateReport* pReport = (*itAssgnReport)->GetDateReport(pDateReport->GetReportGUID());
        if (pReport)
            return *itAssgnReport;
    }
    return pNullAssgnReport;
}

void                        CTUProjectReport::SetReportsType()
{
    TUAssignmentReportsList::iterator itAssgnReport;
    TUDateReportsList::iterator itDateReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        (*itAssgnReport)->SetReportsType();
    }
}

void                        CTUProjectReport::SetAllReportsStatus(int eStatus)
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        (*itTaskReport)->SetStatus((enumTUTaskReportsStatus)eStatus);
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    TUDateReportsList::iterator itDateReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        (*itAssgnReport)->SetAllReportsStatus((enumTUReportsStatus)eStatus);
        (*itAssgnReport)->SetMarkAsCompleteType(TU_MAC_None);
    }
}

void                        CTUProjectReport::ChangeAllReportsStatus(int eFromStatus, int eToStatus)
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->GetStatus() == (enumTUTaskReportsStatus)eFromStatus || TUTaskAll == (enumTUTaskReportsStatus)eFromStatus)
        {
            (*itTaskReport)->SetStatus((enumTUTaskReportsStatus)eToStatus);
        }        
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    TUDateReportsList::iterator itDateReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        (*itAssgnReport)->ChangeAllReportsStatus((enumTUReportsStatus)eFromStatus, (enumTUReportsStatus)eToStatus);
    }
}

void                        CTUProjectReport::RemoveReportsWithStatus(int eStatus)
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();)
    {
        if ((*itTaskReport)->GetStatus() == (enumTUTaskReportsStatus)eStatus)
        {
            delete *itTaskReport;
            *itTaskReport = NULL;
            itTaskReport = m_TaskReportsList.erase(itTaskReport);
        }
        else
            itTaskReport++;
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
        itAssgnReport != m_AssignmentReportsList.end();
        itAssgnReport++)
    {
        (*itAssgnReport)->RemoveReportsWithStatus((enumTUReportsStatus)eStatus);
    }
}

void                        CTUProjectReport::RemoveReportsWithStatus(CString strProjGUID, int eStatus)
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->GetProjectGUID() == strProjGUID && 
            (*itTaskReport)->GetStatus() == (enumTUTaskReportsStatus)eStatus)
        {
            delete *itTaskReport;
            *itTaskReport = NULL;
            itTaskReport = m_TaskReportsList.erase(itTaskReport);
        }
        else
            itTaskReport++;
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
         itAssgnReport != m_AssignmentReportsList.end();
         itAssgnReport++)
    {
        if ((*itAssgnReport)->GetProjectGUID() == strProjGUID)
        {
            (*itAssgnReport)->RemoveReportsWithStatus((enumTUReportsStatus)eStatus);
        }
    }
}

void                        CTUProjectReport::UpdateCurrentValues()
{
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_TaskReportsList.begin();
        itTaskReport != m_TaskReportsList.end();
        itTaskReport++)
    {
        if ((*itTaskReport)->GetStatus() == TUTaskAccepted)
        {
            if ((*itTaskReport)->IsPercentCompleteActive())
            {
                (*itTaskReport)->SetCurrentPercentComplete((*itTaskReport)->GetReportedPercentComplete());
            }
            else if ((*itTaskReport)->IsPercentWorkCompleteActive())
            {
                (*itTaskReport)->SetCurrentPercentWorkComplete((*itTaskReport)->GetReportedPercentWorkComplete());
            }
            else if ((*itTaskReport)->IsActualWorkActive())
            {
                (*itTaskReport)->SetCurrentActualWork((*itTaskReport)->GetReportedActualWork());
            }
			else if ((*itTaskReport)->IsActualStartActive())
			{
				(*itTaskReport)->SetCurrentActualStart((*itTaskReport)->GetReportedActualStart());
			}
			else if ((*itTaskReport)->IsActualFinishActive())
			{
				(*itTaskReport)->SetCurrentActualFinish((*itTaskReport)->GetReportedActualFinish());
			}
            else if ((*itTaskReport)->IsActualDurationActive())
            {
                (*itTaskReport)->SetCurrentActualDuration((*itTaskReport)->GetReportedActualDuration());
            }
        }
    }

    TUAssignmentReportsList::iterator itAssgnReport;
    for (itAssgnReport = m_AssignmentReportsList.begin();
        itAssgnReport != m_AssignmentReportsList.end();
        itAssgnReport++)
    {
        (*itAssgnReport)->UpdateCurrentValues();
    }
}

bool CTUProjectReport::IsTaskUpdateCustomTextField(long long taskUID, int fieldID)
{
    bool returnValue = false;

    if (fieldID != 0)
    {
        TUTaskReportsList::iterator itTaskReport;
        for (itTaskReport = m_TaskReportsList.begin();
            itTaskReport != m_TaskReportsList.end();
            itTaskReport++)
        {
            if ((*itTaskReport)->GetTaskUID() == taskUID)
            {
                if ((*itTaskReport)->GetCustomText1FieldID() == fieldID ||
                    (*itTaskReport)->GetCustomText2FieldID() == fieldID ||
                    (*itTaskReport)->GetCustomText3FieldID() == fieldID)
                {
                    returnValue = true;
                    break;
                }
            }
        }
    }

    return returnValue;
}

bool CTUProjectReport::HasReportWithWorkInsertMode(enumTUTaskWorkInsertMode eWorkInsertMode)
{
	bool returnValue = false;
	TUTaskReportsList::iterator itTaskReport;
	for (itTaskReport = m_TaskReportsList.begin();
		itTaskReport != m_TaskReportsList.end();
		itTaskReport++)
	{
		if ((*itTaskReport)->GetWorkInsertMode() == eWorkInsertMode &&
			(*itTaskReport)->GetStatus() != TUTaskAccepted)
		{
			returnValue = true;
			break;
		}
	}

	return returnValue;
}
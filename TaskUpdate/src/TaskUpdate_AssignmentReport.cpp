#include "TaskUpdate_AssignmentReport.h"

CTUDateReport* CTUAssignmentReport::pNullDateReport = NULL;

CTUAssignmentReport::CTUAssignmentReport(void)
:m_MarkAsCompleteType(TU_MAC_None)
{
}

CTUAssignmentReport::~CTUAssignmentReport(void)
{
    Clear();
}

void CTUAssignmentReport::Clear()
{
    RemoveDateReports();
}
long long CTUAssignmentReport::GetAssignmentUID() const
{
    return m_AssignmentUID;
}
void CTUAssignmentReport::SetAssignmentUID(long long iAssgnUID)
{
    m_AssignmentUID = iAssgnUID;
}
CString CTUAssignmentReport::GetProjectName() const
{
    return m_ProjectName;
}
void CTUAssignmentReport::SetProjectName(CString prjName)
{
    m_ProjectName = prjName;
}
CString CTUAssignmentReport::GetProjectGUID() const
{
    return m_ProjectGUID;
}
void CTUAssignmentReport::SetProjectGUID(CString prjGUID)
{
    m_ProjectGUID = prjGUID;
}
long long CTUAssignmentReport::GetAssignmentBaseUID() const
{
    return m_AssignmentBaseUID;
}
void CTUAssignmentReport::SetAssignmentBaseUID(long long iBaseUID)
{
    m_AssignmentBaseUID = iBaseUID;
}
long long CTUAssignmentReport::GetResourceUID() const
{
    return m_ResourceUID;
}
void CTUAssignmentReport::SetResourceUID(long long iResUID)
{
    m_ResourceUID = iResUID;
}
CString CTUAssignmentReport::GetResourceName() const
{
    return m_ResourceName;
}
void CTUAssignmentReport::SetResourceName(CString strResName)
{
    m_ResourceName = strResName;
}
long long CTUAssignmentReport::GetTaskUID() const
{
    return m_TaskUID;
}
void CTUAssignmentReport::SetTaskUID(long long iTaskUID)
{
    m_TaskUID = iTaskUID;
}
CString CTUAssignmentReport::GetTaskName() const
{
    return m_TaskName;
}
void CTUAssignmentReport::SetTaskName(CString strTaskName)
{
    m_TaskName = strTaskName;
}
enumTUAssgnWorkInsertMode CTUAssignmentReport::GetWorkInsertMode() const
{
    return m_WorkInsertMode;
}
void CTUAssignmentReport::SetWorkInsertMode(enumTUAssgnWorkInsertMode eWorkInsertMode)
{
    m_WorkInsertMode = eWorkInsertMode;
}
bool CTUAssignmentReport::IsMarkedAsCompleted() const
{
    return m_MarkAsCompleted;
}
void CTUAssignmentReport::SetMarkedAsCompleted(bool bMarkAsCompleted)
{
    m_MarkAsCompleted = bMarkAsCompleted;
}
enumMarkAsCompleteType CTUAssignmentReport::GetMarkAsCompleteType() const
{
    return m_MarkAsCompleteType;
}
void CTUAssignmentReport::SetMarkAsCompleteType(enumMarkAsCompleteType eMarkAsCompleteType)
{
    m_MarkAsCompleteType = eMarkAsCompleteType;
}

TUDateReportsList*          CTUAssignmentReport::GetDateReportsList()
{
    return &m_DateReportsList;
}

CTUDateReport*              CTUAssignmentReport::GetPercentReport()
{
    TUDateReportsList::iterator itDateReport;

    for (itDateReport = m_DateReportsList.begin();
        itDateReport != m_DateReportsList.end();
        itDateReport++)
    {
        if ((*itDateReport)->GetType() == TUAssgnPercentReports)
        {
            return *itDateReport;
        }
    }

    return pNullDateReport;
}

CTUDateReport*              CTUAssignmentReport::GetPercentReport(bool bActive)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if (bActive)
        {
            if (!(*itDateReport)->IsActive())
                continue;
        }
        if ((*itDateReport)->GetType() == TUAssgnPercentReports)
        {
            return *itDateReport;
        }
    }
    return pNullDateReport;
}

CTUDateReport*              CTUAssignmentReport::GetTimeReport(CTime reportTime)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->GetType() == TUAssgnTimeReports && (*itDateReport)->GetReportDate() == reportTime)
        {
            return *itDateReport;
        }
    }
    return pNullDateReport;
}
CTUDateReport*              CTUAssignmentReport::GetTimeReport(CTime reportTime, bool bActive)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if (bActive)
        {
            if (!(*itDateReport)->IsActive())
                continue;
        }
        if ((*itDateReport)->GetType() == TUAssgnTimeReports && (*itDateReport)->GetReportDate() == reportTime)
        {
            return *itDateReport;
        }
    }
    return pNullDateReport;
}
CTUDateReport*              CTUAssignmentReport::GetDateReport(CString GUID)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->GetReportGUID() == GUID)
            return *itDateReport;
    }
    return pNullDateReport;
}
CTUDateReport*              CTUAssignmentReport::GetReportByUINT_PTR(UINT_PTR uint_ptr)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->GetUINT_PTR() == uint_ptr)
            return *itDateReport;
    }
    return pNullDateReport;
}
void CTUAssignmentReport::AddDateReport(CTUDateReport* pDateReport)
{
    m_DateReportsList.push_back(pDateReport);
}
void CTUAssignmentReport::RemoveDateReport(CTUDateReport* pDateReport)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->GetReportGUID() == pDateReport->GetReportGUID())
        {
            m_DateReportsList.erase(itDateReport);
            delete pDateReport;
            pDateReport = NULL;
            return;
        }
    }
}
void CTUAssignmentReport::RemoveDateReports()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin(); itDateReport != m_DateReportsList.end();)
    {
        delete *itDateReport;
        *itDateReport = NULL;
        itDateReport = m_DateReportsList.erase(itDateReport);
    }
}
bool CTUAssignmentReport::HasReports() const
{
    return m_DateReportsList.size() > 0;
}
bool CTUAssignmentReport::HasSaveableReport()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->IsEdited() || (*itDateReport)->IsSaved() || (*itDateReport)->GetStatus() == TURejected)
            return true;
    }
    return false;
}

bool CTUAssignmentReport::HasRejectedReport()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->GetStatus() == TURejected)
            return true;
    }
    return false;
}

bool CTUAssignmentReport::HasAcceptedReport()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->GetStatus() == TUAccepted)
            return true;
    }
    return false;
}

bool CTUAssignmentReport::HasSavedReport()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->IsSaved())
            return true;
    }
    return false;
}

bool CTUAssignmentReport::HasEditedReport()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->IsEdited())
            return true;
    }
    return false;
}
bool CTUAssignmentReport::IsDisabled()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
        itDateReport != m_DateReportsList.end();
        itDateReport++)
    {
        if ((*itDateReport)->IsDisabled())
            return true;
    }
    return false;
}
bool CTUAssignmentReport::HasTMNote()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
        itDateReport != m_DateReportsList.end();
        itDateReport++)
    {
        if ((*itDateReport)->HasTMNote())
            return true;
    }
    return false;
}
void CTUAssignmentReport::EmptyInactiveReports()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if (!(*itDateReport)->IsActive())
        {
            (*itDateReport)->EmtpyReport();
        }
    }
}

void CTUAssignmentReport::EmptyReports()
{
    SetWorkInsertMode(TUAssgnUndefiniedReport);
    SetMarkedAsCompleted(false);
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        (*itDateReport)->EmtpyReport();
    }
}

void CTUAssignmentReport::ClearReports()
{
    SetWorkInsertMode(TUAssgnUndefiniedReport);
    SetMarkedAsCompleted(false);
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        (*itDateReport)->ClearReport();
    }
}

void CTUAssignmentReport::ClearReportRects(enumTUAssgnWorkInsertMode eWorkInsertMode)
{
    CRect nullRect(0, 0, 0, 0);
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->GetType() == eWorkInsertMode)
        {
            (*itDateReport)->SetPercRect(nullRect);
            (*itDateReport)->SetWorkRect(nullRect);
            (*itDateReport)->SetOvtWorkRect(nullRect);
        }
    }
}

bool CTUAssignmentReport::HasPreviousEmptyReports(CTUDateReport* pReport)
{
    TUDateReportsList::iterator itDateReport;

    CTUDateReport* pFirstReport = GetFirstFlexEmptyReport();
    if (pFirstReport && pReport->GetReportDate() > pFirstReport->GetReportDate())
    {
        for (itDateReport = m_DateReportsList.begin();
             itDateReport != m_DateReportsList.end();
             ++itDateReport)
        {
            if ((*itDateReport)->GetType() == TUAssgnTimeReports)
            {
                if ((*itDateReport)->GetReportDate() >= pFirstReport->GetReportDate() &&
                    (*itDateReport)->GetReportDate() < pReport->GetReportDate())
                {
                    if (!(*itDateReport)->IsEdited() && !(*itDateReport)->IsSaved() && !(*itDateReport)->IsVirtual())
                    {
                        return true;
                    }
                }
            }
        }
    }
    return false;
}

bool CTUAssignmentReport::HasPreviousDateReports(CTUDateReport* pReport)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         ++itDateReport)
    {
        if ((*itDateReport)->GetType() == TUAssgnTimeReports)
        {
            if ((*itDateReport)->GetReportDate() < pReport->GetReportDate())
            {
                return true;
            }
        }
    }

    return false;
}

bool CTUAssignmentReport::HasNextDateReports(CTUDateReport* pReport)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         ++itDateReport)
    {
        if ((*itDateReport)->GetType() == TUAssgnTimeReports)
        {
            if ((*itDateReport)->GetReportDate() > pReport->GetReportDate())
            {
                return true;
            }
        }
    }

    return false;
}

CTUDateReport*              CTUAssignmentReport::GetFirstFlexEmptyReport()
{
    TUDateReportsList::iterator itDateReport;

    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->GetType() == TUAssgnTimeReports)
        {
            if (!(*itDateReport)->IsVirtual())
            {
                return *itDateReport;
            }
            else if ((*itDateReport)->IsVirtual())
            {
                if ((*itDateReport)->IsEdited() || (*itDateReport)->IsSaved())
                {
                    return *itDateReport;
                }
            }
        }
    }
    return pNullDateReport;
}

void CTUAssignmentReport::SetPreviousEmptyReportsEdited(CTUDateReport* pReport)
{
    TUDateReportsList::iterator itEmptyReports;
    for (itEmptyReports = m_DateReportsList.begin();
         itEmptyReports != m_DateReportsList.end();
         itEmptyReports++)
    {
        if ((*itEmptyReports)->GetReportDate() < pReport->GetReportDate())
        {
            if ((*itEmptyReports)->GetType() == TUAssgnTimeReports)
            {
                if ((!(*itEmptyReports)->IsEdited() && !(*itEmptyReports)->IsSaved() && !(*itEmptyReports)->IsVirtual()) || (*itEmptyReports)->GetStatus() == TURejected)
                {
                    (*itEmptyReports)->SetEdited(true);
                    (*itEmptyReports)->SetStatus(TUOpen);
                    (*itEmptyReports)->SetReportedWork(0);
                    (*itEmptyReports)->SetReportedOvertimeWork(0);
                }
            }
        }
        else if ((*itEmptyReports)->GetReportDate() == pReport->GetReportDate())
        {
            if ((*itEmptyReports)->GetType() == TUAssgnTimeReports)
            {
                if (!(*itEmptyReports)->IsEdited() || (*itEmptyReports)->GetStatus() == TURejected)
                {
                    (*itEmptyReports)->SetEdited(true);
                    (*itEmptyReports)->SetStatus(TUOpen);
                    (*itEmptyReports)->SetReportedWork(0);
                    (*itEmptyReports)->SetReportedOvertimeWork(0);
                }
            }
        }
    }
}

void CTUAssignmentReport::SetPreviousEmptyReportsStatus(CTUDateReport* pReport, enumTUReportsStatus eStatus)
{
    TUDateReportsList::iterator itEmptyReports;
    for (itEmptyReports = m_DateReportsList.begin();
         itEmptyReports != m_DateReportsList.end();
         itEmptyReports++)
    {
        if ((*itEmptyReports)->GetReportDate() <= pReport->GetReportDate())
        {
            if ((*itEmptyReports)->GetType() == TUAssgnTimeReports)
            {
                (*itEmptyReports)->SetStatus(eStatus);
            }
        }
    }
}

void CTUAssignmentReport::SetNexEmptyReportsRejected(CTUDateReport* pReport)
{
    TUDateReportsList::iterator itEmptyReports;
    for (itEmptyReports = m_DateReportsList.begin();
         itEmptyReports != m_DateReportsList.end();
         itEmptyReports++)
    {
        if ((*itEmptyReports)->GetReportDate() >= pReport->GetReportDate())
        {
            if ((*itEmptyReports)->GetType() == TUAssgnTimeReports)
            {
                (*itEmptyReports)->SetStatus(TURejected);
            }
        }
    }
}

void CTUAssignmentReport::ClearNextReports(CTUDateReport* pReport)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->GetReportDate() >= pReport->GetReportDate() && (*itDateReport)->GetType() == TUAssgnTimeReports)
        {
            (*itDateReport)->EmtpyReport();
        }
    }
}

void CTUAssignmentReport::MarkEditedReportsAsSaved()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if ((*itDateReport)->IsEdited())
        {
            (*itDateReport)->SetEdited(false);
            (*itDateReport)->SetSaved();
        }
    }
}

void CTUAssignmentReport::GetDateRange(CTime& tFrom, CTime& tTo)
{
    TUDateReportsList::iterator itDateReport;
    if (!m_DateReportsList.empty())
    {
        itDateReport = m_DateReportsList.begin();
        tFrom = (*itDateReport)->GetReportDate();

        itDateReport = m_DateReportsList.end() - 1;
        tTo = (*itDateReport)->GetReportDate();
        return;
    }
    tFrom = 0;
    tTo = 0;
    return;
}

bool CTUAssignmentReport::GetMarkedAsCompletedForDateReport(CTUDateReport* pDateReport)
{
    if (!IsMarkedAsCompleted())
        return false;

    if (GetWorkInsertMode() == TUAssgnPercentReports)
    {
        return true;
    }
    else
    {
        if (m_DateReportsList.empty())
            return false;
        else
        {
            TUDateReportsList::iterator itDateReport = m_DateReportsList.end() - 1;
            if ((*itDateReport)->GetReportGUID() == pDateReport->GetReportGUID())
                return true;
            else
                return false;
        }
    }
    return false;
}

void CTUAssignmentReport::SetReportsType()
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if (m_WorkInsertMode == TUAssgnTimeReports)
            (*itDateReport)->SetType(TUAssgnTimeReports);
        else if (m_WorkInsertMode == TUAssgnPercentReports)
            (*itDateReport)->SetType(TUAssgnPercentReports);
        else
            (*itDateReport)->SetType(TUAssgnUndefiniedReport);
    }
}

void CTUAssignmentReport::SetAllReportsStatus(enumTUReportsStatus eStatus)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        (*itDateReport)->SetStatus(eStatus);
    }
}
void CTUAssignmentReport::ChangeAllReportsStatus(enumTUReportsStatus eFromStatus, enumTUReportsStatus eToStatus)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        if (eFromStatus == (*itDateReport)->GetStatus() || TUAll == eFromStatus)
        {
            (*itDateReport)->SetStatus(eToStatus);
        }
    }
}
void CTUAssignmentReport::SetAllReportsType(enumTUAssgnWorkInsertMode eType)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();
         itDateReport++)
    {
        (*itDateReport)->SetType(eType);
    }
}
void CTUAssignmentReport::RemoveReportsWithStatus(enumTUReportsStatus eStatus)
{
    TUDateReportsList::iterator itDateReport;
    for (itDateReport = m_DateReportsList.begin();
         itDateReport != m_DateReportsList.end();)
    {
        if ((*itDateReport)->GetStatus() == eStatus)
        {
            delete *itDateReport;
            *itDateReport = NULL;
            itDateReport = m_DateReportsList.erase(itDateReport);
        }
        else
            itDateReport++;
    }
}

void CTUAssignmentReport::UpdateCurrentValues()
{
    if (GetWorkInsertMode() == TUAssgnPercentReports)
    {
        TUDateReportsList::iterator itDateReport;
        for (itDateReport = m_DateReportsList.begin();
            itDateReport != m_DateReportsList.end();
			itDateReport++)
        {
            if ((*itDateReport)->GetStatus() == TUAccepted)
            {
                (*itDateReport)->SetCurrentPercentComplete((*itDateReport)->GetCurrentPercentComplete());
            }
        }
    }
}
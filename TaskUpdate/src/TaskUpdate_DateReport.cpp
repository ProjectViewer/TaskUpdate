#include "TaskUpdate_DateReport.h"

CTUDateReport::CTUDateReport(void)
{
}

CTUDateReport::~CTUDateReport(void)
{
}

CString CTUDateReport::GetReportGUID() const
{
    return m_ReportGUID;
}
void CTUDateReport::SetReportGUID(CString strReportGUID)
{
    m_ReportGUID = strReportGUID;
}
CTime CTUDateReport::GetCreationDate() const
{
    return m_CreationDate;
}
void CTUDateReport::SetCreationDate(CTime tCreationDate)
{
    m_CreationDate = tCreationDate;
}
CTime CTUDateReport::GetReportDate() const
{
    return m_ReportDate;
}
void CTUDateReport::SetReportDate(CTime tReportDate)
{
    m_ReportDate = tReportDate;
}
int CTUDateReport::GetPlannedWork() const
{
    return m_PlannedWork;
}
void CTUDateReport::SetPlannedWork(int iPlannedWork)
{
    m_PlannedWork = iPlannedWork;
}
int CTUDateReport::GetPlannedOvertimeWork() const
{
    return m_PlannedOvertimeWork;
}
void CTUDateReport::SetPlannedOvertimeWork(int iPlannedOvertimeWork)
{
    m_PlannedOvertimeWork = iPlannedOvertimeWork;
}
int CTUDateReport::GetReportedWork() const
{
    return m_ReportedWork;
}
void CTUDateReport::SetReportedWork(int iReportedWork)
{
    m_ReportedWork = iReportedWork;
}
int CTUDateReport::GetReportedOvertimeWork() const
{
    return m_ReportedOvertimeWork;
}
void CTUDateReport::SetReportedOvertimeWork(int iReportedOvertimeWork)
{
    m_ReportedOvertimeWork = iReportedOvertimeWork;
}
int CTUDateReport::GetCurrentPercentComplete() const
{
    return m_CurrentPercentComplete;
}
void CTUDateReport::SetCurrentPercentComplete(int iCurrentPercentComplete)
{
    m_CurrentPercentComplete = iCurrentPercentComplete;
}
int CTUDateReport::GetReportedPercentComplete() const
{
    return m_ReportedPercentComplete;
}
void CTUDateReport::SetReportedPercentComplete(int iReportedPercentComplete)
{
    m_ReportedPercentComplete = iReportedPercentComplete;
}
CString CTUDateReport::GetTMNote() const
{
    return m_TMNote;
}
void CTUDateReport::SetTMNote(CString note)
{
    m_TMNote = note;
}
CString CTUDateReport::GetPMNote() const
{
    return m_PMNote;
}
void CTUDateReport::SetPMNote(CString note)
{
    m_PMNote = note;
}
bool CTUDateReport::HasNote() const
{
    return (!m_TMNote.IsEmpty() || !m_PMNote.IsEmpty());
}
bool CTUDateReport::HasTMNote() const
{
    return !m_TMNote.IsEmpty();
}
CRect CTUDateReport::GetPercRect() const
{
    return m_PercRect;
}
void CTUDateReport::SetPercRect(CRect rect)
{
    m_PercRect = rect;
}
CRect CTUDateReport::GetWorkRect() const
{
    return m_WorkRect;
}
void CTUDateReport::SetWorkRect(CRect rect)
{
    m_WorkRect = rect;
}
CRect CTUDateReport::GetOvtWorkRect() const
{
    return m_OvtWorkRect;
}
void CTUDateReport::SetOvtWorkRect(CRect rect)
{
    m_OvtWorkRect = rect;
}
enumTUReportsStatus CTUDateReport::GetStatus() const
{
    return m_Status;
}
void CTUDateReport::SetStatus(enumTUReportsStatus eStatus)
{
    m_Status = eStatus;
}
enumTUAssgnWorkInsertMode CTUDateReport::GetType() const
{
    return m_Type;
}
void CTUDateReport::SetType(enumTUAssgnWorkInsertMode eType)
{
    m_Type = eType;
}
bool CTUDateReport::IsEdited() const
{
    return m_Edited;
}
void CTUDateReport::SetEdited(bool bEdited)
{
    m_Edited = bEdited;
}
bool CTUDateReport::IsVirtual() const
{
    return m_Virtual;
}
void CTUDateReport::SetVirtual(bool bVirtual)
{
    m_Virtual = bVirtual;
}
bool CTUDateReport::IsActive() const
{
    return m_Active;
}
void CTUDateReport::SetActive(bool bActive)
{
    m_Active = bActive;
}
bool CTUDateReport::IsSaved() const
{
    return m_Saved;
}
void CTUDateReport::SetSaved(bool bSaved)
{
    m_Saved = bSaved;
}
bool CTUDateReport::IsDisabled() const
{
    return m_Disabled;
}
void CTUDateReport::SetDisabled(bool bDisabled)
{
    m_Disabled = bDisabled;
}
UINT_PTR CTUDateReport::GetUINT_PTR() const
{
    return m_UINT_PTR;
}
void CTUDateReport::SetUINT_PTR(UINT_PTR uint_ptr)
{
    m_UINT_PTR = uint_ptr;
}
bool CTUDateReport::IsSelected() const
{
    return m_Selected;
}
void CTUDateReport::SetSelected(bool bSel)
{
    m_Selected = bSel;
}

void CTUDateReport::EmtpyReport()
{
    SetReportedWork(-1);
    SetReportedOvertimeWork(-1);
    SetReportedPercentComplete(-1);
    SetTMNote(_T(""));
    SetStatus(TUOpen);
    SetEdited(false);
    SetSaved(false);
}

void CTUDateReport::ClearReport()
{
    SetReportedWork(-1);
    SetReportedOvertimeWork(-1);
    SetReportedPercentComplete(-1);
    SetTMNote(_T(""));
    SetStatus(TUOpen);
    SetEdited(false);
    SetSaved(false);
    SetActive(true);
}
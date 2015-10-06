#include "TaskUpdate_TaskReport.h"

CTUTaskReport::CTUTaskReport(void)
{
    m_CustomText1FieldID = 0;
    m_CustomText2FieldID = 0;
    m_CustomText3FieldID = 0;
    m_IsSummary = false;
}

CTUTaskReport::~CTUTaskReport(void)
{

}

CString CTUTaskReport::GetReportGUID() const
{
    return m_reportGUID;
}
void CTUTaskReport::SetReportGUID(CString guid)
{
    m_reportGUID = guid;
}
CTime CTUTaskReport::GetCreationDate() const
{
    return m_CreationDate;
}
void CTUTaskReport::SetCreationDate(CTime date)
{
    m_CreationDate  = date;
}

int CTUTaskReport::GetCurrentPercentComplete() const
{
    return m_CurrentPercentComplete;
}
void CTUTaskReport::SetCurrentPercentComplete(int val)
{
    m_CurrentPercentComplete = val;
}
int CTUTaskReport::GetReportedPercentComplete() const
{
    return m_ReportedPercentComplete;
}
void CTUTaskReport::SetReportedPercentComplete(int val)
{
    m_ReportedPercentComplete = val;
}

int CTUTaskReport::GetCurrentPercentWorkComplete() const
{
    return m_CurrentPercentWorkComplete;
}
void CTUTaskReport::SetCurrentPercentWorkComplete(int val)
{
    m_CurrentPercentWorkComplete = val;
}
int CTUTaskReport::GetReportedPercentWorkComplete() const
{
    return m_ReportedPercentWorkComplete;
}
void CTUTaskReport::SetReportedPercentWorkComplete(int val)
{
    m_ReportedPercentWorkComplete = val;
}

int CTUTaskReport::GetCurrentActualWork() const
{
    return m_CurrentActualWork;
}
void CTUTaskReport::SetCurrentActualWork(int val)
{
    m_CurrentActualWork = val;
}
int CTUTaskReport::GetReportedActualWork() const
{
    return m_ReportedActualWork;
}
void CTUTaskReport::SetReportedActualWork(int val)
{
    m_ReportedActualWork = val;
}

CTime CTUTaskReport::GetCurrentActualStart() const
{
	return m_CurrentActualStart;
}
void CTUTaskReport::SetCurrentActualStart(CTime val)
{
	m_CurrentActualStart = val;
}
CTime CTUTaskReport::GetReportedActualStart() const
{
	return m_ReportedActualStart;
}
void CTUTaskReport::SetReportedActualStart(CTime val)
{
	m_ReportedActualStart = val;
}

CTime CTUTaskReport::GetCurrentActualFinish() const
{
	return m_CurrentActualFinish;
}
void CTUTaskReport::SetCurrentActualFinish(CTime val)
{
	m_CurrentActualFinish = val;
}
CTime CTUTaskReport::GetReportedActualFinish() const
{
	return m_ReportedActualFinish;
}
void CTUTaskReport::SetReportedActualFinish(CTime val)
{
	m_ReportedActualFinish = val;
}

int CTUTaskReport::GetCurrentActualDuration() const
{
    return m_CurrentActualDuration;
}
void CTUTaskReport::SetCurrentActualDuration(int val)
{
    m_CurrentActualDuration = val;
}
int CTUTaskReport::GetReportedActualDuration() const
{
    return m_ReportedActualDuration;
}
void CTUTaskReport::SetReportedActualDuration(int val)
{
    m_ReportedActualDuration = val;
}

CString CTUTaskReport::GetCurrentNotes() const
{
    return m_CurrentNotes;
}
void CTUTaskReport::SetCurrentNotes(CString val)
{
    m_CurrentNotes = val;
}
CString CTUTaskReport::GetReportedNotes() const
{
    return m_ReportedNotes;
}
void CTUTaskReport::SetReportedNotes(CString val)
{
    m_ReportedNotes = val;
}

CString CTUTaskReport::GetCustomText1FieldText() const
{
    return m_CustomText1FieldText;
}
void CTUTaskReport::SetCustomText1FieldText(CString text)
{
    m_CustomText1FieldText = text;
}
CString CTUTaskReport::GetCustomText2FieldText() const
{
    return m_CustomText2FieldText;
}
void CTUTaskReport::SetCustomText2FieldText(CString text)
{
    m_CustomText2FieldText = text;
}
CString CTUTaskReport::GetCustomText3FieldText() const
{
    return m_CustomText3FieldText;
}
void CTUTaskReport::SetCustomText3FieldText(CString text)
{
    m_CustomText3FieldText = text;
}

CString CTUTaskReport::GetCustomTextFieldText(unsigned int fieldID) const
{
    CString str = _T("");
    if (m_CustomText1FieldID == fieldID)
    {
        str = m_CustomText1FieldText;
    }
    else if (m_CustomText2FieldID == fieldID)
    {
        str = m_CustomText2FieldText;
    }
    else if (m_CustomText3FieldID == fieldID)
    {
        str = m_CustomText3FieldText;
    }

    return str;
}
void CTUTaskReport::SetCustomTextFieldText(unsigned int fieldID, CString text)
{
    if (m_CustomText1FieldID == fieldID)
    {
        m_CustomText1FieldText = text;
    }
    else if (m_CustomText2FieldID == fieldID)
    {
        m_CustomText2FieldText = text;
    }
    else if (m_CustomText3FieldID == fieldID)
    {
        m_CustomText3FieldText = text;
    }
}


CString CTUTaskReport::GetTMNote() const
{
    return m_TMNote;
}
void CTUTaskReport::SetTMNote(CString note)
{
    m_TMNote = note;
}
CString CTUTaskReport::GetPMNote() const
{
    return m_PMNote;
}
void CTUTaskReport::SetPMNote(CString note)
{
    m_PMNote = note;
}
bool CTUTaskReport::HasNote() const
{
    return (!m_TMNote.IsEmpty() || !m_PMNote.IsEmpty());
}
bool CTUTaskReport::HasTMNote() const
{
    return !m_TMNote.IsEmpty();
}

bool CTUTaskReport::ContainsPercRectPoint(CPoint point)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_PercRect.begin(); itRect != m_PercRect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            return true;
        }
    }
    return false;
}
bool CTUTaskReport::ContainsPercRectPoint(CPoint point, CRect& rc)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_PercRect.begin(); itRect != m_PercRect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            rc = *itRect;
            return true;
        }
    }
    return false;
}
void CTUTaskReport::AddPercRect(CRect rect)
{
    m_PercRect.push_back(rect);
}
bool CTUTaskReport::ContainsPercWorkRectPoint(CPoint point)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_PercWorkRect.begin(); itRect != m_PercWorkRect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            return true;
        }
    }
    return false;
}
bool CTUTaskReport::ContainsPercWorkRectPoint(CPoint point, CRect& rc)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_PercWorkRect.begin(); itRect != m_PercWorkRect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            rc = *itRect;
            return true;
        }
    }
    return false;
}
void CTUTaskReport::AddPercWorkRect(CRect rect)
{
    m_PercWorkRect.push_back(rect);
}
bool CTUTaskReport::ContainsActWorkRectPoint(CPoint point)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_ActWorkRect.begin(); itRect != m_ActWorkRect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            return true;
        }
    }
    return false;
}
bool CTUTaskReport::ContainsActWorkRectPoint(CPoint point, CRect& rc)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_ActWorkRect.begin(); itRect != m_ActWorkRect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            rc = *itRect;
            return true;
        }
    }
    return false;
}
void CTUTaskReport::AddActWorkRect(CRect rect)
{
    m_ActWorkRect.push_back(rect);
}
bool CTUTaskReport::ContainsActStartRectPoint(CPoint point)
{
	std::vector<CRect>::iterator itRect;
	for (itRect = m_ActStartRect.begin(); itRect != m_ActStartRect.end(); itRect++)
	{
		if (itRect->PtInRect(point))
		{
			return true;
		}
	}
	return false;
}
bool CTUTaskReport::ContainsActStartRectPoint(CPoint point, CRect& rc)
{
	std::vector<CRect>::iterator itRect;
	for (itRect = m_ActStartRect.begin(); itRect != m_ActStartRect.end(); itRect++)
	{
		if (itRect->PtInRect(point))
		{
			rc = *itRect;
			return true;
		}
	}
	return false;
}
void CTUTaskReport::AddActStartRect(CRect rect)
{
	m_ActStartRect.push_back(rect);
}
bool CTUTaskReport::ContainsActFinishRectPoint(CPoint point)
{
	std::vector<CRect>::iterator itRect;
	for (itRect = m_ActFinishRect.begin(); itRect != m_ActFinishRect.end(); itRect++)
	{
		if (itRect->PtInRect(point))
		{
			return true;
		}
	}
	return false;
}
bool CTUTaskReport::ContainsActFinishRectPoint(CPoint point, CRect& rc)
{
	std::vector<CRect>::iterator itRect;
	for (itRect = m_ActFinishRect.begin(); itRect != m_ActFinishRect.end(); itRect++)
	{
		if (itRect->PtInRect(point))
		{
			rc = *itRect;
			return true;
		}
	}
	return false;
}
void CTUTaskReport::AddActFinishRect(CRect rect)
{
	m_ActFinishRect.push_back(rect);
}
bool CTUTaskReport::ContainsActDurationRectPoint(CPoint point)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_ActDurationRect.begin(); itRect != m_ActDurationRect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            return true;
        }
    }
    return false;
}
bool CTUTaskReport::ContainsActDurationRectPoint(CPoint point, CRect& rc)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_ActDurationRect.begin(); itRect != m_ActDurationRect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            rc = *itRect;
            return true;
        }
    }
    return false;
}
void CTUTaskReport::AddActDurationRect(CRect rect)
{
    m_ActDurationRect.push_back(rect);
}
bool CTUTaskReport::ContainsNotesRectPoint(CPoint point)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_NotesRect.begin(); itRect != m_NotesRect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            return true;
        }
    }
    return false;
}
bool CTUTaskReport::ContainsNotesRectPoint(CPoint point, CRect& rc)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_NotesRect.begin(); itRect != m_NotesRect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            rc = *itRect;
            return true;
        }
    }
    return false;
}
void CTUTaskReport::AddNotesRect(CRect rect)
{
    m_NotesRect.push_back(rect);
}

bool CTUTaskReport::ContainsCustomText1RectPoint(CPoint point)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_CustomText1Rect.begin(); itRect != m_CustomText1Rect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            return true;
        }
    }
    return false;
}
bool CTUTaskReport::ContainsCustomText1RectPoint(CPoint point, CRect& rc)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_CustomText1Rect.begin(); itRect != m_CustomText1Rect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            rc = *itRect;
            return true;
        }
    }
    return false;
}
void CTUTaskReport::AddCustomText1Rect(CRect rect)
{
    m_CustomText1Rect.push_back(rect);
}
bool CTUTaskReport::ContainsCustomText2RectPoint(CPoint point)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_CustomText2Rect.begin(); itRect != m_CustomText2Rect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            return true;
        }
    }
    return false;
}
bool CTUTaskReport::ContainsCustomText2RectPoint(CPoint point, CRect& rc)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_CustomText2Rect.begin(); itRect != m_CustomText2Rect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            rc = *itRect;
            return true;
        }
    }
    return false;
}
void CTUTaskReport::AddCustomText2Rect(CRect rect)
{
    m_CustomText2Rect.push_back(rect);
}
bool CTUTaskReport::ContainsCustomText3RectPoint(CPoint point)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_CustomText3Rect.begin(); itRect != m_CustomText3Rect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            return true;
        }
    }
    return false;
}
bool CTUTaskReport::ContainsCustomText3RectPoint(CPoint point, CRect& rc)
{
    std::vector<CRect>::iterator itRect;
    for (itRect = m_CustomText3Rect.begin(); itRect != m_CustomText3Rect.end(); itRect++)
    {
        if (itRect->PtInRect(point))
        {
            rc = *itRect;
            return true;
        }
    }
    return false;
}
void CTUTaskReport::AddCustomText3Rect(CRect rect)
{
    m_CustomText3Rect.push_back(rect);
}
bool CTUTaskReport::ContainsCustomTextRectPoint(unsigned int fieldID, CPoint point)
{
    bool returnValue = false;

    if (m_CustomText1FieldID == fieldID)
    {
        returnValue = ContainsCustomText1RectPoint(point);
    }
    else if (m_CustomText2FieldID == fieldID)
    {
        returnValue = ContainsCustomText2RectPoint(point);
    }
    if (m_CustomText3FieldID == fieldID)
    {
        returnValue = ContainsCustomText3RectPoint(point);
    }

    return returnValue;
}
bool CTUTaskReport::ContainsCustomTextRectPoint(unsigned int fieldID, CPoint point, CRect& rc)
{
    bool returnValue = false;

    if (m_CustomText1FieldID == fieldID)
    {
        returnValue = ContainsCustomText1RectPoint(point, rc);
    }
    else if (m_CustomText2FieldID == fieldID)
    {
        returnValue = ContainsCustomText2RectPoint(point, rc);
    }
    if (m_CustomText3FieldID == fieldID)
    {
        returnValue = ContainsCustomText3RectPoint(point, rc);
    }

    return returnValue;
}
void CTUTaskReport::AddCustomTextRect(unsigned int fieldID, CRect rect)
{
    if (m_CustomText1FieldID == fieldID)
    {
        m_CustomText1Rect.push_back(rect);
    }
    else if (m_CustomText2FieldID == fieldID)
    {
        m_CustomText2Rect.push_back(rect);
    }
    if (m_CustomText3FieldID == fieldID)
    {
        m_CustomText3Rect.push_back(rect);
    }
    
}

void CTUTaskReport::ClearReportRects()
{
    m_PercRect.clear();
    m_PercWorkRect.clear();
    m_ActWorkRect.clear();
	m_ActStartRect.clear();
	m_ActFinishRect.clear();
    m_ActDurationRect.clear();
    m_NotesRect.clear();
    m_CustomText1Rect.clear();
    m_CustomText2Rect.clear();
    m_CustomText3Rect.clear();
}
enumTUTaskReportsStatus CTUTaskReport::GetStatus() const
{
    return m_Status;
}
void CTUTaskReport::SetStatus(enumTUTaskReportsStatus eStatus)
{
    m_Status = eStatus;
}

void CTUTaskReport::SetSummary(bool isSummary)
{
    m_IsSummary = isSummary;
}
bool CTUTaskReport::IsSummary() const
{
    return m_IsSummary;
}

bool CTUTaskReport::IsPercentCompleteEdited() const
{
    return m_PercentCompleteEdited;
}
bool CTUTaskReport::IsPercentWorkCompleteEdited() const
{
    return m_PercentWorkCompleteEdited;
}
bool CTUTaskReport::IsActualWorkEdited() const
{
    return m_ActualWorkEdited;
}
bool CTUTaskReport::IsActualStartEdited() const
{
	return m_ActualStartEdited;
}
bool CTUTaskReport::IsActualFinishEdited() const
{
	return m_ActualFinishEdited;
}
bool CTUTaskReport::IsActualDurationEdited() const
{
    return m_ActualDurationEdited;
}
bool CTUTaskReport::IsNotesEdited() const
{
    return m_NotesEdited;
}
bool CTUTaskReport::IsCustomText1Edited() const
{
    return m_CustomText1Edited;
}
bool CTUTaskReport::IsCustomText2Edited() const
{
    return m_CustomText2Edited;
}
bool CTUTaskReport::IsCustomText3Edited() const
{
    return m_CustomText3Edited;
}
bool CTUTaskReport::IsCustomTextEdited(unsigned int fieldID) const
{
    bool returnValue = false;

    if (m_CustomText1FieldID == fieldID)
    {
        returnValue =  m_CustomText1Edited;
    }
    else if (m_CustomText2FieldID == fieldID)
    {
        returnValue =  m_CustomText2Edited;
    }
    else if (m_CustomText3FieldID == fieldID)
    {
        returnValue =  m_CustomText3Edited;
    }

    return returnValue;
}
 bool CTUTaskReport::IsEdited(bool textField) const
 {
     return m_PercentCompleteEdited || m_PercentWorkCompleteEdited || m_ActualWorkEdited || m_ActualStartEdited || m_ActualFinishEdited || m_ActualDurationEdited ||
         (textField && (m_CustomText1Edited || m_CustomText2Edited || m_CustomText3Edited || m_NotesEdited));
 }

void CTUTaskReport::SetPercentCompleteEdited(bool edited)
{
    m_PercentCompleteEdited = edited;
    m_PercentWorkCompleteEdited = false;
    m_ActualWorkEdited = false;
	m_ActualStartEdited = false;
	m_ActualFinishEdited = false;
    m_ActualDurationEdited = false;
}
void CTUTaskReport::SetPercentWorkCompleteEdited(bool edited)
{
    m_PercentCompleteEdited = false;
    m_PercentWorkCompleteEdited = edited;
    m_ActualWorkEdited = false;
	m_ActualStartEdited = false;
	m_ActualFinishEdited = false;
    m_ActualDurationEdited = false;
}
void CTUTaskReport::SetActualWorkEdited(bool edited)
{
    m_PercentCompleteEdited = false;
    m_PercentWorkCompleteEdited = false;
    m_ActualWorkEdited = edited;
	m_ActualStartEdited = false;
	m_ActualFinishEdited = false;
    m_ActualDurationEdited = false;
}
void CTUTaskReport::SetActualStartEdited(bool edited)
{
	m_PercentCompleteEdited = false;
	m_PercentWorkCompleteEdited = false;
	m_ActualWorkEdited = false;
	m_ActualStartEdited = edited;
	m_ActualFinishEdited = false;
    m_ActualDurationEdited = false;
}
void CTUTaskReport::SetActualFinishEdited(bool edited)
{
	m_PercentCompleteEdited = false;
	m_PercentWorkCompleteEdited = false;
	m_ActualWorkEdited = false;
	m_ActualStartEdited = false;
	m_ActualFinishEdited = edited;
    m_ActualDurationEdited = false;
}
void CTUTaskReport::SetActualDurationEdited(bool edited)
{
    m_PercentCompleteEdited = false;
    m_PercentWorkCompleteEdited = false;
    m_ActualWorkEdited = false;
    m_ActualStartEdited = false;
    m_ActualFinishEdited = false;
    m_ActualDurationEdited = edited;
}
void CTUTaskReport::SetNotesEdited(bool edited)
{
    m_NotesEdited = edited;
}
void CTUTaskReport::SetCustomText1Edited(bool edited)
{
    m_CustomText1Edited = edited;
}
void CTUTaskReport::SetCustomText2Edited(bool edited)
{
    m_CustomText2Edited = edited;
}
void CTUTaskReport::SetCustomText3Edited(bool edited)
{
    m_CustomText3Edited = edited;
}
void CTUTaskReport::SetCustomTextEdited(unsigned int fieldID, bool edited)
{
    if (m_CustomText1FieldID == fieldID)
    {
        m_CustomText1Edited = edited;
    }
    else if (m_CustomText2FieldID == fieldID)
    {
        m_CustomText2Edited = edited;
    }
    else if (m_CustomText3FieldID == fieldID)
    {
        m_CustomText3Edited = edited;
    }
}

bool CTUTaskReport::IsSaved() const
{
    return m_Saved;
}
void CTUTaskReport::SetSaved(bool bSaved)
{
    m_Saved = bSaved;
}
bool CTUTaskReport::IsCustomTextSaved() const
{
    return m_CustomTextSaved;
}
void CTUTaskReport::SetCustomTextSaved(bool bSaved)
{
    m_CustomTextSaved = bSaved;
}
bool CTUTaskReport::IsNotesSaved() const
{
    return m_NotesSaved;
}
void CTUTaskReport::SetNotesSaved(bool bSaved)
{
    m_NotesSaved = bSaved;
}
bool CTUTaskReport::IsDisabled() const
{
    return m_Disabled;
}
void CTUTaskReport::SetDisabled(bool bDisabled)
{
    m_Disabled = bDisabled;
}
UINT_PTR CTUTaskReport::GetUINT_PTR() const
{
    return m_UINT_PTR;
}
void CTUTaskReport::SetUINT_PTR(UINT_PTR uint_ptr)
{
    m_UINT_PTR = uint_ptr;
}
bool CTUTaskReport::IsSelected() const
{
    return m_Selected;
}
void CTUTaskReport::SetSelected(bool bSel)
{
    m_Selected = bSel;
}

long long CTUTaskReport::GetTaskUID() const
{
    return m_TaskUID;
}
void CTUTaskReport::SetTaskUID(long long iTaskUID)
{
    m_TaskUID = iTaskUID;
}
CString CTUTaskReport::GetProjectName() const
{
    return m_ProjectName;
}
void CTUTaskReport::SetProjectName(CString prjName)
{
    m_ProjectName = prjName;
}
CString CTUTaskReport::GetProjectGUID() const
{
    return m_ProjectGUID;
}
void CTUTaskReport::SetProjectGUID(CString prjGUID)
{
    m_ProjectGUID = prjGUID;
}
long long CTUTaskReport::GetTaskBaseUID() const
{
    return m_TaskBaseUID;
}
void CTUTaskReport::SetTaskBaseUID(long long iBaseUID)
{
    m_TaskBaseUID = iBaseUID;
}
CString CTUTaskReport::GetTaskName() const
{
    return m_TaskName;
}
void CTUTaskReport::SetTaskName(CString strTaskName)
{
    m_TaskName = strTaskName;
}
enumTUTaskWorkInsertMode CTUTaskReport::GetWorkInsertMode() const
{
    return m_WorkInsertMode;
}
void CTUTaskReport::SetWorkInsertMode(enumTUTaskWorkInsertMode eWorkInsertMode)
{
    m_WorkInsertMode = eWorkInsertMode;
}
bool CTUTaskReport::IsUndefined()
{
    return m_WorkInsertMode == TUTaskUndefiniedReport;
}

bool CTUTaskReport::IsPercentCompleteActive()
{
    return m_WorkInsertMode == TUTaskPercentCompleteReport;
}
bool CTUTaskReport::IsPercentWorkCompleteActive()
{
    return m_WorkInsertMode == TUTaskPercentWorkCompleteReport;
}
bool CTUTaskReport::IsActualWorkActive()
{
    return m_WorkInsertMode == TUTaskActualWorkReport;
}
bool CTUTaskReport::IsActualStartActive()
{
	return m_WorkInsertMode == TUTaskActualStartReport;
}
bool CTUTaskReport::IsActualFinishActive()
{
	return m_WorkInsertMode == TUTaskActualFinishReport;
}
bool CTUTaskReport::IsActualDurationActive()
{
    return m_WorkInsertMode == TUTaskActualDurationReport;
}
void CTUTaskReport::EmptyReport(bool emptyTextField)
{
    SetReportedPercentComplete(-1);
    SetReportedPercentWorkComplete(-1);
    SetReportedActualWork(-1);
	SetReportedActualStart(CTime(0));
	SetReportedActualFinish(CTime(0));
    SetReportedActualDuration(-1);
    

    SetTMNote(_T(""));
    SetStatus(TUTaskOpen);
    SetPercentCompleteEdited(false);
    SetPercentWorkCompleteEdited(false);
    SetActualWorkEdited(false);
	SetActualStartEdited(false);
	SetActualFinishEdited(false);
    SetActualDurationEdited(false);
    

    if (emptyTextField)
    {
        SetReportedNotes(_T(""));
        SetNotesEdited(false);
        SetNotesSaved(false);

        SetCustomText1FieldText(_T(""));
        SetCustomText1Edited(false);

        SetCustomText2FieldText(_T(""));
        SetCustomText2Edited(false);

        SetCustomText3FieldText(_T(""));
        SetCustomText3Edited(false);

        SetCustomTextSaved(false);
    }

    SetSaved(false);
}
bool CTUTaskReport::IsSaveable()
{
    return (IsEdited() || IsSaved() || IsCustomTextSaved() || IsNotesSaved() || GetStatus() == TUTaskRejected);
}

unsigned int CTUTaskReport::GetCustomText1FieldID() const
{
    return m_CustomText1FieldID;
}
void CTUTaskReport::SetCustomText1FieldID(unsigned int filedID)
{
    m_CustomText1FieldID = filedID;
}
unsigned int CTUTaskReport::GetCustomText2FieldID() const
{
    return m_CustomText2FieldID;
}
void CTUTaskReport::SetCustomText2FieldID(unsigned int filedID)
{
    m_CustomText2FieldID = filedID;
}
unsigned int CTUTaskReport::GetCustomText3FieldID() const
{
    return m_CustomText3FieldID;
}
void CTUTaskReport::SetCustomText3FieldID(unsigned int filedID)
{
    m_CustomText3FieldID = filedID;
}

bool CTUTaskReport::IsCustomTextField(unsigned int fieldID)
{
    return m_CustomText1FieldID == fieldID || m_CustomText2FieldID == fieldID || m_CustomText3FieldID == fieldID;

}
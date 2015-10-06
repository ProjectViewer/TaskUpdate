#pragma once

enum enumTUAssgnWorkInsertMode
{
    TUAssgnUndefiniedReport      = 0,
    TUAssgnTimeReports           = 1,
    TUAssgnPercentReports        = 2
};

enum enumTUReportsStatus
{
    TUOpen                  = 0,
    TURejected              = 1,
    TUAccepted              = 2,
    TUPending               = 3,
    TUModified              = 4,
    TUToBeAccepted          = 5,
    TUToBeRejected          = 6,
    TUAll                   = 7
};

class CTUDateReport
{
public:
    CTUDateReport(void);
    ~CTUDateReport(void);

private:
    CString m_ReportGUID;
    CTime m_CreationDate;
    CTime m_ReportDate;
    int m_PlannedWork;
    int m_PlannedOvertimeWork;
    int m_ReportedWork;
    int m_ReportedOvertimeWork;
    int m_CurrentPercentComplete;
    int m_ReportedPercentComplete;
    CString m_TMNote;
    CString m_PMNote;
    CRect m_PercRect;
    CRect m_WorkRect;
    CRect m_OvtWorkRect;
    enumTUReportsStatus m_Status;
    bool m_Edited;                   //for SPV
    bool m_Virtual;                   //for SPV
    bool m_Active;                   //for SPV
    enumTUAssgnWorkInsertMode m_Type;   //for SPV
    bool m_Saved;                   //for SPV
    bool m_Disabled;                   //for SPV
    UINT_PTR m_UINT_PTR;               //for SPV, to store pointer address for tooltips
    bool m_Selected;                   //for SPV, GUI

public:
    CString GetReportGUID() const;
    void SetReportGUID(CString strReportGUID);
    CTime GetCreationDate() const;
    void SetCreationDate(CTime tCreationDate);
    CTime GetReportDate() const;
    void SetReportDate(CTime tReportDate);
    int GetPlannedWork() const;
    void SetPlannedWork(int iPlannedWork);
    int GetPlannedOvertimeWork() const;
    void SetPlannedOvertimeWork(int iPlannedOvertimeWork);
    int GetReportedWork() const;
    void SetReportedWork(int iReportedWork);
    int GetReportedOvertimeWork() const;
    void SetReportedOvertimeWork(int iReportedOvertimeWork);
    int GetCurrentPercentComplete() const;
    void SetCurrentPercentComplete(int iCurrentPercentComplete);
    int GetReportedPercentComplete() const;
    void SetReportedPercentComplete(int iReportedPercentComplete);
    CString GetTMNote() const;
    void SetTMNote(CString note);
    CString GetPMNote() const;
    void SetPMNote(CString note);
    bool HasNote() const;
    bool HasTMNote() const;
    CRect GetPercRect() const;
    void SetPercRect(CRect rect);
    CRect GetWorkRect() const;
    void SetWorkRect(CRect rect);
    CRect GetOvtWorkRect() const;
    void SetOvtWorkRect(CRect rect);
    enumTUReportsStatus GetStatus() const;
    void SetStatus(enumTUReportsStatus eStatus);
    enumTUAssgnWorkInsertMode GetType() const;
    void SetType(enumTUAssgnWorkInsertMode eType);
    bool IsEdited() const;                   //for SPV
    void SetEdited(bool bEdited = true);                   //for SPV
    bool IsVirtual() const;                   //for SPV
    void SetVirtual(bool bVirtual = true);                   //for SPV
    bool IsActive() const;                   //for SPV
    void SetActive(bool bActive = true);                   //for SPV
    bool IsSaved() const;                   //for SPV
    void SetSaved(bool bSaved = true);                   //for SPV
    bool IsDisabled() const;                   //for SPV
    void SetDisabled(bool bDisabled = true);                   //for SPV
    UINT_PTR GetUINT_PTR() const;               //for SPV
    void SetUINT_PTR(UINT_PTR uint_ptr);                   //for SPV
    bool IsSelected() const;                   //for SPV
    void SetSelected(bool bSel);                   //for SPV

    void EmtpyReport();                   //for SPV
    void ClearReport();                   //for SPV
};

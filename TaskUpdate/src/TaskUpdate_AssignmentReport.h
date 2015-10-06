#pragma once
#include "TaskUpdate_DateReport.h"
#include <vector>

enum enumMarkAsCompleteType
{
    TU_MAC_None      = 0,
    TU_MAC_Percent   = 1,
    TU_MAC_Finish    = 2
};

typedef std::vector<CTUDateReport*> TUDateReportsList;

class CTUAssignmentReport
{
public:
    CTUAssignmentReport(void);
    ~CTUAssignmentReport(void);

private:
    long long m_AssignmentUID;
    CString m_ProjectName;
    CString m_ProjectGUID;
    long long m_AssignmentBaseUID;
    CString m_ResourceName;
    long long m_ResourceUID;
    long long m_TaskUID;
    CString m_TaskName;
    enumTUAssgnWorkInsertMode m_WorkInsertMode;
    bool m_MarkAsCompleted;
    TUDateReportsList m_DateReportsList;
    enumMarkAsCompleteType m_MarkAsCompleteType; //For Add-In
    static CTUDateReport* pNullDateReport;

public:
    long long GetAssignmentUID() const;
    void SetAssignmentUID(long long iAssgnUID);
    CString GetProjectName() const;
    void SetProjectName(CString prName);
    CString GetProjectGUID() const;
    void SetProjectGUID(CString prjGUID);
    long long GetAssignmentBaseUID() const;
    void SetAssignmentBaseUID(long long iBaseUID);
    long long GetResourceUID() const;
    void SetResourceUID(long long iResUID);
    CString GetResourceName() const;
    void SetResourceName(CString strResName);
    long long GetTaskUID() const;
    void SetTaskUID(long long iTaskUID);
    CString GetTaskName() const;
    void SetTaskName(CString strTaskName);
    enumTUAssgnWorkInsertMode GetWorkInsertMode() const;
    void SetWorkInsertMode(enumTUAssgnWorkInsertMode eWorkInsertMode);
    bool IsMarkedAsCompleted() const;
    void SetMarkedAsCompleted(bool bMarkAsCompleted = true);
    enumMarkAsCompleteType GetMarkAsCompleteType() const;
    void SetMarkAsCompleteType(enumMarkAsCompleteType m_MarkAsCompleteType);

    CTUDateReport*              GetPercentReport();
    CTUDateReport*              GetTimeReport(CTime reportTime);
    CTUDateReport*              GetPercentReport(bool bActive); //for SPV
    CTUDateReport*              GetTimeReport(CTime reportTime, bool bActive);  //for SPV
    CTUDateReport*              GetDateReport(CString GUID);
    CTUDateReport*              GetReportByUINT_PTR(UINT_PTR uint_ptr); //for SPV

    TUDateReportsList*          GetDateReportsList();
    void AddDateReport(CTUDateReport* pDateReport);
    void RemoveDateReport(CTUDateReport* pDateReport);
    void RemoveDateReports();

    void Clear();

    bool HasReports() const;
    bool HasSaveableReport();                           //for SPV
    bool HasRejectedReport();
    bool HasAcceptedReport();
    bool HasSavedReport();                          //for SPV
    bool HasEditedReport();                         //for SPV
    bool IsDisabled();
    bool HasTMNote();   
    void EmptyInactiveReports();                        //for SPV
    void EmptyReports();                        //forSPV
    void ClearReports();                        //forSPV
    void ClearReportRects(enumTUAssgnWorkInsertMode eWorkInsertMode);                        //for SPV
    bool HasPreviousEmptyReports(CTUDateReport* pReport);                           //for SPV
    bool HasPreviousDateReports(CTUDateReport* pReport);                        //for Add-In
    bool HasNextDateReports(CTUDateReport* pReport);                        //for Add-In
    CTUDateReport*              GetFirstFlexEmptyReport();  //for SPV
    void SetPreviousEmptyReportsEdited(CTUDateReport* pReport);                         //for SPV
    void SetPreviousEmptyReportsStatus(CTUDateReport* pReport, enumTUReportsStatus eStatus);                        //for Add-In
    void SetNexEmptyReportsRejected(CTUDateReport* pReport);                        //for Add-In
    void SetAllReportsStatus(enumTUReportsStatus eStatus);                          //for Add-In
    void ChangeAllReportsStatus(enumTUReportsStatus eFromStatus, enumTUReportsStatus eToStatus);                        //for Add-In
    void SetAllReportsType(enumTUAssgnWorkInsertMode eType);                         //for Add-In
    void ClearNextReports(CTUDateReport* pReport);                          //for SPV
    void MarkEditedReportsAsSaved();                        //forSPV

    void GetDateRange(CTime& tFrom, CTime& tTo);
    bool GetMarkedAsCompletedForDateReport(CTUDateReport* pDateReport);
    void SetReportsType();                       //For Add-In
    void RemoveReportsWithStatus(enumTUReportsStatus eStatus);                        //for Add-In
    void UpdateCurrentValues();
};

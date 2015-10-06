#pragma once
#include "TaskUpdate_AssignmentReport.h"
#include "TaskUpdate_TaskReport.h"

enum enumTUProjectReportingMode
{
    TUTaskReportingMode       = 0,
    TUAssignmentReportingMode = 1
};

enum enumTUReportsType
{
    TUCurrentReports        = 0,
    TUReferentReports       = 1
};

typedef std::vector<CTUAssignmentReport*> TUAssignmentReportsList;
typedef std::vector<CTUTaskReport*> TUTaskReportsList;

class CTUProjectReport
{
public:
    CTUProjectReport(void);
    ~CTUProjectReport(void);

private:
    CString m_ProjectGUID;
    CString m_ProjectName;
    enumTUProjectReportingMode m_workInsertMode;
    TUTaskReportsList m_TaskReportsList;
    TUAssignmentReportsList m_AssignmentReportsList;


    static CTUTaskReport* pNullTaskReport;
    static CTUAssignmentReport* pNullAssgnReport;
    static CTUDateReport* pNullDateReport;

public:
    CString GetProjectGUID() const;
    void SetProjectGUID(CString strPrjGUID);
    CString GetProjectName() const;
    void SetProjectName(CString strPrjName);
    enumTUProjectReportingMode GetWorkInsertMode() const;
    void SetWorkInsertMode (enumTUProjectReportingMode mode);

    //task reports
    TUTaskReportsList*              GetTaskReportsList();
    CTUTaskReport*                  GetTaskReport(long long iTaskUID);
    CTUTaskReport*                  GetTaskReport(CString GUID);
    CTUTaskReport*                  GetTaskReportByUINT_PTR(UINT_PTR uint_ptr); //for SPV
    void                            AddTaskReport(CTUTaskReport* pTaskReport);
    void                            RemoveTaskReport(CTUTaskReport* pTaskReport);

    //assignment reports
    TUAssignmentReportsList*        GetAssignmentReportsList();
    CTUAssignmentReport*            GetAssgnReport(long long assgnUID);
    CTUAssignmentReport*            GetAssgnReport(long long iTaskUID, long long iResUID);
    void                            AddAssgnReport(CTUAssignmentReport* pAssgnReport);
    void                            RemoveAssgnReport(CTUAssignmentReport* pAssgnReport);
    CTUDateReport*                  GetPercentReport(long long assgnUID);
    CTUDateReport*                  GetTimeReport(long long assgnUID, CTime reportTime);
    CTUDateReport*                  GetPercentReport(long long assgnUID, bool bActive); //for SPV
    CTUDateReport*                  GetTimeReport(long long assgnUID, CTime reportTime, bool bActive); //for SPV
    CTUDateReport*                  GetDateReport(CString GUID);
    CTUDateReport*                  GetDateReportByUINT_PTR(UINT_PTR uint_ptr); //for SPV

    void Clear();

    bool HasRejectedReport();
    bool HasAcceptedReport();
    bool HasSaveableReport();                           //for SPV
    bool HasEditedReport();                           //for SPV
    bool HasTMNote();
    void SetReportsCreationDate();                           //call before saving the xml to write the current date
    void MarkEditedReportsAsSaved();                           //for SPV
    void EmptyInactiveReports();                           //for SPV
    void AddBaseUID(long long iBaseUID);                           //to add baseUID to assignments original UID(retrieve the original UID when writing to the xml)
    void RemoveBaseUID(long long iBaseUID);                           //remove baseUID from assignments original UID
    CTUAssignmentReport*            GetDatesAssignmentReport(CTUDateReport* pDateReport);
    void GetAssignmentReportsListForTask(long long iTaskUID, TUAssignmentReportsList& theList);
    void GetAssignmentReportsListForResource(long long iResUID, TUAssignmentReportsList& theList);
    void SetReportsType();                           //For Add-In
    void SetAllReportsStatus(int eStatus);                              //for Add-In
    void ChangeAllReportsStatus(int eFromStatus, int eToStatus);                            //for Add-In
    void RemoveReportsWithStatus(CString strPtojGUID, int eStatus);                            //for Add-In
    void RemoveReportsWithStatus(int eStatus);                            //for Add-In
    void ClearTaskReportRects();                            //for SPV
    void ClearReportRects(enumTUAssgnWorkInsertMode eWorkInsertMode);                            //for SPV
    void UpdateCurrentValues();
    bool IsTaskUpdateCustomTextField(long long taskUID, int fieldID);
	bool HasReportWithWorkInsertMode(enumTUTaskWorkInsertMode eWorkInsertMode);
};

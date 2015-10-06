#pragma once
#include "TaskUpdate_ProjectReport.h"
#include "TaskUpdate_ProjectInfo.h"
typedef	std::pair<CString,CFile*>	FileHandlesPair;
typedef std::map<CString, CFile*>	FileHandlesList;

struct TaskIndexesStruct 
{
	long taskIndex;
	long taskUID;
	MSProject::TaskPtr task;
	CString name;
};

typedef std::vector<TaskIndexesStruct> TaskIndexesList;

class CUpdatesManager
{
public:
    CUpdatesManager(void);
    ~CUpdatesManager(void);

public:
    CTUProjectReport*	getProjReport();
	TUProjectsInfoList* getProjInfoList();
	bool				saveSPVFiles(TUProjectsInfoList* pProjInfoList, bool removeNotes = false, bool removeCustomTexts = false);
	void				lockSPVFiles(TUProjectsInfoList* pProjectInfoList);
	void				unLockSPVFiles(TUProjectsInfoList* pProjectInfoList);
	void				processWritingToFiles(TUProjectsInfoList* pProjInfoList, bool& bCannotWriteManual, int reportingMode);
    void                getUpdatesViaGantt();
    void                getCustomTextUpdatesViaGantt(bool& success, bool& hasCustomtextReport);
    void                getNotesViaGantt(bool& success, bool& hasNoteReport);
    void                removeReportsWithStatus(int eStatus); 
	void				reopenProject();

private: 
	void				fillProjectInfoList();
	void				fillSubProjectsInfoList(MSProject::_MSProjectPtr app, MSProject::_IProjectDocPtr project, CTUProjectInfo* pProjectInfo, int parentBaseUID);
	void				processSPVFiles(TUProjectsInfoList* pProjInfoList);
	void				getReportsFromSpv(CTUProjectInfo* pProjInfo);
	bool				saveSPVFile(CTUProjectInfo* pPrjInfo, bool removeNotes = false, bool removeCustomTexts = false);
	void				getReportsForProject(CTUProjectInfo* pPrjInfo, CTUProjectReport* pProjectReport, CTUProjectReport& projectReportOut, int eStatus);
               
    void				writeAssgnReportsToFile(CTUProjectInfo* pPrjInfo, bool& bCannotWriteManual);
    void				writeTaskReportsToFile(CTUProjectInfo* pPrjInfo, bool& bCannotWriteManual);

    void				processShowingUpdatesOnGantt(TUProjectsInfoList* pProjInfoList);
    void				processShowingNotesOnGantt(TUProjectsInfoList* pProjInfoList, bool& success, bool& hasNoteReport);
    void                showNotesOnGantt(CTUProjectInfo* pPrjInfo, bool& success, bool& hasNoteReport);
    void				processShowingCustomTextUpdatesOnGantt(TUProjectsInfoList* pProjInfoList, bool& success, bool& hasCustomtextReport);
    void                showTaskUpdatesOnGantt(CTUProjectInfo* pPrjInfo);
    void                showCustomTextUpdatesOnGantt(CTUProjectInfo* pPrjInfo, bool& success, bool& hasCustomtextReport);

	void				updateWorkByAssigment(CTUAssignmentReport* pAssgnReport, MSProject::TaskPtr prjTask);
	void				updateWorkByTask(MSProject::_IProjectDocPtr project, CTUAssignmentReport* pAssgnReport, MSProject::TaskPtr prjTask);
	void				updatePercent_TaskWithAssigment(CTUAssignmentReport* pAssgnReport, MSProject::TaskPtr prjTask);
	void				updatePercent_TaskWithUnassignedResources(CTUAssignmentReport* pAssgnReport, MSProject::TaskPtr prjTask);

    void                updateTask_PercentComplete(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask);
    void                updateTask_PercentWorkComplete(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask);
    void                updateTask_ActualWork(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask);
	void                updateTask_ActualStart(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask);
	void                updateTask_ActualFinish(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask);
    void                updateTask_ActualDuration(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask);

    void                getProjectMapValues(MSProject::_IProjectDocPtr project);
    void                showTask_PercentCompleteReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
    void                showTask_PercentWorkCompleteReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
    void                showTask_ActualWorkReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
	void                showTask_ActualStartReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
	void                showTask_ActualFinishReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
    void                showTask_ActualDurationReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);

    void                showTask_PercentCompleteDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
    void                showTask_PercentWorkCompleteDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
    void                showTask_ActualWorkDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
	void                showTask_ActualStartDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
	void                showTask_ActualFinishDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
    void                showTask_ActualDurationDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);

    void                showTask_PercentCompleteGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
    void                showTask_PercentWorkCompleteGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
    void                showTask_ActualWorkGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
	void                showTask_ActualStartGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
	void                showTask_ActualFinishGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);
    void                showTask_ActualDurationGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err);

    void                showTask_Notes(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err, bool& hasNoteReport);
    void                showTask_CustomText1Field(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err, bool& hasCustomTextReport);
    void                showTask_CustomText2Field(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err, bool& hasCustomTextReport);
    void                showTask_CustomText3Field(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err, bool& hasCustomTextReport);

    void                removeAcceptedNotes(CTUProjectReport& projectReport);
    void                removeAcceptedCustomFields(CTUProjectReport& projectReport);
    void                removeUndefinedEmptyReports(CTUProjectReport& projectReport);

	void				initTaskIndexes(TUProjectsInfoList* pProjInfoList);
	void				setTaskIndexes(MSProject::_IProjectDocPtr project);
	MSProject::TaskPtr	taskByUID(MSProject::_IProjectDocPtr project, long uid);

private:
	TUProjectsInfoList			m_ProjInfoList;
	CTUProjectReport			m_ProjectReport;

    FileHandlesList				m_FileHandlerList;
    MSProject::_MSProjectPtr	m_app;

    int m_iPercComplReport;
    int m_iPercComplDiff;
    int m_iPercComplGraphInd;
    int m_iPercWorkComplReport;
    int m_iPercWorkComplDiff;
    int m_iPercWorkComplGraphInd;
    int m_iActWorkReport;
    int m_iActWorkDiff;
    int m_iActWorkGraphInd;
	int m_iActStartReport;
	int m_iActStartDiff;
	int m_iActStartGraphInd;
	int m_iActFinishReport;
	int m_iActFinishDiff;
	int m_iActFinishGraphInd;
    int m_iActDurationReport;
    int m_iActDurationDiff;
    int m_iActDurationGraphInd;

    int m_iCustomTextField1;
    int m_iCustomTextField2;
    int m_iCustomTextField3;


	std::map<MSProject::_IProjectDocPtr, TaskIndexesList> m_taskIndexes;
};


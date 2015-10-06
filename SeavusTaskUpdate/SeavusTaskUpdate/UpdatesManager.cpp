#include "StdAfx.h"
#include "UpdatesManager.h"
#include "TaskUpdate_Utils.h"
#include "Utils.h"
#include "TaskUpdate_XMLParser.h"

CUpdatesManager::CUpdatesManager(void)
{
    fillProjectInfoList();
	initTaskIndexes(&m_ProjInfoList);
	processSPVFiles(&m_ProjInfoList);
}


CUpdatesManager::~CUpdatesManager(void)
{
   
}

 CTUProjectReport*		CUpdatesManager::getProjReport()
{
    return &m_ProjectReport;
}
 TUProjectsInfoList*	CUpdatesManager::getProjInfoList()
 {
	 return &m_ProjInfoList;
 }

void					CUpdatesManager::lockSPVFiles(TUProjectsInfoList* pProjectInfoList) 
{
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjectInfoList->begin(); itPrjInfo != pProjectInfoList->end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->HasChilds())
        {
            lockSPVFiles((*itPrjInfo)->GetSubProjectsList());
        }

        CFile* pFile  = new CFile();
        if (pFile->Open((*itPrjInfo)->GetProjectSPVPath(),CFile::modeRead | CFile::shareExclusive))
        {
            m_FileHandlerList.insert(FileHandlesPair((*itPrjInfo)->GetProjectSPVPath(), pFile));
        }
    }
}
void					CUpdatesManager::unLockSPVFiles(TUProjectsInfoList* pProjectInfoList)
{
	TUProjectsInfoList::iterator itPrjInfo;
	for (itPrjInfo = pProjectInfoList->begin(); itPrjInfo != pProjectInfoList->end(); ++itPrjInfo)
	{
		if ((*itPrjInfo)->HasChilds())
		{
			unLockSPVFiles((*itPrjInfo)->GetSubProjectsList());
		}

		FileHandlesList::iterator itFile = m_FileHandlerList.find((*itPrjInfo)->GetProjectSPVPath());
		if(itFile != m_FileHandlerList.end())
		{
			itFile->second->Close();
			delete itFile->second;
			itFile->second = NULL;
		}
	}
}

void					CUpdatesManager::fillProjectInfoList()
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
    m_app = app;

    MSProject::_IProjectDocPtr project = app->GetActiveProject();  
    if ((BSTR)project->FullName)
    {
		CTUProjectInfo* pRootPrjInfo = new CTUProjectInfo();

        CString updatesDirPath = _T("");
        CString pathSPV = _T("");
        pRootPrjInfo->SetProjectName((BSTR)project->Name);
        pRootPrjInfo->SetProjectPath((BSTR)project->FullName);
        pRootPrjInfo->SetProjectGUID(CTU_Utils::GetCollIDFromFile((BSTR)project->FullName));
		pRootPrjInfo->SetProjectBaseUID(0);
        CTU_Utils::GetFilePaths((BSTR)project->FullName, pathSPV, updatesDirPath);
        pRootPrjInfo->SetProjectSPVPath(pathSPV);
        pRootPrjInfo->SetProjectReportingMode(CTU_Utils::GetReportingModeFromFile(pRootPrjInfo->GetProjectPath()));

		m_ProjInfoList.push_back(pRootPrjInfo);

        fillSubProjectsInfoList(app, project, pRootPrjInfo, 0x00400000);
    }
}
void					CUpdatesManager::fillSubProjectsInfoList(MSProject::_MSProjectPtr app, MSProject::_IProjectDocPtr project, CTUProjectInfo* pProjectInfo, int parentBaseUID)
{
    MSProject::SubprojectsPtr subProjectsList = project->Subprojects;

    for (int i = 1; i <= subProjectsList->GetCount(); i++)
    {
        if (MSProject::SubprojectPtr subProject = subProjectsList->Item[i])
        {         
            if (MSProject::_IProjectDocPtr subProjectDoc = subProject->GetSourceProject())
            {
                int baseUID = parentBaseUID + (subProject->Index * 0x00400000);

                CTUProjectInfo* pSubProjectInfo = pProjectInfo->AddSubProject(CTU_Utils::GetCollIDFromFile((BSTR)subProjectDoc->FullName), 
                    (BSTR)subProjectDoc->Name, 
                    baseUID,    
                    (BSTR)subProjectDoc->FullName, 
                    0); 

                fillSubProjectsInfoList(app, subProjectDoc, pSubProjectInfo, baseUID);
            }
        }
    }
}

void					CUpdatesManager::initTaskIndexes(TUProjectsInfoList* pProjInfoList)
{
	TUProjectsInfoList::iterator itPrjInfo;
	for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
	{
		MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
		MSProject::_IProjectDocPtr project = app->GetActiveProject(); 
		MSProject::_IProjectDocPtr currentProject = NULL;
		if (0 != wcscmp((*itPrjInfo)->GetProjectPath(), project->FullName))
		{
			currentProject = Utils::getSubProject(project, (*itPrjInfo)->GetProjectPath());
			if (!currentProject)
			{
				return;
			}	
		}
		else
		{
			currentProject = project;
		}


		setTaskIndexes(currentProject);

		initTaskIndexes((*itPrjInfo)->GetSubProjectsList());
	}
}
void					CUpdatesManager::setTaskIndexes(MSProject::_IProjectDocPtr project)
{
	std::vector<TaskIndexesStruct> temp;
	for (int i = 1; i <= project->Tasks->GetCount(); i++)
	{
		MSProject::TaskPtr pTask = project->Tasks->GetItem(i);
		if (pTask)
		{
			TaskIndexesStruct struc;
			struc.taskIndex = i;
			struc.taskUID = pTask->UniqueID;
			struc.task = pTask;
			struc.name = (BSTR)pTask->Name;

			temp.push_back(struc);
		}
	}
	m_taskIndexes[project] = temp; 
}

MSProject::TaskPtr	CUpdatesManager::taskByUID(MSProject::_IProjectDocPtr project, long uid)
{
	MSProject::TaskPtr task = 0;

	auto itList = m_taskIndexes.find(project);
	if (itList != m_taskIndexes.end())
	{
		TaskIndexesList& list = itList->second;
		for (auto it = list.begin(); it != list.end(); it++)
		{
			TaskIndexesStruct* struc = &(*it);
			if ( uid == struc->taskUID)
			{
				task = struc->task;
				break;
			}

		}
	}


	return task;
}

void					CUpdatesManager::processSPVFiles(TUProjectsInfoList* pProjInfoList)
{
	TUProjectsInfoList::iterator itPrjInfo;
	for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
	{
		getReportsFromSpv(*itPrjInfo);

		//for all sub projects
		processSPVFiles((*itPrjInfo)->GetSubProjectsList());
	}
}
void					CUpdatesManager::getReportsFromSpv(CTUProjectInfo* pProjInfo)
{
    int reportingMode = CTU_Utils::GetReportingModeFromFile(pProjInfo->GetProjectPath());

	CTUProjectReport* tempReports = new CTUProjectReport();
	CTU_Utils::LoadXML(pProjInfo->GetProjectSPVPath(),*tempReports, reportingMode);//load only task or assignment reports

	tempReports->AddBaseUID(pProjInfo->GetProjectBaseUID());

    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    MSProject::_IProjectDocPtr project = app->GetActiveProject(); 
    MSProject::_IProjectDocPtr currentProject = NULL;
    if (0 != wcscmp(pProjInfo->GetProjectPath(), project->FullName))
    {
        currentProject = Utils::getSubProject(project, pProjInfo->GetProjectPath());
        if (!currentProject)
        {
            return;
        }	
    }
    else
    {
        currentProject = project;
    }



    if (TUTaskReportingMode == reportingMode)
    {
        TUTaskReportsList::iterator itTaskReport;
        for(itTaskReport = tempReports->GetTaskReportsList()->begin();
            itTaskReport != tempReports->GetTaskReportsList()->end();
            itTaskReport++)
        {
            MSProject::TaskPtr prjTask = taskByUID(currentProject, (long)((*itTaskReport)->GetTaskUID() - pProjInfo->GetProjectBaseUID()));
            if (prjTask)
            {
                (*itTaskReport)->SetCurrentPercentComplete(prjTask->PercentComplete);
                (*itTaskReport)->SetCurrentPercentWorkComplete(prjTask->PercentWorkComplete);
                (*itTaskReport)->SetCurrentActualWork(prjTask->ActualWork);
                (*itTaskReport)->SetCurrentActualDuration(prjTask->ActualDuration);

				COleDateTime odt_start;
				COleDateTime odt_finish;
				_bstr_t bstr_start = prjTask->ActualStart;
				_bstr_t bstr_finish = prjTask->ActualFinish;
				odt_start.ParseDateTime(bstr_start, LOCALE_NOUSEROVERRIDE,LANG_USER_DEFAULT); 
				odt_finish.ParseDateTime(bstr_finish, LOCALE_NOUSEROVERRIDE,LANG_USER_DEFAULT); 

				CTime time_start;
				CTime time_finish;
				time_start = (odt_start.m_status == COleDateTime::invalid) ? CTime(0) : CTime(odt_start.GetYear(), odt_start.GetMonth(), odt_start.GetDay(), odt_start.GetHour(), odt_start.GetMinute(), 0);
				time_finish = (odt_finish.m_status == COleDateTime::invalid) ? CTime(0) : CTime(odt_finish.GetYear(), odt_finish.GetMonth(), odt_finish.GetDay(), odt_finish.GetHour(), odt_finish.GetMinute(), 0);
				(*itTaskReport)->SetCurrentActualStart(time_start);
				(*itTaskReport)->SetCurrentActualFinish(time_finish);  

                //change base UID, as it is needed for subproject task UIDs when filtering
                m_ProjectReport.AddTaskReport((*itTaskReport));
			}
        }
    }
    else if (TUAssignmentReportingMode == reportingMode)
    {
        TUAssignmentReportsList::iterator itAssgnReport;
        for(itAssgnReport = tempReports->GetAssignmentReportsList()->begin();
            itAssgnReport != tempReports->GetAssignmentReportsList()->end();
            itAssgnReport++)
        {
            MSProject::TaskPtr prjTask = taskByUID(currentProject,  (long)((*itAssgnReport)->GetTaskUID() - pProjInfo->GetProjectBaseUID()));
            MSProject::AssignmentPtr prjAssignment = Utils::getPjAssignment(currentProject, prjTask, (long)((*itAssgnReport)->GetAssignmentUID() - pProjInfo->GetProjectBaseUID()));
            if (prjTask && (prjAssignment || prjTask->Assignments->Count == 0))
            {
                //change base UID, as it is needed for subproject task UIDs when filtering
                m_ProjectReport.AddAssgnReport((*itAssgnReport));
                //date report types are not parsed, we get them from  their parent assignment
                (*itAssgnReport)->SetAllReportsType((*itAssgnReport)->GetWorkInsertMode());
            }
        }
    }
}

bool					CUpdatesManager::saveSPVFiles(TUProjectsInfoList* pProjInfoList, bool removeNotes, bool removeCustomTexts)
{
	bool bSaveSuccess = true;
	TUProjectsInfoList::iterator itPrjInfo;
	for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
	{
		bSaveSuccess = bSaveSuccess && saveSPVFile(*itPrjInfo, removeNotes, removeCustomTexts);

		//for all sub projects
		bSaveSuccess = bSaveSuccess && saveSPVFiles((*itPrjInfo)->GetSubProjectsList(), removeNotes, removeCustomTexts);
	}

	return bSaveSuccess;
}
bool					CUpdatesManager::saveSPVFile(CTUProjectInfo* pPrjInfo, bool removeNotes, bool removeCustomTexts)
{
	CTUProjectReport reportsToSave;
	getReportsForProject(pPrjInfo, &m_ProjectReport, reportsToSave, TUAll);

	//retrieve the original UIDs before writing them to SPV file
	reportsToSave.RemoveBaseUID(pPrjInfo->GetProjectBaseUID());

    if (removeNotes)
    {
        removeAcceptedNotes(reportsToSave);
    }

    if (removeCustomTexts)
    {
        removeAcceptedNotes(reportsToSave);
        removeAcceptedCustomFields(reportsToSave);
    }

    removeUndefinedEmptyReports(reportsToSave);

	return CTU_Utils::SaveXML(pPrjInfo->GetProjectSPVPath(), &reportsToSave);
}
void					CUpdatesManager::getReportsForProject(CTUProjectInfo* pPrjInfo, CTUProjectReport* pProjectReport, CTUProjectReport& projectReportOut, int eStatus)
{
	projectReportOut.SetProjectGUID(pPrjInfo->GetProjectGUID());
	projectReportOut.SetProjectName(pPrjInfo->GetProjectName());

    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = pProjectReport->GetTaskReportsList()->begin();
        itTaskReport != pProjectReport->GetTaskReportsList()->end();
        itTaskReport++)
    {
        if((*itTaskReport)->GetProjectGUID() == pPrjInfo->GetProjectGUID())
        {
            if(eStatus == (*itTaskReport)->GetStatus() || eStatus == TUTaskAll)
            {
                CTUTaskReport* pTaskReportOut = new CTUTaskReport();

                pTaskReportOut->SetTaskUID((*itTaskReport)->GetTaskUID());
                pTaskReportOut->SetProjectGUID((*itTaskReport)->GetProjectGUID());
                pTaskReportOut->SetProjectName((*itTaskReport)->GetProjectName());
                pTaskReportOut->SetTaskName((*itTaskReport)->GetTaskName());
                pTaskReportOut->SetWorkInsertMode((*itTaskReport)->GetWorkInsertMode());
                pTaskReportOut->SetReportGUID((*itTaskReport)->GetReportGUID());
                pTaskReportOut->SetCreationDate((*itTaskReport)->GetCreationDate());

                pTaskReportOut->SetCurrentPercentComplete((*itTaskReport)->GetCurrentPercentComplete());
                pTaskReportOut->SetReportedPercentComplete((*itTaskReport)->GetReportedPercentComplete());
                pTaskReportOut->SetCurrentPercentWorkComplete((*itTaskReport)->GetCurrentPercentWorkComplete());
                pTaskReportOut->SetReportedPercentWorkComplete((*itTaskReport)->GetReportedPercentWorkComplete());
                pTaskReportOut->SetCurrentActualWork((*itTaskReport)->GetCurrentActualWork());
                pTaskReportOut->SetReportedActualWork((*itTaskReport)->GetReportedActualWork());
				pTaskReportOut->SetCurrentActualStart((*itTaskReport)->GetCurrentActualStart());
				pTaskReportOut->SetReportedActualStart((*itTaskReport)->GetReportedActualStart());
				pTaskReportOut->SetCurrentActualFinish((*itTaskReport)->GetCurrentActualFinish());
				pTaskReportOut->SetReportedActualFinish((*itTaskReport)->GetReportedActualFinish());
                pTaskReportOut->SetCurrentActualDuration((*itTaskReport)->GetCurrentActualDuration());
                pTaskReportOut->SetReportedActualDuration((*itTaskReport)->GetReportedActualDuration());
                pTaskReportOut->SetCurrentNotes((*itTaskReport)->GetCurrentNotes());
                pTaskReportOut->SetReportedNotes((*itTaskReport)->GetReportedNotes());

                pTaskReportOut->SetCustomText1FieldText((*itTaskReport)->GetCustomText1FieldText());
                pTaskReportOut->SetCustomText2FieldText((*itTaskReport)->GetCustomText2FieldText());
                pTaskReportOut->SetCustomText3FieldText((*itTaskReport)->GetCustomText3FieldText());

                pTaskReportOut->SetTMNote((*itTaskReport)->GetTMNote());
                pTaskReportOut->SetPMNote((*itTaskReport)->GetPMNote());
                pTaskReportOut->SetStatus((*itTaskReport)->GetStatus());

                projectReportOut.AddTaskReport(pTaskReportOut);
            }
        }
    }

    TUDateReportsList::iterator		itDateReport;
    TUAssignmentReportsList::iterator	itAssgnReport;
	for (itAssgnReport = pProjectReport->GetAssignmentReportsList()->begin();
		itAssgnReport != pProjectReport->GetAssignmentReportsList()->end();
		itAssgnReport++)
	{
		if((*itAssgnReport)->GetProjectGUID() == pPrjInfo->GetProjectGUID())
		{
			CTUAssignmentReport* pAssgnReportOut = new CTUAssignmentReport();
			pAssgnReportOut->SetAssignmentUID((*itAssgnReport)->GetAssignmentUID());
			pAssgnReportOut->SetProjectGUID((*itAssgnReport)->GetProjectGUID());
			pAssgnReportOut->SetProjectName((*itAssgnReport)->GetProjectName());
			pAssgnReportOut->SetResourceName((*itAssgnReport)->GetResourceName());
			pAssgnReportOut->SetResourceUID((*itAssgnReport)->GetResourceUID());
			pAssgnReportOut->SetTaskName((*itAssgnReport)->GetTaskName());
			pAssgnReportOut->SetTaskUID((*itAssgnReport)->GetTaskUID());
			pAssgnReportOut->SetWorkInsertMode((*itAssgnReport)->GetWorkInsertMode());
			pAssgnReportOut->SetMarkedAsCompleted((*itAssgnReport)->IsMarkedAsCompleted());
            pAssgnReportOut->SetMarkAsCompleteType((*itAssgnReport)->GetMarkAsCompleteType());

			for (itDateReport = (*itAssgnReport)->GetDateReportsList()->begin();
				itDateReport != (*itAssgnReport)->GetDateReportsList()->end();
				itDateReport++)
			{
				if((enumTUReportsStatus)eStatus == (*itDateReport)->GetStatus() || (enumTUReportsStatus)eStatus == TUAll)
				{
					CTUDateReport* pDateReportOut = new CTUDateReport();
					pDateReportOut->SetReportGUID((*itDateReport)->GetReportGUID());
					pDateReportOut->SetCreationDate((*itDateReport)->GetCreationDate());
					pDateReportOut->SetReportDate((*itDateReport)->GetReportDate());
					pDateReportOut->SetPlannedWork((*itDateReport)->GetPlannedWork());
					pDateReportOut->SetReportedWork((*itDateReport)->GetReportedWork());
					pDateReportOut->SetPlannedOvertimeWork((*itDateReport)->GetPlannedOvertimeWork());
					pDateReportOut->SetReportedOvertimeWork((*itDateReport)->GetReportedOvertimeWork());
					pDateReportOut->SetCurrentPercentComplete((*itDateReport)->GetCurrentPercentComplete());
					pDateReportOut->SetReportedPercentComplete((*itDateReport)->GetReportedPercentComplete());
					pDateReportOut->SetTMNote((*itDateReport)->GetTMNote());
					pDateReportOut->SetPMNote((*itDateReport)->GetPMNote());
					pDateReportOut->SetType((*itDateReport)->GetType());
					pDateReportOut->SetStatus((*itDateReport)->GetStatus());
					pAssgnReportOut->AddDateReport(pDateReportOut);
				}
			}

			//if assignment has no date reports remove it
			if(pAssgnReportOut->GetDateReportsList()->size())
			{
				projectReportOut.AddAssgnReport(pAssgnReportOut);
			}
		}
	}
}
void                    CUpdatesManager::removeReportsWithStatus(int eStatus)
{
    m_ProjectReport.RemoveReportsWithStatus(eStatus);
}

void					CUpdatesManager::updatePercent_TaskWithUnassignedResources(CTUAssignmentReport* pAssgnReport, MSProject::TaskPtr prjTask)
{
	TUDateReportsList::iterator itDateReport;
	for (itDateReport = pAssgnReport->GetDateReportsList()->begin(); 
		itDateReport !=pAssgnReport->GetDateReportsList()->end(); 
		itDateReport++ )
	{
		if (((*itDateReport)->GetStatus() == TUToBeAccepted) && (pAssgnReport->GetResourceUID() == 0) 
			&& (pAssgnReport->GetTaskUID() == prjTask->GetUniqueID())
			&& (*itDateReport)->GetReportedPercentComplete() >= 0) 
		{
            try
            {
                prjTask->PercentComplete = (*itDateReport)->GetReportedPercentComplete();
            }
            catch (...)
            {
            }
		}
	}
}
void					CUpdatesManager::updatePercent_TaskWithAssigment(CTUAssignmentReport* pAssgnReport, MSProject::TaskPtr prjTask)
{
	MSProject::AssignmentPtr prjAssgn;
	for (int i = 1;i <= prjTask->Assignments->GetCount(); i++)
	{
		prjAssgn = prjTask->Assignments->GetItem(i);
		if(prjAssgn)
		{
			if (( prjAssgn->GetUniqueID() == pAssgnReport->GetAssignmentUID())||
				( prjAssgn->GetUniqueID() == pAssgnReport->GetAssignmentUID()+0x100000))
			{
				TUDateReportsList::iterator itDateReport ;
				for (itDateReport = pAssgnReport->GetDateReportsList()->begin(); 
					itDateReport != pAssgnReport->GetDateReportsList()->end(); 
					itDateReport++ )
				{
					if (((*itDateReport)->GetStatus() == TUToBeAccepted) && ((*itDateReport)->GetReportedPercentComplete() >= 0 ))
					{
                        try
                        {
                            prjAssgn->PercentWorkComplete = (*itDateReport)->GetReportedPercentComplete();
                        }
                        catch (...)
                        {
                        }
					}
				}
			}
		}
	}
}
void					CUpdatesManager::updateWorkByAssigment(CTUAssignmentReport* pAssgnReport, MSProject::TaskPtr prjTask)
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
    _bstr_t appVersion = app->Version;
    int iMajorVersion = Utils::GetProjectMajorVersion((BSTR)appVersion);

	MSProject::AssignmentPtr prjAssgn;
	for (int i = 1; i <= prjTask->Assignments->GetCount(); i++)
	{
		prjAssgn = prjTask->Assignments->GetItem(i);
		if(prjAssgn)
		{
			if (( prjAssgn->GetUniqueID() == pAssgnReport->GetAssignmentUID())||
				( prjAssgn->GetUniqueID() == pAssgnReport->GetAssignmentUID()+0x100000))
			{
				TUDateReportsList::iterator itDateReport;
				for(itDateReport = pAssgnReport->GetDateReportsList()->begin(); 
					itDateReport != pAssgnReport->GetDateReportsList()->end(); 
					itDateReport++ )
				{
					if((*itDateReport)->GetStatus() == TUToBeAccepted)
					{
						MSProject::TimeScaleValuesPtr tempTimeScaleValuesPtr;
						MSProject::TimeScaleValuePtr tempTimeScalPtr;

						CTime tempDate = (*itDateReport)->GetReportDate();
                        COleDateTime oletime;
                        oletime.SetDate(Utils::GetYear(tempDate), Utils::GetMonth(tempDate), Utils::GetDay(tempDate));
                        CString strDateTime = oletime.Format();

						const _variant_t atemp = strDateTime.AllocSysString();

						tempTimeScaleValuesPtr = prjAssgn->TimeScaleData(atemp, atemp,MSProject::pjAssignmentTimescaledActualWork,MSProject::pjTimescaleDays,1);
                        if (tempTimeScaleValuesPtr)
                        {
                            for (int i = 1 ; i <= tempTimeScaleValuesPtr->GetCount();i++)
                            {
                                tempTimeScalPtr = tempTimeScaleValuesPtr->GetItem(i);
                                if ((tempTimeScalPtr) && ((*itDateReport)->GetReportedWork() > 0))
                                {
                                    tempTimeScaleValuesPtr->GetItem(i)->Value = (*itDateReport)->GetReportedWork() ;                                 
                                }
                            }
                        }

                        

                        BOOL isManual = 0;
                        if(iMajorVersion == 14)
                        {
                            isManual = prjTask->Manual; //0 is for auto, -1 for manual
                        }

                        if (isManual == 0)
                        {
                            tempTimeScaleValuesPtr = prjAssgn->TimeScaleData(atemp, atemp,MSProject::pjAssignmentTimescaledActualOvertimeWork,MSProject::pjTimescaleDays,1);
                            if (tempTimeScaleValuesPtr)
                            {
                                for (int i = 1 ; i <= tempTimeScaleValuesPtr->GetCount();i++)
                                {
                                    tempTimeScalPtr = tempTimeScaleValuesPtr->GetItem(i);
                                    if ((tempTimeScalPtr) && ((*itDateReport)->GetReportedOvertimeWork() > 0))
                                    {
                                        tempTimeScaleValuesPtr->GetItem(i)->Value = (*itDateReport)->GetReportedOvertimeWork() ;
                                    }
                                }
                            }
                        }
					}
				}

                if(pAssgnReport->GetWorkInsertMode() == TUAssgnTimeReports)
                {
                    if (pAssgnReport->GetMarkAsCompleteType() == TU_MAC_Percent)
                    {
                        try
                        {
                            prjAssgn->PercentWorkComplete = 100;
                        }
                        catch (...)
                        {
                        }
                    }
                    if(pAssgnReport->GetMarkAsCompleteType() == TU_MAC_Finish)
                    {
                        try
                        {
                            prjAssgn->Work = prjAssgn->ActualWork;
                        }
                        catch (...)
                        {
                        }               
                    }
                }
			}
		}
	}
}
void					CUpdatesManager::updateWorkByTask(MSProject::_IProjectDocPtr project, CTUAssignmentReport* pAssgnReport, MSProject::TaskPtr prjTask)
{
	TUDateReportsList::iterator itDateReport ;
	for (itDateReport = pAssgnReport->GetDateReportsList()->begin(); 
		itDateReport != pAssgnReport->GetDateReportsList()->end(); 
		itDateReport++ )
	{
		if (((*itDateReport)->GetStatus() == TUToBeAccepted)&&(pAssgnReport->GetResourceUID() == 0) 
			&& (pAssgnReport->GetTaskUID() == prjTask->GetUniqueID()))
		{
			if ((*itDateReport)->GetStatus() == TUToBeAccepted)
			{
				MSProject::TimeScaleValuesPtr tempTimeScaleValuesPtr;
				MSProject::TimeScaleValuePtr tempTimeScalPtr;

				CTime tempDate = (*itDateReport)->GetReportDate();
                COleDateTime oletime;
                oletime.SetDate(Utils::GetYear(tempDate), Utils::GetMonth(tempDate), Utils::GetDay(tempDate));
                CString strDateTime = oletime.Format();

				const _variant_t atemp = strDateTime.AllocSysString();

				tempTimeScaleValuesPtr = prjTask->TimeScaleData(atemp, atemp,MSProject::pjTaskTimescaledActualWork,MSProject::pjTimescaleDays,1);
                if (tempTimeScaleValuesPtr)
                {
                    for (int i = 1 ; i <= tempTimeScaleValuesPtr->GetCount();i++)
                    {
                        tempTimeScalPtr = tempTimeScaleValuesPtr->GetItem(i);
                        if (((tempTimeScalPtr) && (*itDateReport)->GetReportedWork() > 0) && ( project->Calendar->Period(atemp,atemp)->GetWorking()))
                        {
                            tempTimeScaleValuesPtr->GetItem(i)->Value = (*itDateReport)->GetReportedWork() ;
                        }
                    }
                }
			}
        }
	}

    if(pAssgnReport->GetWorkInsertMode() == TUAssgnTimeReports)
    {
        if (pAssgnReport->GetMarkAsCompleteType() == TU_MAC_Percent)
        {
            try
            {
                prjTask->PercentWorkComplete = 100;
            }
            catch (...)
            {
            }             
        }
        if(pAssgnReport->GetMarkAsCompleteType() == TU_MAC_Finish)
        {
            try
            {
                prjTask->Work = prjTask->ActualWork;
            }
            catch (...)
            {
            }          
        }
    }
}

void                    CUpdatesManager::updateTask_PercentComplete(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask)
{
    if (pTaskReport->GetReportedPercentComplete() >= 0)
    {
        try
        {
            prjTask->PercentComplete = pTaskReport->GetReportedPercentComplete();
        }
        catch (...)
        {
        }         
    }
}
void                    CUpdatesManager::updateTask_PercentWorkComplete(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask)
{
    if (pTaskReport->GetReportedPercentWorkComplete() >= 0)
    {
        try
        {
            prjTask->PercentWorkComplete = pTaskReport->GetReportedPercentWorkComplete();
        }
        catch (...)
        {
        }       
    }
}
void                    CUpdatesManager::updateTask_ActualWork(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask)
{
    if (pTaskReport->GetReportedActualWork() >= 0)
    {
        try
        {
            prjTask->ActualWork = pTaskReport->GetReportedActualWork();
        }
        catch (...)
        {
        }       
    }
}
void                    CUpdatesManager::updateTask_ActualStart(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask)
{
	if (pTaskReport->GetReportedActualStart() != CTime(0))
	{
		CTime timeVal = Utils::toGMT(pTaskReport->GetReportedActualStart());
        CString strDateTime = Utils::get_date_time_in_user_format(timeVal);
		const _variant_t varTime = strDateTime.AllocSysString();

        try
        {
            prjTask->ActualStart = varTime;
        }
        catch (...)
        {
        }	
	}
}
void                    CUpdatesManager::updateTask_ActualFinish(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask)
{
	if (pTaskReport->GetReportedActualFinish() != CTime(0))
	{
		CTime timeVal = Utils::toGMT(pTaskReport->GetReportedActualFinish());
        CString strDateTime = Utils::get_date_time_in_user_format(timeVal);
		const _variant_t varTime = strDateTime.AllocSysString();

        try
        {
            prjTask->ActualFinish = varTime;
        }
        catch (...)
        {
        }		
	}
}
void                    CUpdatesManager::updateTask_ActualDuration(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask)
{
    if (pTaskReport->GetReportedActualDuration() >= 0)
    {
        try
        {
            prjTask->ActualDuration = pTaskReport->GetReportedActualDuration();
        }
        catch (...)
        {
        }       
    }
}

void                    CUpdatesManager::getProjectMapValues(MSProject::_IProjectDocPtr project)
{
    CString path = project->FullName;

    CString str = CTU_Utils::GetCustomPropertyValue(path, _T("PercComplReport"));
    m_iPercComplReport = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("PercComplDiff"));
    m_iPercComplDiff = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("PercComplGraph"));
    m_iPercComplGraphInd = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("PercWorkComplReport"));
    m_iPercWorkComplReport = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("PercWorkComplDiff"));
    m_iPercWorkComplDiff = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("PercWorkComplGraph"));
    m_iPercWorkComplGraphInd = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("ActWorkReport"));
    m_iActWorkReport = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("ActWorkDiff"));
    m_iActWorkDiff = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("ActWorkGraph"));
    m_iActWorkGraphInd = str.IsEmpty() ? -1 : _ttoi(str);

	str = CTU_Utils::GetCustomPropertyValue(path, _T("ActStartReport"));
	m_iActStartReport = str.IsEmpty() ? -1 : _ttoi(str);

	str = CTU_Utils::GetCustomPropertyValue(path, _T("ActStartDiff"));
	m_iActStartDiff = str.IsEmpty() ? -1 : _ttoi(str);

	str = CTU_Utils::GetCustomPropertyValue(path, _T("ActStartGraph"));
	m_iActStartGraphInd = str.IsEmpty() ? -1 : _ttoi(str);

	str = CTU_Utils::GetCustomPropertyValue(path, _T("ActFinishReport"));
	m_iActFinishReport = str.IsEmpty() ? -1 : _ttoi(str);

	str = CTU_Utils::GetCustomPropertyValue(path, _T("ActFinishDiff"));
	m_iActFinishDiff = str.IsEmpty() ? -1 : _ttoi(str);

	str = CTU_Utils::GetCustomPropertyValue(path, _T("ActFinishGraph"));
	m_iActFinishGraphInd = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("ActDurationReport"));
    m_iActDurationReport = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("ActDurationDiff"));
    m_iActDurationDiff = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("ActDurationGraph"));
    m_iActDurationGraphInd = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("CustomText1"));
    m_iCustomTextField1 = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("CustomText2"));
    m_iCustomTextField2 = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("CustomText3"));
    m_iCustomTextField3 = str.IsEmpty() ? -1 : _ttoi(str);
}
void					CUpdatesManager::processWritingToFiles(TUProjectsInfoList* pProjInfoList, bool& bCannotWriteManual, int reportingMode)
{
	TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
    {        
        if ((*itPrjInfo)->HasChilds())
        {
		    //for all sub projects
		    processWritingToFiles((*itPrjInfo)->GetSubProjectsList(), bCannotWriteManual, reportingMode);
        }

        if (reportingMode == TUTaskReportingMode)
            writeTaskReportsToFile(*itPrjInfo, bCannotWriteManual);
        else if (reportingMode == TUAssignmentReportingMode)
            writeAssgnReportsToFile(*itPrjInfo, bCannotWriteManual);
	}
}
void                    CUpdatesManager::getUpdatesViaGantt()
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    MSProject::_IProjectDocPtr project = app->GetActiveProject(); 

    getProjectMapValues(project);

    bool canProceed = true;
    TUTaskReportsList::iterator itTaskReport;
    for (itTaskReport = m_ProjectReport.GetTaskReportsList()->begin();
        itTaskReport != m_ProjectReport.GetTaskReportsList()->end();
        itTaskReport++)
    {
        if (TUTaskPercentCompleteReport == (*itTaskReport)->GetWorkInsertMode() && -1 == m_iPercComplReport)
        {
            canProceed = false;
            break;
        }
        if (TUTaskPercentWorkCompleteReport == (*itTaskReport)->GetWorkInsertMode() && -1 == m_iPercWorkComplReport)
        {
            canProceed = false;
            break;
        }
        if (TUTaskActualWorkReport == (*itTaskReport)->GetWorkInsertMode() && -1 == m_iActWorkReport)
        {
            canProceed = false;
            break;
        }
		if (TUTaskActualStartReport == (*itTaskReport)->GetWorkInsertMode() && -1 == m_iActStartReport)
		{
			canProceed = false;
			break;
		}
		if (TUTaskActualFinishReport == (*itTaskReport)->GetWorkInsertMode() && -1 == m_iActFinishReport)
		{
			canProceed = false;
			break;
		}
        if (TUTaskActualDurationReport == (*itTaskReport)->GetWorkInsertMode() && -1 == m_iActDurationReport)
        {
            canProceed = false;
            break;
        }
    }
    if (!canProceed)
    {
        MessageBox(NULL,
            _T("You don’t have specified custom fields for presenting the updates. Please go to Options dialog and in Mapping tab specify the Microsoft Project custom fields that will present the reported values."),
            _T("SeavusTaskUpdate"),
            MB_OK | MB_ICONWARNING);
        return;
    }

    processShowingUpdatesOnGantt(getProjInfoList());
    //save file

    if  (0 != app->FileSave())
    {

    }

    saveSPVFiles(getProjInfoList()); 
}
void                    CUpdatesManager::getNotesViaGantt(bool& success, bool& hasNoteReport)
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    MSProject::_IProjectDocPtr project = app->GetActiveProject(); 


    processShowingNotesOnGantt(getProjInfoList(), success, hasNoteReport);
    //save file

    if  (0 != app->FileSave())
    {

    }
}
void					CUpdatesManager::processShowingNotesOnGantt(TUProjectsInfoList* pProjInfoList, bool& success, bool& hasNoteReport)
{
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
    {        
        if ((*itPrjInfo)->HasChilds())
        {
            //for all sub projects
            processShowingNotesOnGantt((*itPrjInfo)->GetSubProjectsList(), success, hasNoteReport);
        }

        showNotesOnGantt((*itPrjInfo), success, hasNoteReport);
    }
}
void                    CUpdatesManager::showNotesOnGantt(CTUProjectInfo* pPrjInfo, bool& success, bool& hasNoteReport)
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    MSProject::_IProjectDocPtr project = app->GetActiveProject(); 

    MSProject::TaskPtr prjTask;

    // get proper file/subproject to make updates
    MSProject::_IProjectDocPtr currentProject;
    if (0 != wcscmp(pPrjInfo->GetProjectPath(), project->FullName))
    {
        currentProject = Utils::getSubProject(project, pPrjInfo->GetProjectPath());
        if (!currentProject)
        {
            return;
        }	
    }
    else
    {
        currentProject = project;
    }

    CTUProjectReport reportsToSave;
    getReportsForProject(pPrjInfo, &m_ProjectReport, reportsToSave, TUTaskAll);
    reportsToSave.RemoveBaseUID(pPrjInfo->GetProjectBaseUID());

    bool err = false;
    TUTaskReportsList::iterator itTaskReport;
    for(itTaskReport = reportsToSave.GetTaskReportsList()->begin();
        itTaskReport != reportsToSave.GetTaskReportsList()->end();
        itTaskReport++)
    {
        prjTask = taskByUID(currentProject, (int)(*itTaskReport)->GetTaskUID());
        if (prjTask)
        {
            showTask_Notes(*itTaskReport, prjTask, err, hasNoteReport);
        } 
    }

    saveSPVFile(pPrjInfo, true, false);

    if (err)
    {
        MessageBox(NULL,
            _T("Not all note updates are inserted in project plan."),
            _T("SeavusTaskUpdate"),
            MB_OK | MB_ICONWARNING);
    }
    else
    {
        success = true;
    }
}
void                    CUpdatesManager::getCustomTextUpdatesViaGantt(bool& success, bool& hasCustomtextReport)
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    MSProject::_IProjectDocPtr project = app->GetActiveProject(); 

    getProjectMapValues(project);

    processShowingCustomTextUpdatesOnGantt(getProjInfoList(), success, hasCustomtextReport);
    //save file

    if  (0 != app->FileSave())
    {

    }
}
void					CUpdatesManager::processShowingUpdatesOnGantt(TUProjectsInfoList* pProjInfoList)
{
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
    {        
        if ((*itPrjInfo)->HasChilds())
        {
            //for all sub projects
            processShowingUpdatesOnGantt((*itPrjInfo)->GetSubProjectsList());
        }

        showTaskUpdatesOnGantt((*itPrjInfo));
    }
}
void					CUpdatesManager::processShowingCustomTextUpdatesOnGantt(TUProjectsInfoList* pProjInfoList, bool& success, bool& hasCustomTextReport)
{
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjInfoList->begin(); itPrjInfo != pProjInfoList->end(); ++itPrjInfo)
    {        
        if ((*itPrjInfo)->HasChilds())
        {
            //for all sub projects
            processShowingCustomTextUpdatesOnGantt((*itPrjInfo)->GetSubProjectsList(), success, hasCustomTextReport);
        }

        showCustomTextUpdatesOnGantt((*itPrjInfo), success, hasCustomTextReport);
    }
}

void					CUpdatesManager::writeTaskReportsToFile(CTUProjectInfo* pPrjInfo, bool& bCannotWriteManual)
{
	MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
	MSProject::_IProjectDocPtr project = app->GetActiveProject(); 
	MSProject::TaskPtr prjTask;

    // get proper file/subproject to make updates
    MSProject::_IProjectDocPtr currentProject = NULL;
     if (0 != wcscmp(pPrjInfo->GetProjectPath(), project->FullName))
     {
         currentProject = Utils::getSubProject(project, pPrjInfo->GetProjectPath());
         if (!currentProject)
         {
             return;
         }	
     }
     else
     {
         currentProject = project;
     }

	CTUProjectReport reportsToSave;
	getReportsForProject(pPrjInfo, &m_ProjectReport, reportsToSave, TUTaskToBeAccepted);
	reportsToSave.RemoveBaseUID(pPrjInfo->GetProjectBaseUID());

    _bstr_t appVersion = app->Version;
    int iMajorVersion = Utils::GetProjectMajorVersion((BSTR)appVersion);

    _bstr_t appBuild = "0";
    int iMinorVersion = 0;
    if (iMajorVersion == 14)
    {
        appBuild = app->Build;
        iMinorVersion = Utils::GetProjectMinorVersion((BSTR)appBuild);
    }

    bool bProceededWithManual = (iMajorVersion == 14 && iMinorVersion >= 6023) || iMajorVersion < 14;

    TUTaskReportsList::iterator itTaskReport;
    for(itTaskReport = reportsToSave.GetTaskReportsList()->begin();
        itTaskReport != reportsToSave.GetTaskReportsList()->end();
        itTaskReport++)
    {
        prjTask = taskByUID(currentProject, (int)(*itTaskReport)->GetTaskUID());
        if (prjTask)
        {
            BOOL isManual = 0;
            if(iMajorVersion == 14)
            {
                isManual = prjTask->Manual; //0 is for auto, -1 for manual
            }

            if(isManual == 0 || bProceededWithManual)
            {
                switch ((*itTaskReport)->GetWorkInsertMode())
                {
                case TUTaskPercentCompleteReport:
                    updateTask_PercentComplete(*itTaskReport, prjTask);
                	break;
                case TUTaskPercentWorkCompleteReport:
                    updateTask_PercentWorkComplete(*itTaskReport, prjTask);
                    break;
                case TUTaskActualWorkReport:
                    updateTask_ActualWork(*itTaskReport, prjTask);
                    break;
				case TUTaskActualStartReport:
					updateTask_ActualStart(*itTaskReport, prjTask);
					break;
				case TUTaskActualFinishReport:
					updateTask_ActualFinish(*itTaskReport, prjTask);
					break;
                case TUTaskActualDurationReport:
                    updateTask_ActualDuration(*itTaskReport, prjTask);
                    break;
                }
            }
            else
            {
                CTUTaskReport* pManualTaskReport = m_ProjectReport.GetTaskReport((*itTaskReport)->GetTaskUID() + pPrjInfo->GetProjectBaseUID());
                if (pManualTaskReport)
                {
                    pManualTaskReport->SetStatus(TUTaskOpen);
                }
                bCannotWriteManual = true;
            }
        } 
    }
}
void					CUpdatesManager::writeAssgnReportsToFile(CTUProjectInfo* pPrjInfo, bool& bCannotWriteManual)
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    MSProject::_IProjectDocPtr project = app->GetActiveProject(); 
    MSProject::TaskPtr prjTask;

    // get proper file/subproject to make updates
    MSProject::_IProjectDocPtr currentProject;
    if (0 != wcscmp(pPrjInfo->GetProjectPath(), project->FullName))
    {
        currentProject = Utils::getSubProject(project, pPrjInfo->GetProjectPath());
        if (!currentProject)
        {
            return;
        }	
    }	
    else
    {
        currentProject = project;
    }

    CTUProjectReport reportsToSave;
    getReportsForProject(pPrjInfo, &m_ProjectReport, reportsToSave, TUToBeAccepted);
    reportsToSave.RemoveBaseUID(pPrjInfo->GetProjectBaseUID());

    _bstr_t appVersion = app->Version;
    int iMajorVersion = Utils::GetProjectMajorVersion((BSTR)appVersion);

    _bstr_t appBuild = "0";
    int iMinorVersion = 0;
    if (iMajorVersion == 14)
    {
        appBuild = app->Build;
        iMinorVersion = Utils::GetProjectMinorVersion((BSTR)appBuild);
    }

    bool bProceededWithManual = (iMajorVersion == 14 && iMinorVersion >= 6023) || iMajorVersion < 14;

    TUAssignmentReportsList::iterator itAssgnReport;
    for(itAssgnReport = reportsToSave.GetAssignmentReportsList()->begin();
        itAssgnReport != reportsToSave.GetAssignmentReportsList()->end();
        itAssgnReport++)
    {
        // check for insert mode (day report or by percent)
        if ((*itAssgnReport)->GetWorkInsertMode() == TUAssgnTimeReports)
        {
            prjTask = taskByUID(currentProject, (int)(*itAssgnReport)->GetTaskUID());

            if (prjTask)
            {
                BOOL isManual = 0;
                if(iMajorVersion == 14)
                {
                    isManual = prjTask->Manual; //0 is for auto, -1 for manual
                }

                if(isManual == 0 || bProceededWithManual)
                {
                    int tempAssigmCount = prjTask->Assignments->GetCount();
                    if(!tempAssigmCount)
                    {
                        updateWorkByTask(project, *itAssgnReport, prjTask);
                    }
                    else
                    {
                        updateWorkByAssigment(*itAssgnReport, prjTask);
                    }             
                }
                else
                {
                    CTUAssignmentReport* pManualAssgnReport = m_ProjectReport.GetAssgnReport((*itAssgnReport)->GetAssignmentUID() + pPrjInfo->GetProjectBaseUID());
                    if (pManualAssgnReport)
                    {
                        pManualAssgnReport->SetAllReportsStatus(TUOpen);
                    }
                    bCannotWriteManual = true;
                }
            }
        }
        else if ((*itAssgnReport)->GetWorkInsertMode() == TUAssgnPercentReports)
        {
            prjTask = taskByUID(currentProject, (int)(*itAssgnReport)->GetTaskUID());

            if (prjTask)
            {
                BOOL isManual = 0;
                if(iMajorVersion == 14)
                {
                    isManual = prjTask->Manual; //0 is for auto, -1 for manual
                }

                if(isManual == 0 || bProceededWithManual)
                {
                    int tempAssigmCount = prjTask->Assignments->GetCount();
                    if(!tempAssigmCount)
                    {
                        updatePercent_TaskWithUnassignedResources(*itAssgnReport, prjTask);
                    }
                    else
                    {
                        updatePercent_TaskWithAssigment(*itAssgnReport, prjTask);
                    }
                }
                else
                {
                    CTUAssignmentReport* pManualAssgnReport = m_ProjectReport.GetAssgnReport((*itAssgnReport)->GetAssignmentUID() + pPrjInfo->GetProjectBaseUID());
                    if (pManualAssgnReport)
                    {
                        pManualAssgnReport->SetAllReportsStatus(TUOpen);
                    }
                    bCannotWriteManual = true;
                }
            } 
        } 
    }
}
void                    CUpdatesManager::showTaskUpdatesOnGantt(CTUProjectInfo* pPrjInfo)
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    MSProject::_IProjectDocPtr project = app->GetActiveProject(); 

    MSProject::TaskPtr prjTask;

    // get proper file/subproject to make updates
    MSProject::_IProjectDocPtr currentProject;
    if (0 != wcscmp(pPrjInfo->GetProjectPath(), project->FullName))
    {
        currentProject = Utils::getSubProject(project, pPrjInfo->GetProjectPath());
        if (!currentProject)
        {
            return;
        }	
    }
    else
    {
        currentProject = project;
    }
   
    m_ProjectReport.ChangeAllReportsStatus(TUTaskAll, TUTaskAccepted);

    CTUProjectReport reportsToSave;
    getReportsForProject(pPrjInfo, &m_ProjectReport, reportsToSave, TUTaskAll);
    reportsToSave.RemoveBaseUID(pPrjInfo->GetProjectBaseUID());

    bool err = false;
    TUTaskReportsList::iterator itTaskReport;
    for(itTaskReport = reportsToSave.GetTaskReportsList()->begin();
        itTaskReport != reportsToSave.GetTaskReportsList()->end();
        itTaskReport++)
    {
        prjTask = taskByUID(currentProject, (int)(*itTaskReport)->GetTaskUID());
        if (prjTask)
        {
            switch ((*itTaskReport)->GetWorkInsertMode())
            {
            case TUTaskPercentCompleteReport:
                showTask_PercentCompleteReport(*itTaskReport, prjTask, err);
                showTask_PercentCompleteDiff(*itTaskReport, prjTask, err);
                showTask_PercentCompleteGraphInd(*itTaskReport, prjTask, err);
                break;
            case TUTaskPercentWorkCompleteReport:
                showTask_PercentWorkCompleteReport(*itTaskReport, prjTask, err);
                showTask_PercentWorkCompleteDiff(*itTaskReport, prjTask, err);
                showTask_PercentWorkCompleteGraphInd(*itTaskReport, prjTask, err);
                break;
            case TUTaskActualWorkReport:
                showTask_ActualWorkReport(*itTaskReport, prjTask, err);
                showTask_ActualWorkDiff(*itTaskReport, prjTask, err);
                showTask_ActualWorkGraphInd(*itTaskReport, prjTask, err);
                break;
			case TUTaskActualStartReport:
				showTask_ActualStartReport(*itTaskReport, prjTask, err);
				showTask_ActualStartDiff(*itTaskReport, prjTask, err);
				showTask_ActualStartGraphInd(*itTaskReport, prjTask, err);
				break;
			case TUTaskActualFinishReport:
				showTask_ActualFinishReport(*itTaskReport, prjTask, err);
				showTask_ActualFinishDiff(*itTaskReport, prjTask, err);
				showTask_ActualFinishGraphInd(*itTaskReport, prjTask, err);
				break;
            case TUTaskActualDurationReport:
                showTask_ActualDurationReport(*itTaskReport, prjTask, err);
                showTask_ActualDurationDiff(*itTaskReport, prjTask, err);
                showTask_ActualDurationGraphInd(*itTaskReport, prjTask, err);
                break;
            }
        } 
    }

    if (err)
    {
        MessageBox(NULL,
            _T("Not all updates are inserted in project plan due to conflicts in the definition of the custom fields. Please, review:\n- The mapping process of the custom fields\n- Received updates."),
            _T("SeavusTaskUpdate"),
            MB_OK | MB_ICONWARNING);   
    }
}
void                    CUpdatesManager::showCustomTextUpdatesOnGantt(CTUProjectInfo* pPrjInfo, bool& success, bool& hasCustomTextReport)
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    MSProject::_IProjectDocPtr project = app->GetActiveProject(); 

    MSProject::TaskPtr prjTask;

    // get proper file/subproject to make updates
    MSProject::_IProjectDocPtr currentProject;
    if (0 != wcscmp(pPrjInfo->GetProjectPath(), project->FullName))
    {
        currentProject = Utils::getSubProject(project, pPrjInfo->GetProjectPath());
        if (!currentProject)
        {
            return;
        }	
    }
    else
    {
        currentProject = project;
    }

    CTUProjectReport reportsToSave;
    getReportsForProject(pPrjInfo, &m_ProjectReport, reportsToSave, TUTaskAll);
    reportsToSave.RemoveBaseUID(pPrjInfo->GetProjectBaseUID());

    bool err = false;
    TUTaskReportsList::iterator itTaskReport;
    for(itTaskReport = reportsToSave.GetTaskReportsList()->begin();
        itTaskReport != reportsToSave.GetTaskReportsList()->end();
        itTaskReport++)
    {
        prjTask = taskByUID(currentProject, (int)(*itTaskReport)->GetTaskUID());
        if (prjTask)
        {
            showTask_CustomText1Field(*itTaskReport, prjTask, err, hasCustomTextReport);
            showTask_CustomText2Field(*itTaskReport, prjTask, err, hasCustomTextReport);
            showTask_CustomText3Field(*itTaskReport, prjTask, err, hasCustomTextReport);
        } 
    }

    saveSPVFile(pPrjInfo, false, true);

    if (err)
    {
        MessageBox(NULL,
            _T("Not all updates are inserted in project plan due to conflicts in the definition of the custom fields. Please, review:\n- The mapping process of the custom fields\n- Received updates."),
            _T("SeavusTaskUpdate"),
            MB_OK | MB_ICONWARNING);
    }
    else
    {
        success = true;
    }
}

void                    CUpdatesManager::showTask_PercentCompleteReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iPercComplReport && pTaskReport->GetReportedPercentComplete() >= 0)
    {
        Utils::setCustomField(m_iPercComplReport, ePercComplReport);

        CString strVal;
        strVal.Format(_T("%d%%"), pTaskReport->GetReportedPercentComplete());
        _bstr_t val = strVal;

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iPercComplReport, val);
        }
        catch (...)
        {
            err = true;
        }
    }
}
void                    CUpdatesManager::showTask_PercentWorkCompleteReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iPercWorkComplReport && pTaskReport->GetReportedPercentWorkComplete() >= 0)
    {
        Utils::setCustomField(m_iPercWorkComplReport, ePercWorkComplReport);

        CString strVal;
        strVal.Format(_T("%d%%"), pTaskReport->GetReportedPercentWorkComplete());
        _bstr_t val = strVal;

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iPercWorkComplReport, val);
        }
        catch (...)
        {
            err = true;
        } 
    }
}
void                    CUpdatesManager::showTask_ActualWorkReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iActWorkReport)
    {
        Utils::setCustomField(m_iActWorkReport, eActWorkReport);

        int minutesVal = pTaskReport->GetReportedActualWork();
        CString strVal = Utils::FormatWork(minutesVal);

        _bstr_t val = strVal;

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iActWorkReport, val);
        }
        catch (...)
        {
            err = true;
        }    
    }
}
void                    CUpdatesManager::showTask_ActualStartReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
	if (-1 != m_iActStartReport)
	{
		Utils::setCustomField(m_iActStartReport, eActStartReport);

		CTime timeVal = Utils::toGMT(pTaskReport->GetReportedActualStart());
        CString strDateTime = Utils::get_date_time_in_user_format(timeVal);
        _bstr_t bsrtVal = strDateTime;

        try 
        {
			prjTask->SetField((MSProject::PjField)m_iActStartReport, bsrtVal);
		}
		catch (...)
		{
			err = true;
		}    
	}
}
void                    CUpdatesManager::showTask_ActualFinishReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{

	if (-1 != m_iActFinishReport)
	{
		Utils::setCustomField(m_iActFinishReport, eActFinishReport);

		CTime timeVal = Utils::toGMT(pTaskReport->GetReportedActualFinish());
		CString strDateTime = Utils::get_date_time_in_user_format(timeVal);
		_bstr_t bsrtVal = strDateTime;

		try 
		{
			prjTask->SetField((MSProject::PjField)m_iActFinishReport, bsrtVal);
		}
		catch (...)
		{
			err = true;
		}    
	}
}
void                    CUpdatesManager::showTask_ActualDurationReport(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iActDurationReport)
    {
        Utils::setCustomField(m_iActDurationReport, eActDurationReport);

        int minutesVal = pTaskReport->GetReportedActualDuration();
        CString strVal = Utils::FormatWork(minutesVal);

        _bstr_t val = strVal;

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iActDurationReport, val);
        }
        catch (...)
        {
            err = true;
        }    
    }
}
void                    CUpdatesManager::showTask_PercentCompleteDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iPercComplDiff && pTaskReport->GetReportedPercentComplete() >= 0)
    {
        Utils::setCustomField(m_iPercComplDiff, ePercComplDiff);

        int diff = pTaskReport->GetReportedPercentComplete() - (int)prjTask->PercentComplete;
        CString strVal;
        strVal.Format(_T("%d%%"), diff);
        _bstr_t val = strVal;

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iPercComplDiff, val);
        }
        catch (...)
        {
            err = true;
        }
    }
}
void                    CUpdatesManager::showTask_PercentWorkCompleteDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iPercWorkComplDiff && pTaskReport->GetReportedPercentWorkComplete() >= 0)
    {
        Utils::setCustomField(m_iPercWorkComplDiff, ePercWorkComplDiff);

        int diff = pTaskReport->GetReportedPercentWorkComplete() - (int)prjTask->PercentWorkComplete;
        CString strVal;
        strVal.Format(_T("%d%%"), diff);
        _bstr_t val = strVal;

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iPercWorkComplDiff, val);
        }
        catch (...)
        {
            err = true;
        }  
    }
}
void                    CUpdatesManager::showTask_ActualWorkDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iActWorkDiff)
    {
        Utils::setCustomField(m_iActWorkDiff, eActWorkDiff);

        int minutesVal = pTaskReport->GetReportedActualWork();
        int diff = minutesVal - (int)prjTask->ActualWork;
        CString strVal = Utils::FormatWork(diff);

        _bstr_t val = strVal;

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iActWorkDiff, val);
        }
        catch (...)
        {
            err = true;
        }    
    }
}
void                    CUpdatesManager::showTask_ActualStartDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iActStartDiff)
    {
		MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
		MSProject::_IProjectDocPtr project = app->GetActiveProject(); 
		double hoursPerDay = project->GetHoursPerDay();

        Utils::setCustomField(m_iActStartDiff, eActStartDiff);

		COleDateTime odt_val;
		_bstr_t odt_bstrVal = prjTask->ActualStart;
		odt_val.ParseDateTime(odt_bstrVal, LOCALE_NOUSEROVERRIDE,LANG_USER_DEFAULT); 

		CTime time_start;
		time_start = (odt_val.m_status == COleDateTime::invalid) 
			? CTime(0) 
			: CTime(odt_val.GetYear(), odt_val.GetMonth(), odt_val.GetDay(), odt_val.GetHour(), odt_val.GetMinute(), 0);

        CTime val = Utils::toGMT(pTaskReport->GetReportedActualStart());

        CTimeSpan diff = val - time_start;
        double dVal = ((diff.GetTotalHours() / (24 / hoursPerDay))) / hoursPerDay;

		CString strVal;
		strVal.Format(_T("%.2f days"), dVal);

		if (time_start == CTime(0) || val == CTime(0))
		{
			strVal = _T("NA");
		}

        _bstr_t bstrVal = strVal;

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iActStartDiff, bstrVal);
        }
        catch (...)
        {
            err = true;
        }    
    }
}
void                    CUpdatesManager::showTask_ActualFinishDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
	if (-1 != m_iActFinishDiff)
	{
		MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
		MSProject::_IProjectDocPtr project = app->GetActiveProject(); 
		double hoursPerDay = project->GetHoursPerDay();

		Utils::setCustomField(m_iActFinishDiff, eActFinishDiff);

		COleDateTime odt_val;
		_bstr_t odt_bstrVal = prjTask->ActualFinish;
		odt_val.ParseDateTime(odt_bstrVal, LOCALE_NOUSEROVERRIDE,LANG_USER_DEFAULT); 

		CTime time_finish;
		time_finish = (odt_val.m_status == COleDateTime::invalid) 
			? CTime(0) 
			: CTime(odt_val.GetYear(), odt_val.GetMonth(), odt_val.GetDay(), odt_val.GetHour(), odt_val.GetMinute(), 0);

		CTime val = Utils::toGMT(pTaskReport->GetReportedActualFinish());
		
        CTimeSpan diff = val - time_finish;
        double dVal = ((diff.GetTotalHours() / (24 / hoursPerDay))) / hoursPerDay;

		CString strVal;
		strVal.Format(_T("%.2f days"), dVal);

		if (time_finish == CTime(0) || val == CTime(0))
		{
			strVal = _T("NA");
		}

		_bstr_t bstrVal = strVal;

		try 
		{
			prjTask->SetField((MSProject::PjField)m_iActFinishDiff, bstrVal);
		}
		catch (...)
		{
			err = true;
		}    
	}
}
void                    CUpdatesManager::showTask_ActualDurationDiff(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iActDurationDiff)
    {
        Utils::setCustomField(m_iActDurationDiff, eActDurationDiff);

        int minutesVal = pTaskReport->GetReportedActualDuration();
        int diff = minutesVal - (int)prjTask->ActualDuration;
        CString strVal = Utils::FormatWork(diff);

        _bstr_t val = strVal;

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iActDurationDiff, val);
        }
        catch (...)
        {
            err = true;
        }    
    }
}
void                    CUpdatesManager::showTask_PercentCompleteGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iPercComplGraphInd && pTaskReport->GetReportedPercentComplete() >= 0)
    {
        Utils::setCustomField(m_iPercComplGraphInd, ePercComplGraphInd);

        int diff = pTaskReport->GetReportedPercentComplete() - (int)prjTask->PercentComplete;
        BSTR val = diff == 0 ? L"No" : L"Yes";

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iPercComplGraphInd, val);
        }
        catch (...)
        {
            err = true;
        }   
    }
}
void                    CUpdatesManager::showTask_PercentWorkCompleteGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iPercWorkComplGraphInd && pTaskReport->GetReportedPercentWorkComplete() >= 0)
    {
        Utils::setCustomField(m_iPercWorkComplGraphInd, ePercWorkComplGraphInd);

        int diff = pTaskReport->GetReportedPercentWorkComplete() - (int)prjTask->PercentWorkComplete;
        BSTR val = diff == 0 ? L"No" : L"Yes";

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iPercWorkComplGraphInd, val);
        }
        catch (...)
        {
            err = true;
        } 
    }
}
void                    CUpdatesManager::showTask_ActualWorkGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iActWorkGraphInd)
    {
        Utils::setCustomField(m_iActWorkGraphInd, eActWorkGraphInd);

        int minutesVal = pTaskReport->GetReportedActualWork();
        int diff = minutesVal - (int)prjTask->ActualWork;
        bool bEqual  = (diff <= 1 && diff >= -1);
        BSTR val = bEqual ? L"No" : L"Yes";

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iActWorkGraphInd, val);
        }
        catch (...)
        {
            err = true;
        } 
    }
}
void                    CUpdatesManager::showTask_ActualStartGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
	if (-1 != m_iActStartGraphInd)
	{
		Utils::setCustomField(m_iActStartGraphInd, eActStartGraphInd);

		COleDateTime odt_val;
		_bstr_t odt_bstrVal = prjTask->ActualStart;
		odt_val.ParseDateTime(odt_bstrVal, LOCALE_NOUSEROVERRIDE,LANG_USER_DEFAULT); 

		CTime time_start;
		time_start = (odt_val.m_status == COleDateTime::invalid) 
			? CTime(0) 
			: CTime(odt_val.GetYear(), odt_val.GetMonth(), odt_val.GetDay(), odt_val.GetHour(), odt_val.GetMinute(), 0);

		CTime val = Utils::toGMT(pTaskReport->GetReportedActualStart());
		CTimeSpan diff = val - time_start;

		bool bEqual  = (diff <= 1 && diff >= -1);
		BSTR bstrVal = bEqual ? L"No" : L"Yes";

		try 
		{
			prjTask->SetField((MSProject::PjField)m_iActStartGraphInd, bstrVal);
		}
		catch (...)
		{
			err = true;
		} 
	}
}
void                    CUpdatesManager::showTask_ActualFinishGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
	if (-1 != m_iActFinishGraphInd)
	{
		Utils::setCustomField(m_iActFinishGraphInd, eActFinishGraphInd);

		COleDateTime odt_val;
		_bstr_t odt_bstrVal = prjTask->ActualFinish;
		odt_val.ParseDateTime(odt_bstrVal, LOCALE_NOUSEROVERRIDE,LANG_USER_DEFAULT); 

		CTime time_finish;
		time_finish = (odt_val.m_status == COleDateTime::invalid) 
			? CTime(0) 
			: CTime(odt_val.GetYear(), odt_val.GetMonth(), odt_val.GetDay(), odt_val.GetHour(), odt_val.GetMinute(), 0);

		CTime val = Utils::toGMT(pTaskReport->GetReportedActualFinish());
		CTimeSpan diff = val - time_finish;

		bool bEqual  = (diff <= 1 && diff >= -1);
		BSTR bstrVal = bEqual ? L"No" : L"Yes";

		try 
		{
			prjTask->SetField((MSProject::PjField)m_iActFinishGraphInd, bstrVal);
		}
		catch (...)
		{
			err = true;
		} 
	}
}
void                    CUpdatesManager::showTask_ActualDurationGraphInd(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err)
{
    if (-1 != m_iActDurationGraphInd)
    {
        Utils::setCustomField(m_iActDurationGraphInd, eActDurationGraphInd);

        int minutesVal = pTaskReport->GetReportedActualDuration();
        int diff = minutesVal - (int)prjTask->ActualDuration;
        bool bEqual  = (diff <= 1 && diff >= -1);
        BSTR val = bEqual ? L"No" : L"Yes";

        try 
        {
            prjTask->SetField((MSProject::PjField)m_iActDurationGraphInd, val);
        }
        catch (...)
        {
            err = true;
        } 
    }
}
void                    CUpdatesManager::showTask_Notes(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err, bool& hasNoteReport)
{
    CString val = pTaskReport->GetReportedNotes();
    if (!val.IsEmpty())
    {
        hasNoteReport = true;

        try
        {
            CString str = (LPCTSTR)prjTask->Notes;
            if (!str.IsEmpty())
            {
                str.Append(_T("\n"));
                str.Append(_T("\n"));
            }
            str.Append(val);
            prjTask->Notes = str.AllocSysString();

            pTaskReport->SetReportedNotes(_T(""));
            pTaskReport->SetCurrentNotes(_T(""));
        }
        catch (...)
        {
            err = true;
        }       
    }
}
void                    CUpdatesManager::showTask_CustomText1Field(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err, bool& hasCustomTextReport)
{
    CString val = pTaskReport->GetCustomText1FieldText();
    if (!val.IsEmpty())
    {
        hasCustomTextReport = true;

        if (-1 != m_iCustomTextField1)
        {
            Utils::setCustomField(m_iCustomTextField1, eCustomText1);

            try 
            {
                prjTask->SetField((MSProject::PjField)m_iCustomTextField1, val.AllocSysString());
                pTaskReport->SetCustomText1FieldText(_T(""));
            }
            catch (...)
            {
                err = true;
            }
        }
    }
}
void                    CUpdatesManager::showTask_CustomText2Field(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err, bool& hasCustomTextReport)
{
    CString val = pTaskReport->GetCustomText2FieldText();
    if (!val.IsEmpty())
    {
        hasCustomTextReport = true;

        if (-1 != m_iCustomTextField2)
        {
            Utils::setCustomField(m_iCustomTextField2, eCustomText2);

            try 
            {
                prjTask->SetField((MSProject::PjField)m_iCustomTextField2, val.AllocSysString());
                pTaskReport->SetCustomText2FieldText(_T(""));
            }
            catch (...)
            {
                err = true;
            }
        }
    }
}
void                    CUpdatesManager::showTask_CustomText3Field(CTUTaskReport* pTaskReport, MSProject::TaskPtr prjTask, bool& err, bool& hasCustomTextReport)
{
    CString val = pTaskReport->GetCustomText3FieldText();
    if (!val.IsEmpty())
    {
        hasCustomTextReport = true;

        if (-1 != m_iCustomTextField3)
        {
            Utils::setCustomField(m_iCustomTextField3, eCustomText3);

            try 
            {
                prjTask->SetField((MSProject::PjField)m_iCustomTextField3, val.AllocSysString());
                pTaskReport->SetCustomText2FieldText(_T(""));
            }
            catch (...)
            {
                err = true;
            }      
        }
    }
}
void					CUpdatesManager::reopenProject()
{
	TUProjectsInfoList::iterator itRoot = m_ProjInfoList.begin();
	CTUProjectInfo* pRoot = *itRoot;
	if(pRoot)
	{
		_bstr_t name = (LPCTSTR)pRoot->GetProjectPath();

		m_app->FileOpen(name,vtMissing,vtMissing,vtMissing,
			vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,vtMissing,MSProject::pjDoNotOpenPool,
			vtMissing,vtMissing,vtMissing,vtMissing);
	}
}

void CUpdatesManager::removeAcceptedNotes(CTUProjectReport& projectReport)
{
    TUTaskReportsList::iterator itTaskReport;
    for(itTaskReport = projectReport.GetTaskReportsList()->begin();
        itTaskReport != projectReport.GetTaskReportsList()->end();
        itTaskReport++)
    {
        (*itTaskReport)->SetCurrentNotes(_T(""));
        (*itTaskReport)->SetReportedNotes(_T(""));
    }
}

void CUpdatesManager::removeAcceptedCustomFields(CTUProjectReport& projectReport)
{
    TUTaskReportsList::iterator itTaskReport;
    for(itTaskReport = projectReport.GetTaskReportsList()->begin();
        itTaskReport != projectReport.GetTaskReportsList()->end();
        itTaskReport++)
    {
        (*itTaskReport)->SetCustomText1FieldText(_T(""));
        (*itTaskReport)->SetCustomText2FieldText(_T(""));
        (*itTaskReport)->SetCustomText3FieldText(_T(""));
    }
}

void CUpdatesManager::removeUndefinedEmptyReports(CTUProjectReport& projectReport)
{
    TUTaskReportsList::iterator itTaskReport;
    for(itTaskReport = projectReport.GetTaskReportsList()->begin();
        itTaskReport != projectReport.GetTaskReportsList()->end();)
    {
        if ((*itTaskReport)->IsUndefined())
        {
            if ((*itTaskReport)->GetCustomText1FieldText().IsEmpty() &&
                (*itTaskReport)->GetCustomText2FieldText().IsEmpty() &&
                (*itTaskReport)->GetCustomText3FieldText().IsEmpty() &&
                (*itTaskReport)->GetReportedNotes().IsEmpty())
            {
                delete *itTaskReport;
                *itTaskReport = NULL;
                itTaskReport = projectReport.GetTaskReportsList()->erase(itTaskReport);
            }
            else
            {
                itTaskReport++;
            }
        }
        else
        {
            itTaskReport++;
        }
    }
}

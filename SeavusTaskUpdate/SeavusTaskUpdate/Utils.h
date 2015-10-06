#pragma once
#include "TaskUpdate_Utils.h"
#include <map>

enum enumCustomFields
{
    ePercComplReport,      
    ePercWorkComplReport,  
    eActWorkReport,        
	eActStartReport,       
	eActFinishReport,      
    eActDurationReport,    
    ePercComplDiff,        
    ePercWorkComplDiff,    
    eActWorkDiff,          
	eActStartDiff,         
	eActFinishDiff,     
    eActDurationDiff, 
    ePercComplGraphInd,
    ePercWorkComplGraphInd,
    eActWorkGraphInd,
	eActStartGraphInd,
	eActFinishGraphInd,
    eActDurationGraphInd,
    eCustomText1,
    eCustomText2,
    eCustomText3
};
typedef std::map<short, CString> LabelList;

class Utils
{
public:
    Utils(void);
    ~Utils(void);

    static BOOL isProjectSaved(MSProject::_IProjectDocPtr project);
    static bool areSubProjectsSaved(MSProject::_IProjectDocPtr project);    
    static BOOL saveProject(MSProject::_IProjectDocPtr project);
    static void initWorkUnits();
    static CString FormatWork(int valMinutes);
    static int GetProjectMinorVersion(CString strBuild);
    static int GetProjectMajorVersion(CString strBuild);
    static void initAppUnits();
    static void setCustomField(int fieldID, enumCustomFields eCustomField);
    static bool isTaskUpdateCustomField(int fieldID);
    static void initCustomFieldNames();

    static bool isProjectCollaborative(CString path, CTUProjectInfo* pPrjInfo);
    static bool hasCollaborativeChild(MSProject::_IProjectDocPtr prj, CTUProjectInfo* pPrjInfo);
    static void closeCollaborativeProjects(MSProject::_MSProjectPtr app, CTUProjectInfo* pPrjInfo);

    static MSProject::AssignmentPtr  getPjAssignment(MSProject::_IProjectDocPtr project, MSProject::TaskPtr pTask, long assignmentUID);
    static MSProject::_IProjectDocPtr getSubProject(MSProject::_IProjectDocPtr project, CString path);
	static CTime toGMT(CTime t);
public:
    static LabelList m_minuteLabels;
    static LabelList m_hourLabels;
    static LabelList m_dayLabels;
    static LabelList m_weekLabels;
    static LabelList m_monthLabels;

    static short m_minuteLabelDisplay;
    static short m_hourLabelDisplay;
    static short m_dayLabelDisplay;
    static short m_weekLabelDisplay;
    static short m_monthLabelDisplay;

    static double m_hoursPerDay ;
    static double m_hoursPerWeek;
    static double m_daysPerMonth;

    static MSProject::PjUnit m_defWorkUnit;

    static std::map<enumCustomFields, CString> m_customFieldNames;
    static CString get_date_in_user_format (CTime& time);
    static CString get_time_in_user_format (CTime& time);
    static CString get_date_time_in_user_format(CTime& time);
	static int GetYear(CTime time);
	static int GetMonth(CTime time);
	static int GetDay(CTime time);
	static CString formatMSTime(CTime time);
};


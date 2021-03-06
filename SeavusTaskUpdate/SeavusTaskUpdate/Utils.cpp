#include "StdAfx.h"
#include "Utils.h"
#include <atlpath.h>

LabelList Utils::m_minuteLabels;
LabelList Utils::m_hourLabels;
LabelList Utils::m_dayLabels;
LabelList Utils::m_weekLabels;
LabelList Utils::m_monthLabels;

short Utils::m_minuteLabelDisplay = 1;
short Utils::m_hourLabelDisplay = 1;
short Utils::m_dayLabelDisplay = 2;
short Utils::m_weekLabelDisplay = 1;
short Utils::m_monthLabelDisplay = 1;

double Utils::m_hoursPerDay = 8;
double Utils::m_hoursPerWeek = 40;
double Utils::m_daysPerMonth = 20;

MSProject::PjUnit Utils::m_defWorkUnit = MSProject::pjHour;

std::map<enumCustomFields, CString> Utils::m_customFieldNames;

void Utils::initCustomFieldNames()
{
    m_customFieldNames[ePercComplReport] = _T("Reported % Complete");
    m_customFieldNames[ePercWorkComplReport] = _T("Reported % Work Complete");
    m_customFieldNames[eActWorkReport] = _T("Reported Actual Work");
	m_customFieldNames[eActStartReport] = _T("Reported Actual Start");
	m_customFieldNames[eActFinishReport] = _T("Reported Actual Finish");
    m_customFieldNames[eActDurationReport] = _T("Reported Actual Duration");
    m_customFieldNames[ePercComplDiff] = _T("% Complete Difference");
    m_customFieldNames[ePercWorkComplDiff] = _T("% Work Complete Difference");
    m_customFieldNames[eActWorkDiff] = _T("Actual Work Difference");
	m_customFieldNames[eActStartDiff] = _T("Actual Start Difference");
	m_customFieldNames[eActFinishDiff] = _T("Actual Finish Difference");
    m_customFieldNames[eActDurationDiff] = _T("Actual Duration Difference");
    m_customFieldNames[ePercComplGraphInd] = _T("% Complete Graph. Ind.");
    m_customFieldNames[ePercWorkComplGraphInd] = _T("% Work Complete Graph. Ind.");
    m_customFieldNames[eActWorkGraphInd] = _T("Actual Work Graph. Ind.");
	m_customFieldNames[eActStartGraphInd] = _T("Actual Start Graph. Ind.");
	m_customFieldNames[eActFinishGraphInd] = _T("Actual Finish Graph. Ind.");
    m_customFieldNames[eActDurationGraphInd] = _T("Actual Duration Graph. Ind.");
    m_customFieldNames[eCustomText1] = _T("1st Custom Text");
    m_customFieldNames[eCustomText2] = _T("2nd Custom Text");
    m_customFieldNames[eCustomText3] = _T("3rd Custom Text");
}
BOOL Utils::isProjectSaved(MSProject::_IProjectDocPtr project)
{
    BOOL returnValue = true;

    if(returnValue = ATLPath::FileExists((BSTR)project->FullName))
    {
        returnValue = project->Saved 
            && areSubProjectsSaved(project);
    }

    return returnValue;
}

bool Utils::areSubProjectsSaved(MSProject::_IProjectDocPtr project)
{
    bool returnValue = true;

    MSProject::SubprojectsPtr subProjectsList = project->Subprojects;
    for(int i = 1; i <= subProjectsList->GetCount(); i++)
    {
        if (MSProject::SubprojectPtr subProject = subProjectsList->Item[i])
        {         
            MSProject::_IProjectDocPtr subProjectDoc = subProject->GetSourceProject();
            returnValue = returnValue &&
                subProjectDoc->Saved && areSubProjectsSaved(subProjectDoc);
        }
    }

    return returnValue;
}

BOOL Utils::saveProject(MSProject::_IProjectDocPtr project)
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
    return app->FileSave();
}

CString Utils::FormatWork(int valMinutes)
{
    CString strFormated;
    double dFormatedVal;

    //convert the minutes to selected format
    switch (m_defWorkUnit)
    {
    case MSProject::pjMinute:
        dFormatedVal = (double)valMinutes;
        break;
    case MSProject::pjHour:
        dFormatedVal = (double)(valMinutes / 60.0f);
        break;
    case MSProject::pjDay:
        dFormatedVal = (double)(valMinutes / (60.0f * m_hoursPerDay));
        break;
    case MSProject::pjWeek:
        dFormatedVal = (double)(valMinutes / (60.0f * m_hoursPerWeek));
        break;
    case MSProject::pjMonthUnit:
        dFormatedVal = (double)(valMinutes / (60.0f * m_daysPerMonth * m_hoursPerDay));
        break;
    }

    //final format of the value
    CString strUnitFormated;
    strFormated = CTU_Utils::FormatFloat(dFormatedVal);

    //add unit label
    switch (m_defWorkUnit)
    {
    case MSProject::pjMinute:
        strFormated += _T(" ") + m_minuteLabels[m_minuteLabelDisplay];
        break;
    case MSProject::pjHour:
        strFormated += _T(" ") + m_hourLabels[m_hourLabelDisplay];
        break;
    case MSProject::pjDay:
        strFormated += _T(" ") + m_dayLabels[m_dayLabelDisplay];
        break;
    case MSProject::pjWeek:
        strFormated += _T(" ") + m_weekLabels[m_weekLabelDisplay];
        break;
    case MSProject::pjMonthUnit:
        strFormated += _T(" ") + m_monthLabels[m_monthLabelDisplay];
        break;
    }

    return strFormated;
}
int Utils::GetProjectMinorVersion(CString strBuild)
{
    int build = -1;
    std::vector<CString> strVec;
    CString strToken;
    int nCurPos = 0;
    strToken = strBuild.Tokenize(_T(",."), nCurPos);
    while (strToken != _T(""))
    {
        strVec.push_back(strToken);
        strToken = strBuild.Tokenize(_T(",."), nCurPos);   
    }

    if (strVec.size() > 2)
        build = _wtoi(strVec[2]);

    return build;
}
int Utils::GetProjectMajorVersion(CString strBuild)
{
    int build = -1;
    std::vector<CString> strVec;
    CString strToken;
    int nCurPos = 0;
    strToken = strBuild.Tokenize(_T(",."), nCurPos);
    while (strToken != _T(""))
    {
        strVec.push_back(strToken);
        strToken = strBuild.Tokenize(_T(",."), nCurPos);   
    }

    if (strVec.size() > 0)
        build = _wtoi(strVec[0]);

    return build;
}

void Utils::initAppUnits()
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));   
    MSProject::_IProjectDocPtr project = app->GetActiveProject(); 

    m_minuteLabelDisplay = project->MinuteLabelDisplay;
    m_hourLabelDisplay = project->HourLabelDisplay;
    m_dayLabelDisplay = project->DayLabelDisplay;
    m_weekLabelDisplay = project->WeekLabelDisplay;
    m_monthLabelDisplay = project->MonthLabelDisplay;

    m_defWorkUnit = project->DefaultWorkUnits;
    m_hoursPerDay = project->HoursPerDay;
    m_hoursPerWeek = project->HoursPerWeek;
    m_daysPerMonth = project->DaysPerMonth;

    long localeID = app->LocaleID();
    switch (localeID)
    {
    case 1033://English
        {
            m_minuteLabels[0] = _T("m");
            m_minuteLabels[1] = _T("mins");
            m_minuteLabels[2] = _T("minutes");
            m_hourLabels[0] = _T("h");
            m_hourLabels[1] = _T("hrs");
            m_hourLabels[2] = _T("hours");
            m_dayLabels[0] = _T("d");
            m_dayLabels[1] = _T("dys");
            m_dayLabels[2] = _T("days");
            m_weekLabels[0] = _T("w");
            m_weekLabels[1] = _T("wks");
            m_weekLabels[2] = _T("weeks");
            m_monthLabels[0] = _T("mo");
            m_monthLabels[1] = _T("mons");
            m_monthLabels[2] = _T("months");
        }
        break;
    case 1031://German
        {
            m_minuteLabels[0] = _T("min");
            m_minuteLabels[1] = _T("Min.");
            m_minuteLabels[2] = _T("Minuten");
            m_hourLabels[0] = _T("h");
            m_hourLabels[1] = _T("Std.");
            m_hourLabels[2] = _T("Stunden");
            m_dayLabels[0] = _T("t");
            m_dayLabels[1] = _T("Tage");
            m_weekLabels[0] = _T("W");
            m_weekLabels[1] = _T("Wochen");
            m_monthLabels[0] = _T("M");
            m_monthLabels[1] = _T("Monate");
        }
        break;
    case 1036://French
        {
            m_minuteLabels[0] = _T("m");
            m_minuteLabels[1] = _T("min");
            m_minuteLabels[2] = _T("minutes");
            m_hourLabels[0] = _T("h");
            m_hourLabels[1] = _T("h");
            m_hourLabels[2] = _T("heure");
            m_dayLabels[0] = _T("j");
            m_dayLabels[1] = _T("jrs");
            m_dayLabels[2] = _T("jours");
            m_weekLabels[0] = _T("s");
            m_weekLabels[1] = _T("sm");
            m_weekLabels[2] = _T("semaines");
            m_monthLabels[0] = _T("ms");
            m_monthLabels[1] = _T("mois");
        }
        break;
    case 1029://Czech
        {
            m_minuteLabels[0] = _T("m");
            m_minuteLabels[1] = _T("minut");
            m_hourLabels[0] = _T("h");
            m_hourLabels[1] = _T("hodin");
            m_dayLabels[0] = _T("d");
            m_dayLabels[1] = _T("dny");
            m_weekLabels[0] = _T("t");
            m_weekLabels[1] = _T("týdny");
            m_monthLabels[0] = _T("měs");
            m_monthLabels[1] = _T("měs.");
            m_monthLabels[2] = _T("měsíce");
        }
        break;
    case 1041://Japanese
        {
            m_minuteLabels[0] = _T("分");
            m_minuteLabels[1] = _T("分間");
            m_hourLabels[0] = _T("時");
            m_hourLabels[1] = _T("時間");
            m_dayLabels[0] = _T("日");
            m_dayLabels[1] = _T("日間");
            m_weekLabels[0] = _T("週");
            m_weekLabels[1] = _T("週間");
            m_monthLabels[0] = _T("月");
            m_monthLabels[1] = _T("月間");
        }
        break;
    case 1049://Russian
        {
            m_minuteLabels[0] = _T("м");
            m_minuteLabels[1] = _T("мин");
            m_minuteLabels[2] = _T("минут");
            m_hourLabels[0] = _T("ч");
            m_hourLabels[1] = _T("часов");
            m_dayLabels[0] = _T("д");
            m_dayLabels[1] = _T("дн");
            m_dayLabels[2] = _T("дней");
            m_weekLabels[0] = _T("н");
            m_weekLabels[1] = _T("нед");
            m_weekLabels[2] = _T("недель");
            m_monthLabels[0] = _T("мес");
            m_monthLabels[1] = _T("месяцев");
        }
        break;
    case 3082://Spanish
        {
            m_minuteLabels[0] = _T("m");
            m_minuteLabels[1] = _T("mins");
            m_minuteLabels[2] = _T("minutos");
            m_hourLabels[0] = _T("h");
            m_hourLabels[1] = _T("hrs");
            m_hourLabels[2] = _T("horas");
            m_dayLabels[0] = _T("d");
            m_dayLabels[1] = _T("dís");
            m_dayLabels[2] = _T("días");
            m_weekLabels[0] = _T("s");
            m_weekLabels[1] = _T("sem.");
            m_weekLabels[2] = _T("semanas");
            m_monthLabels[0] = _T("me");
            m_monthLabels[1] = _T("mss");
            m_monthLabels[2] = _T("meses");
        }
        break;
    case 2052://Chinese
        {
            m_minuteLabels[0] = _T("分钟工的");
            m_hourLabels[0] = _T("工的");
            m_dayLabels[0] = _T("个工作日");
            m_weekLabels[0] = _T("周工的");
            m_monthLabels[0] = _T("月工的");
        }
        break;
    default:
        {
            m_minuteLabels[0] = _T("");
            m_minuteLabels[1] = _T("");
            m_minuteLabels[2] = _T("");
            m_hourLabels[0] = _T("");
            m_hourLabels[1] = _T("");
            m_hourLabels[2] = _T("");
            m_dayLabels[0] = _T("");
            m_dayLabels[1] = _T("");
            m_dayLabels[2] = _T("");
            m_weekLabels[0] = _T("");
            m_weekLabels[1] = _T("");
            m_weekLabels[2] = _T("");
            m_monthLabels[0] = _T("");
            m_monthLabels[1] = _T("");
            m_monthLabels[2] = _T("");
        }
        break;
    }
}

void Utils::setCustomField(int fieldID, enumCustomFields eCustomField)
{
        
    if (-1 != fieldID)
    {    
        MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));  

        bool graphInd = 
            ePercComplGraphInd == eCustomField || 
            ePercWorkComplGraphInd == eCustomField || 
            eActWorkGraphInd == eCustomField ||
            eActStartGraphInd == eCustomField ||
            eActFinishGraphInd == eCustomField ||
            eActDurationGraphInd == eCustomField;
        const _variant_t name = Utils::m_customFieldNames[eCustomField].AllocSysString();

        int appVer = Utils::GetProjectMajorVersion((BSTR)app->Version);
        if (appVer == 14)
        {
            app->CustomFieldDelete((MSProject::PjCustomField)fieldID);
        }
        app->CustomFieldRename((MSProject::PjCustomField)fieldID, L"");
        app->CustomFieldRename((MSProject::PjCustomField)fieldID, name);
        app->CustomFieldProperties((MSProject::PjCustomField)fieldID,
            MSProject::pjFieldAttributeNone,
            MSProject::pjCalcNone,
            graphInd);

        if (graphInd)
        {
            app->CustomFieldIndicatorAdd((MSProject::PjCustomField)fieldID, 
                MSProject::pjCompareEquals, 
                "Yes", 
                MSProject::pjIndicatorSphereRed,
                MSProject::pjCriteriaNonSummary);

            app->CustomFieldIndicatorAdd((MSProject::PjCustomField)fieldID, 
                MSProject::pjCompareEquals, 
                "No", 
                MSProject::pjIndicatorSphereGreen,
                MSProject::pjCriteriaNonSummary);
        }
    }
}

bool Utils::isTaskUpdateCustomField(int fieldID)
{

    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application));

    CString name = app->CustomFieldGetName((MSProject::PjCustomField)fieldID);

    std::map<enumCustomFields, CString>::iterator itName;
    for (itName = m_customFieldNames.begin(); itName != m_customFieldNames.end(); itName++)
    {
        if (itName->second == name)
            return true;
    }

    return false;
}

void Utils::closeCollaborativeProjects(MSProject::_MSProjectPtr app, CTUProjectInfo* pPrjInfo)
{
    int count = app->Projects->Count;
    for (int i=1; i <= count; i++)
    {
		try
		{
			MSProject::_IProjectDocPtr prj = app->Projects->GetItem(i);  
			if (prj)
			{
				_bstr_t prjPath = prj->FullName;
				if (hasCollaborativeChild(prj, pPrjInfo))
				{
					try //the sub projects even if they are not opened in separate window will be listed as projects but they cannot be activated, thats the reason for the try-catch
					{
						prj->Activate();
						app->FileClose( MSProject::pjSave);

						closeCollaborativeProjects(app, pPrjInfo);//start closing again
						break;
					}
					catch (...)
					{

					}
				}
			}
		}
		catch(...)
		{
		}
    }
}

bool Utils::isProjectCollaborative(CString path, CTUProjectInfo* pPrjInfo)
{
    CTUProjectInfo* pPrj = pPrjInfo->GetProjectInfoByPath(path);
    return pPrj ? true : false;
}
bool Utils::hasCollaborativeChild(MSProject::_IProjectDocPtr project, CTUProjectInfo* pPrjInfo)
{
    _bstr_t projectPath = project->FullName;

    if (isProjectCollaborative((BSTR)projectPath, pPrjInfo))
        return true;

    MSProject::_IProjectDocPtr subProjectDoc = NULL;
    MSProject::SubprojectsPtr subProjectsList = project->Subprojects;

    for (int i = 1; i <= subProjectsList->GetCount(); i++)
    {
        if (MSProject::SubprojectPtr subProject = subProjectsList->Item[i])
        {         
            subProjectDoc = subProject->GetSourceProject();
            _bstr_t path = subProjectDoc->FullName;

            if (isProjectCollaborative((BSTR)path, pPrjInfo))
            {          
                return true;
            }
            else
            {
                if (hasCollaborativeChild(subProjectDoc, pPrjInfo))
                {
                    return true;
                }
            }
        }
    }

    return false;
}

MSProject::AssignmentPtr      Utils::getPjAssignment(MSProject::_IProjectDocPtr project, MSProject::TaskPtr pTask, long assignmentUID)
{
    if (pTask)
    {
        for (int i = 1; i <= pTask->Assignments->Count; i++)
        {
            MSProject::AssignmentPtr  pAssignment = pTask->Assignments->GetItem(i);
            if (pAssignment && ((pAssignment->UniqueID == assignmentUID) || (pAssignment->UniqueID == (assignmentUID + 0x100000))))
            {
                return pAssignment;
            }
        }
    }

    return NULL;
}

MSProject::_IProjectDocPtr Utils::getSubProject(MSProject::_IProjectDocPtr project, CString path)
{
    MSProject::_IProjectDocPtr subProjectDoc = NULL;
    MSProject::SubprojectsPtr subProjectsList = project->Subprojects;

    for (int i = 1; i <= subProjectsList->GetCount(); i++)
    {
        if (MSProject::SubprojectPtr subProject = subProjectsList->Item[i])
        {         
            subProjectDoc = subProject->GetSourceProject();
            if (0 == wcscmp(subProjectDoc->FullName, path))
            {          
                return subProjectDoc;
            }
            else
            {
                subProjectDoc = getSubProject(subProjectDoc, path);
                if (subProjectDoc)
                    return subProjectDoc;
            }
        }
    }

    return NULL;
}

CTime Utils::toGMT(CTime t)
{
	if (t != CTime(0))
	{
		tm gmtTime;
        tm localTime;
        t.GetGmtTm(&gmtTime);
        t.GetLocalTm(&localTime);
        int diff = gmtTime.tm_hour - localTime.tm_hour;

		t += diff*60*60; 
	}

	return t;
}

CString Utils::get_date_in_user_format (CTime& time)
{
    CString strTmpFormat;
    CString strDate;
    WCHAR* szData = NULL;
    int num_chars = 
        GetLocaleInfoW(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, szData, 0);
    if (num_chars != 0)
    {
        szData = new WCHAR[num_chars+1];
        szData[num_chars] = '\0';
        GetLocaleInfoW(LOCALE_USER_DEFAULT, 
            LOCALE_SSHORTDATE, szData, num_chars);
        CString strTmp (szData);
        int ind = 0;
        int len = strTmp.GetLength();
        while (ind < len)
        {
            switch (strTmp[ind])
            {
            case 'y':
                {
                    int year_type = 0;
                    while (ind < len && strTmp[ind] == 'y'){ 
                        ind++;
                        year_type++;
                    }
                    ind--;
                    switch (year_type){
                    case 4: strTmpFormat.Format(_T("%d"), time.GetYear());
                        strDate += strTmpFormat; break;
                    case 2: strTmpFormat.Format(_T("%02d"), time.GetYear() % 100);
                        strDate += strTmpFormat; break;
                    case 1: strTmpFormat.Format(_T("%d"), time.GetYear() % 10);
                        strDate += strTmpFormat; break;
                    }
                    break;
                }
            case 'M':
                {
                    int month_type = 0;
                    while (ind < len && strTmp[ind] == 'M'){ 
                        ind++;
                        month_type++;
                    }
                    ind--;
                    switch (month_type){
                    case 4:
                        {
                            WCHAR szMonth[500]={0};
                            if (0<GetLocaleInfoW(LOCALE_USER_DEFAULT, 
                                LOCALE_SMONTHNAME1+time.GetMonth()-1, szMonth, 499)){
                                    strDate += szMonth;
                            }
                            break;
                        }
                    case 3:
                        {
                            WCHAR szMonth[500]={0};
                            if (0<GetLocaleInfoW(LOCALE_USER_DEFAULT, 
                                LOCALE_SABBREVMONTHNAME1+time.GetMonth()-1, 
                                szMonth, 499)){
                                    strDate += szMonth;
                            }
                            break;
                        }
                    case 2: strTmpFormat.Format(_T("%02d"), time.GetMonth());
                        strDate += strTmpFormat; break;
                    case 1: strTmpFormat.Format(_T("%d"), time.GetMonth());
                        strDate += strTmpFormat; break;
                    }
                    break;
                }
            case 'd':
                {
                    int day_type = 0;
                    while (ind < len && strTmp[ind] == 'd'){ 
                        ind++;
                        day_type++;
                    }
                    ind--;
                    switch (day_type){
                    case 4:
                        {
                            UINT DayOfWeekFull[] = {
                                LOCALE_SDAYNAME7,   // Sunday
                                LOCALE_SDAYNAME1,   
                                LOCALE_SDAYNAME2,
                                LOCALE_SDAYNAME3,
                                LOCALE_SDAYNAME4, 
                                LOCALE_SDAYNAME5, 
                                LOCALE_SDAYNAME6   // Saturday
                            };
                            WCHAR szDayOfWeek[500]={0};
                            if (0<GetLocaleInfoW(LOCALE_USER_DEFAULT, 
                                DayOfWeekFull[time.GetDayOfWeek()-1], 
                                szDayOfWeek, 499)){
                                    strDate += szDayOfWeek;
                            }
                            break;
                        }
                    case 3:
                        {
                            UINT DayOfWeekAbbr[] = {
                                LOCALE_SABBREVDAYNAME7,   // Sunday
                                LOCALE_SABBREVDAYNAME1,   
                                LOCALE_SABBREVDAYNAME2,
                                LOCALE_SABBREVDAYNAME3,
                                LOCALE_SABBREVDAYNAME4, 
                                LOCALE_SABBREVDAYNAME5, 
                                LOCALE_SABBREVDAYNAME6   // Saturday
                            };
                            WCHAR szDayOfWeek[500]={0};
                            if (0<GetLocaleInfoW(LOCALE_USER_DEFAULT, 
                                DayOfWeekAbbr[time.GetDayOfWeek()-1], 
                                szDayOfWeek, 499)){
                                    strDate += szDayOfWeek;
                            }
                            break;
                        }
                    case 2: strTmpFormat.Format(_T("%02d"), time.GetDay());
                        strDate += strTmpFormat; break;
                    case 1: strTmpFormat.Format(_T("%d"), time.GetDay());
                        strDate += strTmpFormat; break;
                    }
                    break;
                }
            default:
                strDate += CString(strTmp[ind]);
                break;
            }
            ind++;
        }
        delete szData;
    }

    if (strDate.IsEmpty()){
        strDate = time.Format(_T("%x")); // fallback mechanism
    }

    return strDate;
}

CString Utils::get_time_in_user_format (CTime& time)
{
    CString strTmpFormat;
    CString strTime;
    WCHAR* szData = NULL;
    int num_chars = GetLocaleInfoW(LOCALE_USER_DEFAULT, 
        LOCALE_STIMEFORMAT, szData, 0);
    if (num_chars != 0)
    {
        szData = new WCHAR[num_chars+1];
        szData[num_chars] = '\0';
        GetLocaleInfoW(LOCALE_USER_DEFAULT, 
            LOCALE_STIMEFORMAT, szData, num_chars);
        CString strTmp (szData);
        int ind = 0;
        int len = strTmp.GetLength();
        while (ind < len)
        {
            switch (strTmp[ind])
            {
            case 't':
                {
                    int time_marker_type = 0;
                    while (ind < len && strTmp[ind] == 't'){ 
                        ind++;
                        time_marker_type++;
                    }
                    ind--;
                    switch (time_marker_type){
                    case 2:
                    case 1:
                        {
                            WCHAR szTimemarker[500]={0};
                            LCTYPE am_or_pm = LOCALE_S1159; //AM
                            if (time.GetHour() >= 0 && time.GetHour() < 12){
                                am_or_pm = LOCALE_S1159; //AM
                            }else{
                                am_or_pm = LOCALE_S2359; //PM
                            }
                            if (0<GetLocaleInfoW(LOCALE_USER_DEFAULT, 
                                am_or_pm, szTimemarker, 499)){
                                    if (time_marker_type == 1){
                                        strTime += CString(szTimemarker, 1);
                                    }else{
                                        strTime += szTimemarker;
                                    }
                            }
                            break;
                        }
                    }
                    break;
                }
            case 's':
                {
                    int seconds_type = 0;
                    while (ind < len && strTmp[ind] == 's'){ 
                        ind++;
                        seconds_type++;
                    }
                    ind--;
                    switch (seconds_type){
                    case 2: strTmpFormat.Format(_T("%02d"), time.GetSecond());
                        strTime += strTmpFormat; break;
                    case 1: strTmpFormat.Format(_T("%d"), time.GetSecond());
                        strTime += strTmpFormat; break;
                    }
                    break;
                }
            case 'm':
                {
                    int minute_type = 0;
                    while (ind < len && strTmp[ind] == 'm'){ 
                        ind++;
                        minute_type++;
                    }
                    ind--;
                    switch (minute_type){
                    case 2: strTmpFormat.Format(_T("%02d"), time.GetMinute());
                        strTime += strTmpFormat; break;
                    case 1: strTmpFormat.Format(_T("%d"), time.GetMinute());
                        strTime += strTmpFormat; break;
                    }
                    break;
                }
            case 'H':
                {
                    int hour_type = 0;
                    while (ind < len && strTmp[ind] == 'H'){ 
                        ind++;
                        hour_type++;
                    }
                    ind--;
                    switch (hour_type){
                    case 2: strTmpFormat.Format(_T("%02d"), time.GetHour());
                        strTime += strTmpFormat; break;
                    case 1: strTmpFormat.Format(_T("%d"), time.GetHour());
                        strTime += strTmpFormat; break;
                    }
                    break;
                }
            case 'h':
                {
                    int hour_12_format = time.GetHour() % 12;
                    if (hour_12_format==0){
                        hour_12_format = 12;
                    }
                    int hour_type = 0;
                    while (ind < len && strTmp[ind] == 'h'){ 
                        ind++;
                        hour_type++;
                    }
                    ind--;
                    switch (hour_type){
                    case 2: strTmpFormat.Format(_T("%02d"), hour_12_format);
                        strTime += strTmpFormat; break;
                    case 1: strTmpFormat.Format(_T("%d"), hour_12_format);
                        strTime += strTmpFormat; break;
                    }
                    break;
                }
            default:
                strTime += CString(strTmp[ind]);
                break;
            }
            ind++;
        }
        delete szData;
    }

    if (strTime.IsEmpty()){
        strTime = time.Format(_T("%X")); //fallback mechanism
    }

    return strTime;
}

CString Utils::get_date_time_in_user_format(CTime& time)
{
    CString dt = Utils::get_date_in_user_format(time);
    CString tm = Utils::get_time_in_user_format(time);

    CString strDateTime;
    strDateTime.Append(dt);
    strDateTime.Append(_T(" "));
    strDateTime.Append(tm);

    return strDateTime;
}

int Utils::GetYear(CTime time)
{
	int result = 1;

	if (time != 0)
	{
		tm gmtTime;
		time.GetGmtTm(&gmtTime);
		result = gmtTime.tm_year + 1900;
	}

	return result;
}

int Utils::GetMonth(CTime time)
{
	int result = 1;

	if (time != 0)
	{
		tm gmtTime;
		time.GetGmtTm(&gmtTime);
		result = gmtTime.tm_mon + 1;
	}

	return result;
}

int Utils::GetDay(CTime time)
{
	int result = 1;

	if (time != 0)
	{
		tm gmtTime;
		time.GetGmtTm(&gmtTime);
		result = gmtTime.tm_mday;
	}

	return result;
}

CString Utils::formatMSTime(CTime time)
{

	MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 

	CString strDateTime = Utils::get_date_time_in_user_format(time);
	const _variant_t varTime = strDateTime.AllocSysString();

	_bstr_t varFormattedTime = app->DateFormat(varTime);


	return (BSTR)varFormattedTime;
}
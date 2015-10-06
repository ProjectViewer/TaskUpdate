#pragma once
#include "DocCustomProperties.h"
#include <map>
#include <vector>
#include "TaskUpdate_ProjectInfo.h"
#include "TaskUpdate_ProjectReport.h"
#include "TaskUpdate_XMLParser.h"
#include "TaskUpdate_Selection.h"

enum enumTUUnit
{
    TUUnitNone              = -1,
    TUUnitMin               = 0,
    TUUnitHour              = 1,
    TUUnitDay               = 2,
    TUUnitWeek              = 3,
    TUUnitMonth             = 4,
};
typedef std::map<enumTUUnit, std::vector<CString> > TUUnitsMap;

class CTUProjectReport;
class CTUProjectInfo;
typedef std::vector<CTUProjectInfo*> TUProjectsInfoList;

class CTU_Utils
{
public:
    
    static void GetFilePaths(CString mppFilePath, CString& spvFilePath, CString& updatesDirPath);
    static bool FileExists(CString filePath);
    static bool DirExists(CString dirPath);
    static bool LoadXML(CString strPath, CTUProjectReport& theProjectReport, int reportingMode); //fills project reports object from XML from the given path
    static bool SaveXML(CString strPath, CTUProjectReport* theProjectReport);                    //saves project reports object to XML to the given path
    static bool IsValidGUID(CString strGUID);
    static bool IsNumber(CString s);
    static bool IsFileTUActive(CString path);
    static CString GenerateGUID();
    static CString GetCollIDFromFile(CString path);
    static int GetReportingModeFromFile(CString path);
    static bool GetIsActiveFromFile(CString path);
    static CString GetCustomPropertyValue(CString path, CString name);
    static enumTUUnit GetUnit(TUUnitsMap* theUnitsList, CString str);
    static CString FormatFloat(double val);
    static bool IsFileWriteble(CString path);
    static bool MPPFilesExist(TUProjectsInfoList* pProjInfoList);
    static bool SPVFilesExist(TUProjectsInfoList* pProjInfoList);
    static bool AreSPVFilesWritable(TUProjectsInfoList* pProjInfoList);
    static bool DeleteSPVFiles(TUProjectsInfoList* pProjInfoList);
    static unsigned int GetCustomText1FieldID(CString path);
    static unsigned int GetCustomText2FieldID(CString path);
    static unsigned int GetCustomText3FieldID(CString path);
    static bool SaveSelectionXML(CString path, CTaskUpdate_Selection selection);
    static bool LoadSelectionXML(CString path, CTaskUpdate_Selection& selection);
};

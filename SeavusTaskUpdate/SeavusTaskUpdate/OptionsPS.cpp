// OptionsPS.cpp : implementation file
//

#include "stdafx.h"
#include "OptionsPS.h"
#include "CollaborationManager.h"
#include "Utils.h"

// COptionsPS

IMPLEMENT_DYNAMIC(COptionsPS, CPropertySheet)

COptionsPS::COptionsPS(UINT nIDCaption, CWnd* pParentWnd, UINT iSelectPage)
	:CPropertySheet(nIDCaption, pParentWnd, iSelectPage)
    ,m_CollaborationManager(new CCollaborationManager())
{
 
}

COptionsPS::COptionsPS(LPCTSTR pszCaption, CWnd* pParentWnd, UINT iSelectPage)
	:CPropertySheet(pszCaption, pParentWnd, iSelectPage)
    ,m_CollaborationManager(new CCollaborationManager())
{
    AddPage(&m_ppMapping);
    AddPage(&m_ppImportSettings);
    AddPage(&m_ppCustomTextMapping);

    SetActivePage(0);
}

COptionsPS::~COptionsPS()
{
    AddPage(&m_ppMapping);
    AddPage(&m_ppImportSettings);
    AddPage(&m_ppCustomTextMapping);

    SetActivePage(0);
}


BEGIN_MESSAGE_MAP(COptionsPS, CPropertySheet)
    ON_BN_CLICKED(IDOK, OnOK)
END_MESSAGE_MAP()


// COptionsPS message handlers


BOOL COptionsPS::OnInitDialog()
{
    BOOL bResult = CPropertySheet::OnInitDialog();

    // Remove the Apply Now button
    CWnd* pWndOk = GetDlgItem(IDOK);
    CWnd* pWndCancel = GetDlgItem(IDCANCEL);
    CWnd* pWndApplyNow = GetDlgItem(ID_APPLY_NOW);
    pWndOk->ShowWindow(TRUE);
    pWndApplyNow->ShowWindow(FALSE);
    CRect rectApplyNow, rectCancel, rectOK;

    pWndOk->GetWindowRect(rectOK);
    pWndApplyNow->GetWindowRect(rectApplyNow);
    pWndCancel->GetWindowRect(rectCancel);

    ScreenToClient(rectOK);
    ScreenToClient(rectApplyNow);
    ScreenToClient(rectCancel);

    pWndApplyNow->MoveWindow(rectCancel);
    pWndCancel->MoveWindow(rectApplyNow);
    pWndOk->MoveWindow(rectCancel);

    updateImportSettingsTab();
    updateMappingTab();
    updateCustomTextMappingTab();

    m_ppMapping.init();

    return bResult;
}
void COptionsPS::updateMappingTab()
{
    CString path = m_CollaborationManager->getProjectInfo()->GetProjectPath();

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
}
void COptionsPS::updateImportSettingsTab()
{
    CString path = m_CollaborationManager->getProjectInfo()->GetProjectPath();

    CString str = CTU_Utils::GetCustomPropertyValue(path, _T("ImportViaCustomFields"));
    m_ppImportSettings.m_bImportViaCustomFields = (str == _T("1")) ? true : false;
}
void COptionsPS::updateCustomTextMappingTab()
{
    CString path = m_CollaborationManager->getProjectInfo()->GetProjectPath();

    CString str = CTU_Utils::GetCustomPropertyValue(path, _T("CustomText1"));
    m_iCustomTextField1 = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("CustomText2"));
    m_iCustomTextField2 = str.IsEmpty() ? -1 : _ttoi(str);

    str = CTU_Utils::GetCustomPropertyValue(path, _T("CustomText3"));
    m_iCustomTextField3 = str.IsEmpty() ? -1 : _ttoi(str);
}
void COptionsPS::OnOK()
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
    MSProject::_IProjectDocPtr project = app->GetActiveProject();  

    //get active project path
    _bstr_t projectPath = project->FullName;

    //close the active file
    app->FileClose( MSProject::pjSave);

    //close all projects that need to be modified   
    Utils::closeCollaborativeProjects(app, m_CollaborationManager->getProjectInfo());

    //update files
    std::vector<std::string> valuesList;
    valuesList.push_back(int2string(m_iPercComplReport));
    valuesList.push_back(int2string(m_iPercComplDiff));
    valuesList.push_back(int2string(m_iPercComplGraphInd));
    valuesList.push_back(int2string(m_iPercWorkComplReport));
    valuesList.push_back(int2string(m_iPercWorkComplDiff));
    valuesList.push_back(int2string(m_iPercWorkComplGraphInd));
    valuesList.push_back(int2string(m_iActWorkReport));
    valuesList.push_back(int2string(m_iActWorkDiff));
    valuesList.push_back(int2string(m_iActWorkGraphInd));
	valuesList.push_back(int2string(m_iActStartReport));
	valuesList.push_back(int2string(m_iActStartDiff));
	valuesList.push_back(int2string(m_iActStartGraphInd));
	valuesList.push_back(int2string(m_iActFinishReport));
	valuesList.push_back(int2string(m_iActFinishDiff));
	valuesList.push_back(int2string(m_iActFinishGraphInd));
    valuesList.push_back(int2string(m_iActDurationReport));
    valuesList.push_back(int2string(m_iActDurationDiff));
    valuesList.push_back(int2string(m_iActDurationGraphInd));
    valuesList.push_back(int2string(m_iCustomTextField1));
    valuesList.push_back(int2string(m_iCustomTextField2));
    valuesList.push_back(int2string(m_iCustomTextField3));
    writeMappingValuesToFile(m_CollaborationManager->GetProjInfoList(), valuesList);

    writeCustomDocPropertyToFile(m_CollaborationManager->GetProjInfoList(), L"ImportViaCustomFields", bool2string(m_ppImportSettings.m_bImportViaCustomFields));
    
    //re-open active project
    app->FileOpen(projectPath, 
        vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
        vtMissing,vtMissing,vtMissing, MSProject::pjDoNotOpenPool, vtMissing, vtMissing, vtMissing, vtMissing );

    setCustomFields();

    CPropertySheet::OnClose();
}

void COptionsPS::writeMappingValuesToFile(TUProjectsInfoList* pProjectInfoList, std::vector<std::string>& valuesList)
{
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjectInfoList->begin(); itPrjInfo != pProjectInfoList->end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->HasChilds())
        {
            writeMappingValuesToFile((*itPrjInfo)->GetSubProjectsList(), valuesList);
        }

        (*itPrjInfo)->SetMappingValues(valuesList);
    }
}
void COptionsPS::writeCustomDocPropertyToFile(TUProjectsInfoList* pProjectInfoList, const std::wstring strKey, std::string strVal)
{
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjectInfoList->begin(); itPrjInfo != pProjectInfoList->end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->HasChilds())
        {
            writeCustomDocPropertyToFile((*itPrjInfo)->GetSubProjectsList(), strKey, strVal);
        }

        (*itPrjInfo)->SetCustomDocProperty(strKey, strVal);
    }
}
void COptionsPS::writeCollPropertiesToFile(TUProjectsInfoList* pProjectInfoList, bool isCollaborative, int reportingMode)
{
    TUProjectsInfoList::iterator itPrjInfo;
    for (itPrjInfo = pProjectInfoList->begin(); itPrjInfo != pProjectInfoList->end(); ++itPrjInfo)
    {
        if ((*itPrjInfo)->HasChilds())
        {
            writeCollPropertiesToFile((*itPrjInfo)->GetSubProjectsList(), isCollaborative, reportingMode);
        }

        (*itPrjInfo)->SetFileCollaborational((*itPrjInfo)->GetProjectPath(), isCollaborative, reportingMode);
    }
}

void COptionsPS::setCustomFields()
{
    MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 

    //on change delete old task update custom fields
    for (int index = 1; index <= 30; index++)
    {
        int fieldID = getCustomTextFieldID(index);
        if (Utils::isTaskUpdateCustomField(fieldID))
        {
            int appVer = Utils::GetProjectMajorVersion((BSTR)app->Version);
            if (appVer == 14)
            {
                app->CustomFieldDelete((MSProject::PjCustomField)fieldID);
            }
            app->CustomFieldRename((MSProject::PjCustomField)fieldID, L"");
            app->CustomFieldProperties((MSProject::PjCustomField)fieldID,
                MSProject::pjFieldAttributeNone,
                MSProject::pjCalcNone,
                false);
        }
    }

    Utils::setCustomField(m_iPercComplReport, ePercComplReport);
    Utils::setCustomField(m_iPercComplDiff, ePercComplDiff);
    Utils::setCustomField(m_iPercComplGraphInd, ePercComplGraphInd);

    Utils::setCustomField(m_iPercWorkComplReport, ePercWorkComplReport);
    Utils::setCustomField(m_iPercWorkComplDiff, ePercWorkComplDiff);
    Utils::setCustomField(m_iPercWorkComplGraphInd, ePercWorkComplGraphInd);

    Utils::setCustomField(m_iActWorkReport, eActWorkReport);
    Utils::setCustomField(m_iActWorkDiff, eActWorkDiff);
    Utils::setCustomField(m_iActWorkGraphInd, eActWorkGraphInd);

	Utils::setCustomField(m_iActStartReport, eActStartReport);
	Utils::setCustomField(m_iActStartDiff, eActStartDiff);
	Utils::setCustomField(m_iActStartGraphInd, eActStartGraphInd);

	Utils::setCustomField(m_iActFinishReport, eActFinishReport);
	Utils::setCustomField(m_iActFinishDiff, eActFinishDiff);
	Utils::setCustomField(m_iActFinishGraphInd, eActFinishGraphInd);

    Utils::setCustomField(m_iActDurationReport, eActDurationReport);
    Utils::setCustomField(m_iActDurationDiff, eActDurationDiff);
    Utils::setCustomField(m_iActDurationGraphInd, eActDurationGraphInd);

    Utils::setCustomField(m_iCustomTextField1, eCustomText1);
    Utils::setCustomField(m_iCustomTextField2, eCustomText2);
    Utils::setCustomField(m_iCustomTextField3, eCustomText3);
}

std::string COptionsPS::int2string(int val)
{
    std::stringstream strStream;
    strStream << val;
    return (val != -1) ? strStream.str() : "";
}

std::string COptionsPS::bool2string(bool val)
{
    return val ? "1" : "0";
}

int COptionsPS::getCustomTextFieldID(int index)
{
    switch (index)
    {
    case 1:
        return MSProject::pjCustomTaskText1;
        break;
    case 2:
        return MSProject::pjCustomTaskText2;
        break;
    case 3:
        return MSProject::pjCustomTaskText3;
        break;
    case 4:
        return MSProject::pjCustomTaskText4;
        break;
    case 5:
        return MSProject::pjCustomTaskText5;
        break;
    case 6:
        return MSProject::pjCustomTaskText6;
        break;
    case 7:
        return MSProject::pjCustomTaskText7;
        break;
    case 8:
        return MSProject::pjCustomTaskText8;
        break;
    case 9:
        return MSProject::pjCustomTaskText9;
        break;
    case 10:
        return MSProject::pjCustomTaskText10;
        break;
    case 11:
        return MSProject::pjCustomTaskText11;
        break;
    case 12:
        return MSProject::pjCustomTaskText12;
        break;
    case 13:
        return MSProject::pjCustomTaskText13;
        break;
    case 14:
        return MSProject::pjCustomTaskText14;
        break;
    case 15:
        return MSProject::pjCustomTaskText15;
        break;
    case 16:
        return MSProject::pjCustomTaskText16;
        break;
    case 17:
        return MSProject::pjCustomTaskText17;
        break;
    case 18:
        return MSProject::pjCustomTaskText18;
        break;
    case 19:
        return MSProject::pjCustomTaskText19;
        break;
    case 20:
        return MSProject::pjCustomTaskText20;
        break;
    case 21:
        return MSProject::pjCustomTaskText21;
        break;
    case 22:
        return MSProject::pjCustomTaskText22;
        break;
    case 23:
        return MSProject::pjCustomTaskText23;
        break;
    case 24:
        return MSProject::pjCustomTaskText24;
        break;
    case 25:
        return MSProject::pjCustomTaskText25;
        break;
    case 26:
        return MSProject::pjCustomTaskText26;
        break;
    case 27:
        return MSProject::pjCustomTaskText27;
        break;
    case 28:
        return MSProject::pjCustomTaskText28;
        break;
    case 29:
        return MSProject::pjCustomTaskText29;
        break;
    case 30:
        return MSProject::pjCustomTaskText30;
        break;
    }

    return -1;
}

bool COptionsPS::customFieldIsMapped(int fieldID, int excludeFieldID)
{
    if (m_iCustomTextField1 == fieldID && m_iCustomTextField1 != excludeFieldID)
        return true;
    if (m_iCustomTextField2 == fieldID && m_iCustomTextField2 != excludeFieldID)
        return true;
    if (m_iCustomTextField3 == fieldID && m_iCustomTextField3 != excludeFieldID)
        return true;

    if (m_iPercComplReport == fieldID && m_iPercComplReport != excludeFieldID)
        return true;
    if (m_iPercComplDiff == fieldID && m_iPercComplDiff != excludeFieldID)
        return true;
    if (m_iPercComplGraphInd == fieldID && m_iPercComplGraphInd != excludeFieldID)
        return true;

    if (m_iPercWorkComplReport == fieldID && m_iPercWorkComplReport != excludeFieldID)
        return true;
    if (m_iPercWorkComplDiff == fieldID && m_iPercWorkComplDiff != excludeFieldID)
        return true;
    if (m_iPercWorkComplGraphInd == fieldID && m_iPercWorkComplGraphInd != excludeFieldID)
        return true;

    if (m_iActWorkReport == fieldID && m_iActWorkReport != excludeFieldID)
        return true;
    if (m_iActWorkDiff == fieldID && m_iActWorkDiff != excludeFieldID)
        return true;
    if (m_iActWorkGraphInd == fieldID && m_iActWorkGraphInd != excludeFieldID)
        return true;

	if (m_iActStartReport == fieldID && m_iActStartReport != excludeFieldID)
		return true;
	if (m_iActStartDiff == fieldID && m_iActStartDiff != excludeFieldID)
		return true;
	if (m_iActStartGraphInd == fieldID && m_iActStartGraphInd != excludeFieldID)
		return true;

	if (m_iActFinishReport == fieldID && m_iActFinishReport != excludeFieldID)
		return true;
	if (m_iActFinishDiff == fieldID && m_iActFinishDiff != excludeFieldID)
		return true;
	if (m_iActFinishGraphInd == fieldID && m_iActFinishGraphInd != excludeFieldID)
		return true;

    if (m_iActDurationReport == fieldID && m_iActDurationReport != excludeFieldID)
        return true;
    if (m_iActDurationDiff == fieldID && m_iActDurationDiff != excludeFieldID)
        return true;
    if (m_iActDurationGraphInd == fieldID && m_iActDurationGraphInd != excludeFieldID)
        return true;

    return false;
}
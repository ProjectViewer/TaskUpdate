#pragma once
#include "MappingPP.h"
#include "ImportSettingPP.h"
#include "CustomTextMappingPP.h"
#include <sstream>
#include <vector>

class CCollaborationManager;
// COptionsPS

class COptionsPS : public CPropertySheet
{
	DECLARE_DYNAMIC(COptionsPS)

public:
	COptionsPS(UINT nIDCaption, CWnd* pParentWnd = NULL, UINT iSelectPage = 0);
	COptionsPS(LPCTSTR pszCaption, CWnd* pParentWnd = NULL, UINT iSelectPage = 0);
	virtual ~COptionsPS();

protected:
	DECLARE_MESSAGE_MAP()

private:
    boost::shared_ptr<CCollaborationManager> m_CollaborationManager;

    CMappingPP m_ppMapping;
    CImportSettingPP m_ppImportSettings;
    CCustomTextMappingPP m_ppCustomTextMapping;

public:
    void updateMappingTab();
    void updateImportSettingsTab();
    void updateCustomTextMappingTab();

    bool customFieldIsMapped(int fieldID, int excludeFieldID);
    int getCustomTextFieldID(int index);

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

private:
    virtual BOOL OnInitDialog();
    void OnOK();
    
    void writeMappingValuesToFile(TUProjectsInfoList* pProjectInfoList, std::vector<std::string>& valuesList);
    void writeCustomDocPropertyToFile(TUProjectsInfoList* pProjectInfoList, const std::wstring strKey, std::string strDesc);
    void writeCollPropertiesToFile(TUProjectsInfoList* pProjectInfoList, bool bIsCollaborative, int reportingMode);
    void setCustomFields();

    std::string int2string(int val);
    std::string bool2string(bool val);
};



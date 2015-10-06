#pragma once
#include "TaskUpdate_ProjectInfo.h"
#include "TaskUpdate_ComboBox.h"

class COptionsPS;

class CMappingPP : public CPropertyPage
{
	DECLARE_DYNAMIC(CMappingPP)

public:
	CMappingPP();
	virtual ~CMappingPP();

// Dialog Data
	enum { IDD = IDD_DLG_MAPPING_PP };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

private:
    COptionsPS* m_pOptions;

    CButton m_radPercCompl;
    CButton m_radPercWorkCompl;
    CButton m_radActWork;
	CButton m_radActStart;
	CButton m_radActFinish;
    CButton m_radActDuration;

    CTaskUpdate_ComboBox m_comReport;
    CTaskUpdate_ComboBox m_comDiff;
    CTaskUpdate_ComboBox m_comGraphInd;

public:
    int m_iRadSelection;

private:
    virtual BOOL OnSetActive();
    afx_msg void OnBnClickedRadioMapPerccompl();
    afx_msg void OnBnClickedRadioMapPercworkcompl();
    afx_msg void OnBnClickedRadioMapActwork();
	afx_msg void OnBnClickedRadioMapActstart();
	afx_msg void OnBnClickedRadioMapActfinish();
    afx_msg void OnBnClickedRadioMapActduration();
    afx_msg void OnCbnCloseupComboRep();
    afx_msg void OnCbnCloseupComboDiff();
    afx_msg void OnCbnCloseupComboGrphind(); 

    void updateProjectMapValues();
    void updateReportFields();
    void updateDiffFields();
    void updateGraphIndFields();
    void updateRadioButtons();

public:
    void init();
};

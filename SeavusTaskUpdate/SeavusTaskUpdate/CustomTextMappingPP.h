#pragma once
#include "TaskUpdate_ComboBox.h"

class COptionsPS;

class CCustomTextMappingPP : public CPropertyPage
{
	DECLARE_DYNAMIC(CCustomTextMappingPP)

public:
	CCustomTextMappingPP();
    virtual ~CCustomTextMappingPP();

// Dialog Data
	enum { IDD = IDD_DLG_CUSTOMTEXTMAPPING_PP };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

private:
    COptionsPS* m_pOptions;

    virtual BOOL OnSetActive();
    void updateControls(int ID, CTaskUpdate_ComboBox* pCombo);
    void updateMapValues();

    afx_msg void OnCbnCloseupField1();
    afx_msg void OnCbnCloseupField2();
    afx_msg void OnCbnCloseupField3(); 

    CTaskUpdate_ComboBox m_comboField1;
    CTaskUpdate_ComboBox m_comboField2;
    CTaskUpdate_ComboBox m_comboField3;
};

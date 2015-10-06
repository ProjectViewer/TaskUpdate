#pragma once
#include "afxwin.h"
// CImportSettingPP dialog

class CImportSettingPP : public CPropertyPage
{
	DECLARE_DYNAMIC(CImportSettingPP)

public:
	CImportSettingPP();
	virtual ~CImportSettingPP();

// Dialog Data
	enum { IDD = IDD_DLG_IMPORTSETTING_PP };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

public:
    bool m_bImportViaCustomFields;

private:
    CButton m_BtnImportViaCustomFields;

private:
    virtual BOOL OnSetActive();
    void updateControls();
    afx_msg void OnBnClickedBtnImportCustomFields();
};

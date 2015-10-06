// ImportSettingPP.cpp : implementation file
//

#include "stdafx.h"
#include "ImportSettingPP.h"
#include "afxdialogex.h"
#include "OptionsPS.h"

// CImportSettingPP dialog

IMPLEMENT_DYNAMIC(CImportSettingPP, CPropertyPage)

CImportSettingPP::CImportSettingPP()
	: CPropertyPage(CImportSettingPP::IDD)
{
}

CImportSettingPP::~CImportSettingPP()
{
}

void CImportSettingPP::DoDataExchange(CDataExchange* pDX)
{
    CPropertyPage::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_BTN_IMPORT_CUSTOMFIELDS, m_BtnImportViaCustomFields);
}


BEGIN_MESSAGE_MAP(CImportSettingPP, CPropertyPage)
    ON_BN_CLICKED(IDC_BTN_IMPORT_CUSTOMFIELDS, &CImportSettingPP::OnBnClickedBtnImportCustomFields)
END_MESSAGE_MAP()


// CImportSettingPP message handlers


BOOL CImportSettingPP::OnSetActive()
{
    COptionsPS* pOptions = (COptionsPS*)GetParent();
    if (pOptions)
    {
        pOptions->updateImportSettingsTab();
    }

    updateControls();

    return CPropertyPage::OnSetActive();
}

void CImportSettingPP::updateControls()
{
    m_BtnImportViaCustomFields.SetCheck(m_bImportViaCustomFields ? BST_CHECKED : BST_UNCHECKED);
}

void CImportSettingPP::OnBnClickedBtnImportCustomFields()
{
    m_bImportViaCustomFields = (BST_CHECKED == m_BtnImportViaCustomFields.GetCheck()) ? true : false;
}

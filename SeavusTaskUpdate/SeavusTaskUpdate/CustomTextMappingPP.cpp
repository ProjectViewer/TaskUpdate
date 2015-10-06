// CustomTextMapping.cpp : implementation file
//

#include "stdafx.h"
#include "CustomTextMappingPP.h"
#include "afxdialogex.h"
#include "OptionsPS.h"


// CCustomTextMapping dialog

IMPLEMENT_DYNAMIC(CCustomTextMappingPP, CPropertyPage)

CCustomTextMappingPP::CCustomTextMappingPP()
	: CPropertyPage(CCustomTextMappingPP::IDD)
{
}

CCustomTextMappingPP::~CCustomTextMappingPP()
{
}

void CCustomTextMappingPP::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_COMBO_FIELD1, m_comboField1);
    DDX_Control(pDX, IDC_COMBO_FIELD2, m_comboField2);
    DDX_Control(pDX, IDC_COMBO_FIELD3, m_comboField3);
}


BEGIN_MESSAGE_MAP(CCustomTextMappingPP, CPropertyPage)
    ON_CBN_CLOSEUP(IDC_COMBO_FIELD1, &CCustomTextMappingPP::OnCbnCloseupField1)
    ON_CBN_CLOSEUP(IDC_COMBO_FIELD2, &CCustomTextMappingPP::OnCbnCloseupField2)
    ON_CBN_CLOSEUP(IDC_COMBO_FIELD3, &CCustomTextMappingPP::OnCbnCloseupField3)
END_MESSAGE_MAP()


BOOL CCustomTextMappingPP::OnSetActive()
{
    m_pOptions = (COptionsPS*)GetParent();

    updateControls(m_pOptions->m_iCustomTextField1, &m_comboField1);
    updateControls(m_pOptions->m_iCustomTextField2, &m_comboField2);
    updateControls(m_pOptions->m_iCustomTextField3, &m_comboField3);

    return CPropertyPage::OnSetActive();
}

void CCustomTextMappingPP::updateMapValues()
{
    m_pOptions->m_iCustomTextField1 = m_comboField1.GetSelectedItemsKey();
    m_pOptions->m_iCustomTextField2 = m_comboField2.GetSelectedItemsKey();
    m_pOptions->m_iCustomTextField3 = m_comboField3.GetSelectedItemsKey();

    updateControls(m_pOptions->m_iCustomTextField1, &m_comboField1);
    updateControls(m_pOptions->m_iCustomTextField2, &m_comboField2);
    updateControls(m_pOptions->m_iCustomTextField3, &m_comboField3);
}
void CCustomTextMappingPP::OnCbnCloseupField1()
{
    updateMapValues();
}
void CCustomTextMappingPP::OnCbnCloseupField2()
{
    updateMapValues();
}
void CCustomTextMappingPP::OnCbnCloseupField3()
{
    updateMapValues();
}

void CCustomTextMappingPP::updateControls(int ID, CTaskUpdate_ComboBox* pCombo)
{
    COptionsPS* pOptions = (COptionsPS*)GetParent();
    if (pOptions)
    {
        MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
        pCombo->ResetContent();
        pCombo->InsertString(0, _T(""));
        pCombo->SetItemKey(0, TUCMB_NONE);
        pCombo->SetItemDescription(0, _T(""));	

        int i = 1;
        for(int iItem = 1; iItem <= 30; iItem++)
        {
            int fieldID = pOptions->getCustomTextFieldID(iItem);


            if (pOptions->customFieldIsMapped(fieldID, ID))
            {
                continue;
            }

            CString strText;
            strText.Format(_T("Text%d"), iItem);

            CString str = app->CustomFieldGetName((MSProject::PjCustomField)fieldID);
            if (str.IsEmpty())
            {
                str = strText;
            }
            else 
            {
                str += _T(" (") + strText + _T(")");
            }

            pCombo->InsertString(i, str);
            pCombo->SetItemKey(i, fieldID);
            pCombo->SetItemDescription(i, str);	

            i++;
        }

        pCombo->SelectItemWithKey(ID);
    }
}
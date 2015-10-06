// MappingPP.cpp : implementation file
//

#include "stdafx.h"
#include "MappingPP.h"
#include "afxdialogex.h"
#include "TaskUpdate_Utils.h"
#include "OptionsPS.h"
// CMappingPP dialog

IMPLEMENT_DYNAMIC(CMappingPP, CPropertyPage)

CMappingPP::CMappingPP()
	: CPropertyPage(CMappingPP::IDD)
    , m_iRadSelection(1)
{

}

CMappingPP::~CMappingPP()
{
}

void CMappingPP::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_RADIO_MAP_PERCCOMPL, m_radPercCompl);
    DDX_Control(pDX, IDC_RADIO_MAP_PERCWORKCOMPL, m_radPercWorkCompl);
    DDX_Control(pDX, IDC_RADIO_MAP_ACTWORK, m_radActWork);
	DDX_Control(pDX, IDC_RADIO_MAP_ACTSTART, m_radActStart);
	DDX_Control(pDX, IDC_RADIO_MAP_ACTFINISH, m_radActFinish);
    DDX_Control(pDX, IDC_RADIO_MAP_ACTDURATION, m_radActDuration);
    DDX_Control(pDX, IDC_COMBO_REP, m_comReport);
    DDX_Control(pDX, IDC_COMBO_DIFF, m_comDiff);
    DDX_Control(pDX, IDC_COMBO_GRPHIND, m_comGraphInd);
}


BEGIN_MESSAGE_MAP(CMappingPP, CPropertyPage)
    ON_BN_CLICKED(IDC_RADIO_MAP_PERCCOMPL, &CMappingPP::OnBnClickedRadioMapPerccompl)
    ON_BN_CLICKED(IDC_RADIO_MAP_PERCWORKCOMPL, &CMappingPP::OnBnClickedRadioMapPercworkcompl)
    ON_BN_CLICKED(IDC_RADIO_MAP_ACTWORK, &CMappingPP::OnBnClickedRadioMapActwork)
	ON_BN_CLICKED(IDC_RADIO_MAP_ACTSTART, &CMappingPP::OnBnClickedRadioMapActstart)
	ON_BN_CLICKED(IDC_RADIO_MAP_ACTFINISH, &CMappingPP::OnBnClickedRadioMapActfinish)
    ON_BN_CLICKED(IDC_RADIO_MAP_ACTDURATION, &CMappingPP::OnBnClickedRadioMapActduration)
    ON_CBN_CLOSEUP(IDC_COMBO_REP, &CMappingPP::OnCbnCloseupComboRep)
    ON_CBN_CLOSEUP(IDC_COMBO_DIFF, &CMappingPP::OnCbnCloseupComboDiff)
    ON_CBN_CLOSEUP(IDC_COMBO_GRPHIND, &CMappingPP::OnCbnCloseupComboGrphind)
END_MESSAGE_MAP()


// CMappingPP message handlers


BOOL CMappingPP::OnSetActive()
{
    m_pOptions = (COptionsPS*)GetParent();

    updateRadioButtons();

    updateReportFields();
    updateDiffFields();
    updateGraphIndFields();

    return CPropertyPage::OnSetActive();
}
void CMappingPP::init()
{
    updateRadioButtons();

    updateReportFields();
    updateDiffFields();
    updateGraphIndFields();
}
void CMappingPP::updateProjectMapValues()
{
    switch (m_iRadSelection)
    {
    case 1:
        m_pOptions->m_iPercComplReport = m_comReport.GetSelectedItemsKey();
        m_pOptions->m_iPercComplDiff = m_comDiff.GetSelectedItemsKey();
        m_pOptions->m_iPercComplGraphInd = m_comGraphInd.GetSelectedItemsKey();
        break;
    case 2:
        m_pOptions->m_iPercWorkComplReport = m_comReport.GetSelectedItemsKey();
        m_pOptions->m_iPercWorkComplDiff = m_comDiff.GetSelectedItemsKey();
        m_pOptions->m_iPercWorkComplGraphInd = m_comGraphInd.GetSelectedItemsKey();
        break;
    case 3:
        m_pOptions->m_iActWorkReport = m_comReport.GetSelectedItemsKey();
        m_pOptions->m_iActWorkDiff = m_comDiff.GetSelectedItemsKey();
        m_pOptions->m_iActWorkGraphInd = m_comGraphInd.GetSelectedItemsKey();
        break;
	case 4:
		m_pOptions->m_iActStartReport = m_comReport.GetSelectedItemsKey();
		m_pOptions->m_iActStartDiff = m_comDiff.GetSelectedItemsKey();
		m_pOptions->m_iActStartGraphInd = m_comGraphInd.GetSelectedItemsKey();
		break;
	case 5:
		m_pOptions->m_iActFinishReport = m_comReport.GetSelectedItemsKey();
		m_pOptions->m_iActFinishDiff = m_comDiff.GetSelectedItemsKey();
		m_pOptions->m_iActFinishGraphInd = m_comGraphInd.GetSelectedItemsKey();
		break;
    case 6:
        m_pOptions->m_iActDurationReport = m_comReport.GetSelectedItemsKey();
        m_pOptions->m_iActDurationDiff = m_comDiff.GetSelectedItemsKey();
        m_pOptions->m_iActDurationGraphInd = m_comGraphInd.GetSelectedItemsKey();
        break;
    }

    updateReportFields();
    updateDiffFields();
    updateGraphIndFields();
}
void CMappingPP::updateReportFields()
{
    COptionsPS* pOptions = (COptionsPS*)GetParent();
    if (pOptions)
    {
        MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
        m_comReport.ResetContent();
        m_comReport.InsertString(0, _T(""));
        m_comReport.SetItemKey(0, TUCMB_NONE);
        m_comReport.SetItemDescription(0, _T(""));	

        int i = 1;
        for(int iItem = 1; iItem <= 30; iItem++)
        {
            int fieldID = pOptions->getCustomTextFieldID(iItem);

            if (BST_CHECKED == m_radPercCompl.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iPercComplReport))
                    continue;
            }
            else if (BST_CHECKED == m_radPercWorkCompl.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iPercWorkComplReport))
                    continue;
            }
            else if (BST_CHECKED == m_radActWork.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActWorkReport))
                    continue;
            }      
			else if (BST_CHECKED == m_radActStart.GetCheck())
			{
				if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActStartReport))
					continue;
			}
			else if (BST_CHECKED == m_radActFinish.GetCheck())
			{
				if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActFinishReport))
					continue;
			}
            else if (BST_CHECKED == m_radActDuration.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActDurationReport))
                    continue;
            }

            CString strText;
            strText.Format(_T("Text%d"), iItem);

            CString str = app->CustomFieldGetName((MSProject::PjCustomField)fieldID);
            if (str.IsEmpty())
                str = strText;
            else 
                str += _T(" (") + strText + _T(")");

            m_comReport.InsertString(i, str);
            m_comReport.SetItemKey(i, fieldID);
            m_comReport.SetItemDescription(i, str);	

            i++;
        }

        if (BST_CHECKED == m_radPercCompl.GetCheck())
            m_comReport.SelectItemWithKey(m_pOptions->m_iPercComplReport);
        else if (BST_CHECKED == m_radPercWorkCompl.GetCheck())
            m_comReport.SelectItemWithKey(m_pOptions->m_iPercWorkComplReport);
        else if (BST_CHECKED == m_radActWork.GetCheck())
            m_comReport.SelectItemWithKey(m_pOptions->m_iActWorkReport);
		else if (BST_CHECKED == m_radActStart.GetCheck())
			m_comReport.SelectItemWithKey(m_pOptions->m_iActStartReport);
		else if (BST_CHECKED == m_radActFinish.GetCheck())
			m_comReport.SelectItemWithKey(m_pOptions->m_iActFinishReport);
        else if (BST_CHECKED == m_radActDuration.GetCheck())
            m_comReport.SelectItemWithKey(m_pOptions->m_iActDurationReport);
    }
}
void CMappingPP::updateDiffFields()
{
    COptionsPS* pOptions = (COptionsPS*)GetParent();
    if (pOptions)
    {
        MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
        m_comDiff.ResetContent();
        m_comDiff.InsertString(0, _T(""));
        m_comDiff.SetItemKey(0, TUCMB_NONE);
        m_comDiff.SetItemDescription(0, _T(""));	

        int i = 1;
        for(int iItem = 1; iItem <= 30; iItem++)
        {
            int fieldID = pOptions->getCustomTextFieldID(iItem);

            if (BST_CHECKED == m_radPercCompl.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iPercComplDiff))
                    continue;
            }
            else if (BST_CHECKED == m_radPercWorkCompl.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iPercWorkComplDiff))
                    continue;
            }
            else if (BST_CHECKED == m_radActWork.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActWorkDiff))
                    continue;
            } 
			else if (BST_CHECKED == m_radActStart.GetCheck())
			{
				if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActStartDiff))
					continue;
			} 
			else if (BST_CHECKED == m_radActFinish.GetCheck())
			{
				if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActFinishDiff))
					continue;
			}
            else if (BST_CHECKED == m_radActDuration.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActDurationDiff))
                    continue;
            }

            CString strText;
            strText.Format(_T("Text%d"), iItem);

            CString str = app->CustomFieldGetName((MSProject::PjCustomField)fieldID);
            if (str.IsEmpty())
                str = strText;
            else 
                str += _T(" (") + strText + _T(")");

            m_comDiff.InsertString(i, str);
            m_comDiff.SetItemKey(i, fieldID);
            m_comDiff.SetItemDescription(i, str);	

            i++;
        }

        if (BST_CHECKED == m_radPercCompl.GetCheck())
            m_comDiff.SelectItemWithKey(m_pOptions->m_iPercComplDiff);
        else if (BST_CHECKED == m_radPercWorkCompl.GetCheck())
            m_comDiff.SelectItemWithKey(m_pOptions->m_iPercWorkComplDiff);
        else if (BST_CHECKED == m_radActWork.GetCheck())
            m_comDiff.SelectItemWithKey(m_pOptions->m_iActWorkDiff);
		else if (BST_CHECKED == m_radActStart.GetCheck())
			m_comDiff.SelectItemWithKey(m_pOptions->m_iActStartDiff);
		else if (BST_CHECKED == m_radActFinish.GetCheck())
			m_comDiff.SelectItemWithKey(m_pOptions->m_iActFinishDiff);
        else if (BST_CHECKED == m_radActDuration.GetCheck())
            m_comDiff.SelectItemWithKey(m_pOptions->m_iActDurationDiff);
    }
}
void CMappingPP::updateGraphIndFields()
{
    COptionsPS* pOptions = (COptionsPS*)GetParent();
    if (pOptions)
    {
        MSProject::_MSProjectPtr app(__uuidof(MSProject::Application)); 
        m_comGraphInd.ResetContent();
        m_comGraphInd.InsertString(0, _T(""));
        m_comGraphInd.SetItemKey(0, TUCMB_NONE);
        m_comGraphInd.SetItemDescription(0, _T(""));	

        int i = 1;
        for(int iItem = 1; iItem <= 30; iItem++)
        {
            int fieldID = pOptions->getCustomTextFieldID(iItem);

            if (BST_CHECKED == m_radPercCompl.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iPercComplGraphInd))
                    continue;
            }
            else if (BST_CHECKED == m_radPercWorkCompl.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iPercWorkComplGraphInd))
                    continue;
            }
            else if (BST_CHECKED == m_radActWork.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActWorkGraphInd))
                    continue;
            } 
			else if (BST_CHECKED == m_radActStart.GetCheck())
			{
				if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActStartGraphInd))
					continue;
			} 
			else if (BST_CHECKED == m_radActFinish.GetCheck())
			{
				if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActFinishGraphInd))
					continue;
			} 
            else if (BST_CHECKED == m_radActDuration.GetCheck())
            {
                if (pOptions->customFieldIsMapped(fieldID, m_pOptions->m_iActDurationGraphInd))
                    continue;
            } 

            CString strText;
            strText.Format(_T("Text%d"), iItem);

            CString str = app->CustomFieldGetName((MSProject::PjCustomField)fieldID);
            if (str.IsEmpty())
                str = strText;
            else 
                str += _T(" (") + strText + _T(")");

            m_comGraphInd.InsertString(i, str);
            m_comGraphInd.SetItemKey(i, fieldID);
            m_comGraphInd.SetItemDescription(i, str);	

            i++;
        }

        if (BST_CHECKED == m_radPercCompl.GetCheck())
            m_comGraphInd.SelectItemWithKey(m_pOptions->m_iPercComplGraphInd);
        else if (BST_CHECKED == m_radPercWorkCompl.GetCheck())
            m_comGraphInd.SelectItemWithKey(m_pOptions->m_iPercWorkComplGraphInd);
        else if (BST_CHECKED == m_radActWork.GetCheck())
            m_comGraphInd.SelectItemWithKey(m_pOptions->m_iActWorkGraphInd);
		else if (BST_CHECKED == m_radActStart.GetCheck())
			m_comGraphInd.SelectItemWithKey(m_pOptions->m_iActStartGraphInd);
		else if (BST_CHECKED == m_radActFinish.GetCheck())
			m_comGraphInd.SelectItemWithKey(m_pOptions->m_iActFinishGraphInd);
        else if (BST_CHECKED == m_radActDuration.GetCheck())
            m_comGraphInd.SelectItemWithKey(m_pOptions->m_iActDurationGraphInd);
    }
}
void CMappingPP::updateRadioButtons()
{
    switch (m_iRadSelection)
    {
    case 1:
        m_radPercCompl.SetCheck(BST_CHECKED);
        m_radPercWorkCompl.SetCheck(BST_UNCHECKED);
        m_radActWork.SetCheck(BST_UNCHECKED);
		m_radActStart.SetCheck(BST_UNCHECKED);
		m_radActFinish.SetCheck(BST_UNCHECKED);
        m_radActDuration.SetCheck(BST_UNCHECKED);
    	break;
    case 2:
        m_radPercCompl.SetCheck(BST_UNCHECKED);
        m_radPercWorkCompl.SetCheck(BST_CHECKED);
        m_radActWork.SetCheck(BST_UNCHECKED);
		m_radActStart.SetCheck(BST_UNCHECKED);
		m_radActFinish.SetCheck(BST_UNCHECKED);
        m_radActDuration.SetCheck(BST_UNCHECKED);
        break;
    case 3:
        m_radPercCompl.SetCheck(BST_UNCHECKED);
        m_radPercWorkCompl.SetCheck(BST_UNCHECKED);
        m_radActWork.SetCheck(BST_CHECKED);
		m_radActStart.SetCheck(BST_UNCHECKED);
		m_radActFinish.SetCheck(BST_UNCHECKED);
        m_radActDuration.SetCheck(BST_UNCHECKED);
		break;
	case 4:
		m_radPercCompl.SetCheck(BST_UNCHECKED);
		m_radPercWorkCompl.SetCheck(BST_UNCHECKED);
		m_radActWork.SetCheck(BST_UNCHECKED);
		m_radActStart.SetCheck(BST_CHECKED);
		m_radActFinish.SetCheck(BST_UNCHECKED);
        m_radActDuration.SetCheck(BST_UNCHECKED);
		break;
	case 5:
		m_radPercCompl.SetCheck(BST_UNCHECKED);
		m_radPercWorkCompl.SetCheck(BST_UNCHECKED);
		m_radActWork.SetCheck(BST_UNCHECKED);
		m_radActStart.SetCheck(BST_UNCHECKED);
		m_radActFinish.SetCheck(BST_CHECKED);
        m_radActDuration.SetCheck(BST_UNCHECKED);
        break;
    case 6:
        m_radPercCompl.SetCheck(BST_UNCHECKED);
        m_radPercWorkCompl.SetCheck(BST_UNCHECKED);
        m_radActWork.SetCheck(BST_UNCHECKED);
        m_radActStart.SetCheck(BST_UNCHECKED);
        m_radActFinish.SetCheck(BST_UNCHECKED);
        m_radActDuration.SetCheck(BST_CHECKED);
        break;
    }
}
void CMappingPP::OnBnClickedRadioMapPerccompl()
{
    m_iRadSelection = 1;
    updateRadioButtons();

    updateReportFields();
    updateDiffFields();
    updateGraphIndFields();
}
void CMappingPP::OnBnClickedRadioMapPercworkcompl()
{
    m_iRadSelection = 2;
    updateRadioButtons();

    updateReportFields();
    updateDiffFields();
    updateGraphIndFields();
}
void CMappingPP::OnBnClickedRadioMapActwork()
{
    m_iRadSelection = 3;
    updateRadioButtons();

    updateReportFields();
    updateDiffFields();
    updateGraphIndFields();
}
void CMappingPP::OnBnClickedRadioMapActstart()
{
	m_iRadSelection = 4;
	updateRadioButtons();

	updateReportFields();
	updateDiffFields();
	updateGraphIndFields();
}
void CMappingPP::OnBnClickedRadioMapActfinish()
{
	m_iRadSelection = 5;
	updateRadioButtons();

	updateReportFields();
	updateDiffFields();
	updateGraphIndFields();
}
void CMappingPP::OnBnClickedRadioMapActduration()
{
    m_iRadSelection = 6;
    updateRadioButtons();

    updateReportFields();
    updateDiffFields();
    updateGraphIndFields();
}
void CMappingPP::OnCbnCloseupComboRep()
{
    updateProjectMapValues();
}
void CMappingPP::OnCbnCloseupComboDiff()
{
    updateProjectMapValues();
}
void CMappingPP::OnCbnCloseupComboGrphind()
{
    updateProjectMapValues();
}


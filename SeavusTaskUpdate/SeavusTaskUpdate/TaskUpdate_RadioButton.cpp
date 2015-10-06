#include "TaskUpdate_RadioButton.h"


IMPLEMENT_DYNAMIC(CRadioButton, CButton)

    CRadioButton::CRadioButton()
{
}

CRadioButton::~CRadioButton()
{
}

BEGIN_MESSAGE_MAP(CRadioButton, CButton)
    ON_WM_CTLCOLOR_REFLECT()
    ON_WM_ERASEBKGND()
END_MESSAGE_MAP()

HBRUSH CRadioButton::CtlColor(CDC* pDC, UINT nCtlColor)
{
    CBrush br;
    br.CreateSolidBrush(RGB(255,255,255));

    return HBRUSH(br);
}

BOOL CRadioButton::OnEraseBkgnd(CDC* pDC)
{
    pDC->SetBkMode(TRANSPARENT);
    return CButton::OnEraseBkgnd(pDC);
}
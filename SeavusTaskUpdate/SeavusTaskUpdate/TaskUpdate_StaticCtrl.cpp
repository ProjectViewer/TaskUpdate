#include "TaskUpdate_StaticCtrl.h"


IMPLEMENT_DYNAMIC(CTaskUpdate_StaticCtrl, CStatic)

CTaskUpdate_StaticCtrl::CTaskUpdate_StaticCtrl()
{
}

CTaskUpdate_StaticCtrl::~CTaskUpdate_StaticCtrl()
{
}

BEGIN_MESSAGE_MAP(CTaskUpdate_StaticCtrl, CStatic)
    ON_WM_CTLCOLOR_REFLECT()
    ON_WM_ERASEBKGND()
END_MESSAGE_MAP()

HBRUSH CTaskUpdate_StaticCtrl::CtlColor(CDC* pDC, UINT nCtlColor)
{
    CBrush br;
    br.CreateSolidBrush(RGB(255,255,255));

    return HBRUSH(br);
}

BOOL CTaskUpdate_StaticCtrl::OnEraseBkgnd(CDC* pDC)
{
    pDC->SetBkMode(TRANSPARENT);
    return CStatic::OnEraseBkgnd(pDC);
}
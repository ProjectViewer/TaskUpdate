#pragma once
class CTaskUpdate_StaticCtrl : public CStatic
{
    DECLARE_DYNAMIC(CTaskUpdate_StaticCtrl)

public:
    CTaskUpdate_StaticCtrl();
    virtual ~CTaskUpdate_StaticCtrl();

protected:
    afx_msg HBRUSH CtlColor(CDC* pDC, UINT nCtlColor);
    afx_msg BOOL OnEraseBkgnd(CDC* pDC);

protected:
    afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);

    DECLARE_MESSAGE_MAP()
};
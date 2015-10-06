#pragma once
class CRadioButton : public CButton
{
    DECLARE_DYNAMIC(CRadioButton)

public:
    CRadioButton();
    virtual ~CRadioButton();

protected:

public:
    afx_msg HBRUSH CtlColor(CDC* pDC, UINT nCtlColor);
    afx_msg BOOL OnEraseBkgnd(CDC* pDC);

    DECLARE_MESSAGE_MAP()
};
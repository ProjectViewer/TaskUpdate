#pragma once
#include <vector>
struct itemTreePropsStruct 
{
	int				UniqueID;
	CString			strName;
	CString			strPath;
};

class CTaskUpdate_TreeCtrl :public CTreeCtrl
{
public:
	CTaskUpdate_TreeCtrl(void);
	~CTaskUpdate_TreeCtrl(void);

protected:
	DECLARE_MESSAGE_MAP()

	std::vector<itemTreePropsStruct*> itemPropsPtrList;//keeps the item color pointers passed to items LPARAM, for deleting them on destructor
	virtual INT_PTR		OnToolHitTest(CPoint point, TOOLINFO * pTI) const;
	BOOL				OnToolTipText( UINT id, NMHDR * pNMHDR, LRESULT * pResult);

public:
	void				SetItemName(HTREEITEM iItem, CString name);
	CString				GetItemName(HTREEITEM iItem);
	void				SetItemPath(HTREEITEM iItem, CString path);
	CString				GetItemPath(HTREEITEM iItem);
	void				SetItemUID(HTREEITEM iItem, int uid);
	int					GetItemUID(HTREEITEM iItem) const;
	HTREEITEM			GetItemByUID(int uid);

};


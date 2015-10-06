#pragma once
#include <vector>

#define TUCMB_ALL_PROJECTS	_T("")
#define TUCMB_ALL_TASKS		-1
#define TUCMB_ALL_STATUSES  -1
#define TUCMB_ALL_RESOURCES	-1
#define TUCMB_NOTVALID		-2
#define TUCMB_NONE          -1

struct itemPropsStr
{
	int				iKey;
	CString			strGUID;
	CString			strDescription;
};


class CTaskUpdate_ComboBox : public CComboBox
{
	DECLARE_DYNAMIC(CTaskUpdate_ComboBox)
public:
	CTaskUpdate_ComboBox(void);
	~CTaskUpdate_ComboBox(void);

protected:
	DECLARE_MESSAGE_MAP()

private:
	std::vector<itemPropsStr*> itemPropsPtrList;//keeps the item props pointers passed to items LPARAM, for deleting them on destructor


public:
	void			SetItemKey(int iItem, int key);
	int				GetItemKey(int iItem);

	void			SetItemDescription(int iItem, CString desc);
	CString			GetItemDescription(int iItem);

	void			SetItemGUID(int iItem, CString guid);
	CString			GetItemGUID(int iItem);

	void			SelectItemWithKey(int iKey);
	int				GeItemWithKey(int iKey);

	void			SelectItemWithGUID(CString guid);
	int				GeItemWithGUID(CString guid);

	CString			GetSelectedItemsGUID();
	int				GetSelectedItemsKey();
};


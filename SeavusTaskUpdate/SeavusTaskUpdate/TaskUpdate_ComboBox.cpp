#include "TaskUpdate_ComboBox.h"

IMPLEMENT_DYNAMIC(CTaskUpdate_ComboBox, CComboBox)

CTaskUpdate_ComboBox::CTaskUpdate_ComboBox(void)
{
	itemPropsPtrList.clear();
}

CTaskUpdate_ComboBox::~CTaskUpdate_ComboBox(void)
{
	std::vector<itemPropsStr*>::iterator it;
	for(it = itemPropsPtrList.begin(); it != itemPropsPtrList.end();)
	{
		delete *it;
		*it = NULL;
		it = itemPropsPtrList.erase(it);
	}
}

BEGIN_MESSAGE_MAP(CTaskUpdate_ComboBox, CComboBox)
END_MESSAGE_MAP()

void			CTaskUpdate_ComboBox::SetItemKey(int iItem, int key)
 {
	 DWORD_PTR ptr = GetItemData(iItem);
	 itemPropsStr* itemProps = (itemPropsStr*)ptr;
	 if(itemProps)
	 {
		 itemProps->iKey = key;
	 }
	 else
	 {
		 itemProps = new itemPropsStr();
		 itemProps->iKey = key;
		 itemProps->strDescription = _T("");
		 itemProps->strGUID = _T("");


		 itemPropsPtrList.push_back(itemProps);
		 SetItemData(iItem, (LPARAM)itemProps);
	 }
 }
 int			CTaskUpdate_ComboBox::GetItemKey(int iItem)
 {
	 DWORD_PTR ptr = GetItemData(iItem);
	 itemPropsStr* itemProps = (itemPropsStr*)ptr;
	 return itemProps ? itemProps->iKey : -1;
 }
 void			CTaskUpdate_ComboBox::SetItemDescription(int iItem, CString desc)
 {
	 DWORD_PTR ptr = GetItemData(iItem);
	 itemPropsStr* itemProps = (itemPropsStr*)ptr;
	 if(itemProps)
	 {
		 itemProps->strDescription = desc;
	 }
	 else
	 {
		 itemProps = new itemPropsStr();
		 itemProps->iKey = TUCMB_NOTVALID;
		 itemProps->strDescription = desc;
		 itemProps->strGUID = _T("");


		 itemPropsPtrList.push_back(itemProps);
		 SetItemData(iItem, (LPARAM)itemProps);
	 }
 }
 CString		CTaskUpdate_ComboBox::GetItemDescription(int iItem)
 {
	 DWORD_PTR ptr = GetItemData(iItem);
	 itemPropsStr* itemProps = (itemPropsStr*)ptr;
	 return itemProps ? itemProps->strDescription : _T("");
 }

 void			CTaskUpdate_ComboBox::SetItemGUID(int iItem, CString guid)
 {
	 DWORD_PTR ptr = GetItemData(iItem);
	 itemPropsStr* itemProps = (itemPropsStr*)ptr;
	 if(itemProps)
	 {
		 itemProps->strGUID = guid;
	 }
	 else
	 {
		 itemProps = new itemPropsStr();
		 itemProps->iKey = TUCMB_NOTVALID;
		 itemProps->strDescription = _T("");
		 itemProps->strGUID = guid;


		 itemPropsPtrList.push_back(itemProps);
		 SetItemData(iItem, (LPARAM)itemProps);
	 }
 }
CString		CTaskUpdate_ComboBox::GetItemGUID(int iItem)
 {
	 DWORD_PTR ptr = GetItemData(iItem);
	 itemPropsStr* itemProps = (itemPropsStr*)ptr;
	 return itemProps ? itemProps->strGUID : _T("");
 }


void			CTaskUpdate_ComboBox::SelectItemWithKey(int iKey)
{
	for(int iItem = 0; iItem < GetCount(); ++iItem)
	{
		if(GetItemKey(iItem)==iKey)
			SetCurSel(iItem);
	}
}
int			CTaskUpdate_ComboBox::GeItemWithKey(int iKey)
{
	for(int iItem = 0; iItem < GetCount(); ++iItem)
	{
		if(GetItemKey(iItem)==iKey)
			return iItem;
	}
	return TUCMB_NOTVALID;
}

void			CTaskUpdate_ComboBox::SelectItemWithGUID(CString guid)
{
	for(int iItem = 0; iItem < GetCount(); ++iItem)
	{
		if(GetItemGUID(iItem)==guid)
			SetCurSel(iItem);
	}
}
int			CTaskUpdate_ComboBox::GeItemWithGUID(CString guid)
{
	for(int iItem = 0; iItem < GetCount(); ++iItem)
	{
		if(GetItemGUID(iItem)==guid)
			return iItem;
	}
	return TUCMB_NOTVALID;
}

CString			CTaskUpdate_ComboBox::GetSelectedItemsGUID()
{
	int iItem= GetCurSel();
	if(iItem>-1)
	{
		return GetItemGUID(iItem);
	}
	return TUCMB_ALL_PROJECTS;
}

int				CTaskUpdate_ComboBox::GetSelectedItemsKey()
{
		int iItem= GetCurSel();
	if(iItem>-1)
	{
		return GetItemKey(iItem);
	}
	return TUCMB_NOTVALID;
}
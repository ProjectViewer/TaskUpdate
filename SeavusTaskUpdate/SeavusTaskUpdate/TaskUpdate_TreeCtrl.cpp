#include "TaskUpdate_TreeCtrl.h"

//IMPLEMENT_DYNAMIC(CTaskUpdate_TreeCtrl, CTreeCtrl)

CTaskUpdate_TreeCtrl::CTaskUpdate_TreeCtrl(void)
{
	itemPropsPtrList.clear();
	EnableToolTips();
}


CTaskUpdate_TreeCtrl::~CTaskUpdate_TreeCtrl(void)
{
	std::vector<itemTreePropsStruct*>::iterator it;
	for(it = itemPropsPtrList.begin(); it != itemPropsPtrList.end();)
	{
		delete *it;
		*it = NULL;
		it = itemPropsPtrList.erase(it);
	}
}

BEGIN_MESSAGE_MAP(CTaskUpdate_TreeCtrl, CTreeCtrl)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXTW, 0, 0xFFFF, OnToolTipText)
	ON_NOTIFY_EX_RANGE(TTN_NEEDTEXTA, 0, 0xFFFF, OnToolTipText)
END_MESSAGE_MAP()

INT_PTR			CTaskUpdate_TreeCtrl::OnToolHitTest(CPoint point, TOOLINFO * pTI)const
{
	RECT rect;
	UINT nFlags;
	HTREEITEM hitem = HitTest( point, &nFlags );
	if(nFlags & TVHT_ONITEM)
	{
		GetItemRect( hitem, &rect, TRUE );
		pTI->hwnd = m_hWnd;
		pTI->uId = (UINT)(hitem);
		pTI->lpszText = LPSTR_TEXTCALLBACK;
		pTI->rect = rect;
		return pTI->uId;
	}
	return -1;
}
BOOL			CTaskUpdate_TreeCtrl::OnToolTipText( UINT id, NMHDR * pNMHDR, LRESULT * pResult )
{
	TOOLTIPTEXTA* pTTTA = (TOOLTIPTEXTA*)pNMHDR;
	TOOLTIPTEXTW* pTTTW = (TOOLTIPTEXTW*)pNMHDR;
	CString strTipText(_T(""));
	 UINT nID = (UINT)pNMHDR->idFrom;
	
	 if( nID == (UINT)m_hWnd &&
		 (( pNMHDR->code == TTN_NEEDTEXTA && 
		 pTTTA->uFlags & TTF_IDISHWND ) ||
		 ( pNMHDR->code == TTN_NEEDTEXTW && 
		 pTTTW->uFlags & TTF_IDISHWND ) ) )
		 return FALSE;

	// Get the mouse position
	const MSG* pMessage;
	CPoint pt;
	pMessage = GetCurrentMessage(); // get mouse pos 
	pt = pMessage->pt;
	ScreenToClient( &pt );

	UINT nFlags;
	HTREEITEM hitem =  HitTest( pt, &nFlags ); //Get item pointed by mouse
	if(hitem && (nFlags & TVHT_ONITEM))
		strTipText = GetItemPath(hitem);
	else
		return FALSE;

#ifndef _UNICODE
	if (pNMHDR->code == TTN_NEEDTEXTA)
		lstrcpyn(pTTTA->szText, strTipText, 80);
	else
		_mbstowcsz(pTTTW->szText, strTipText, 80);
#else
	if (pNMHDR->code == TTN_NEEDTEXTA)
		_wcstombsz(pTTTA->szText, strTipText, 80);
	else
		lstrcpyn(pTTTW->szText, strTipText, 80);
#endif
	*pResult = 0;

	return TRUE;    // message was handled
}
void			CTaskUpdate_TreeCtrl::SetItemUID(HTREEITEM iItem, int uid)
{
	DWORD_PTR ptr = GetItemData(iItem);
	itemTreePropsStruct* itemProps = (itemTreePropsStruct*)ptr;
	if(itemProps)
	{
		itemProps->UniqueID = uid;
	}
	else
	{
		itemProps = new itemTreePropsStruct();
		itemProps->strPath = _T("");
		itemProps->strName = _T("");
		itemProps->UniqueID = uid;


		itemPropsPtrList.push_back(itemProps);
		SetItemData(iItem, (LPARAM)itemProps);
	}
}
void			CTaskUpdate_TreeCtrl::SetItemName(HTREEITEM iItem, CString name)
{
	DWORD_PTR ptr = GetItemData(iItem);
	itemTreePropsStruct* itemProps = (itemTreePropsStruct*)ptr;
	if(itemProps)
	{
		itemProps->strName = name;
	}
	else
	{
		itemProps = new itemTreePropsStruct();
		itemProps->strPath = _T("");
		itemProps->strName = name;
		itemProps->UniqueID = -1;

		itemPropsPtrList.push_back(itemProps);
		SetItemData(iItem, (LPARAM)itemProps);
	}
}
void			CTaskUpdate_TreeCtrl::SetItemPath(HTREEITEM iItem, CString path)
{
	DWORD_PTR ptr = GetItemData(iItem);
	itemTreePropsStruct* itemProps = (itemTreePropsStruct*)ptr;
	if(itemProps)
	{
		itemProps->strPath = path;
	}
	else
	{
		itemProps = new itemTreePropsStruct();
		itemProps->strName = _T("");
		itemProps->strPath = path;
		itemProps->UniqueID = -1;

		itemPropsPtrList.push_back(itemProps);
		SetItemData(iItem, (LPARAM)itemProps);
	}
}
CString			CTaskUpdate_TreeCtrl::GetItemName(HTREEITEM iItem)
 {
	 DWORD_PTR ptr = GetItemData(iItem);
	 itemTreePropsStruct* itemProps = (itemTreePropsStruct*)ptr;
	 return itemProps ? itemProps->strName : _T("");
 }
CString			CTaskUpdate_TreeCtrl::GetItemPath(HTREEITEM iItem)
 {
	 DWORD_PTR ptr = GetItemData(iItem);
	 itemTreePropsStruct* itemProps = (itemTreePropsStruct*)ptr;
	 return itemProps ? itemProps->strPath : _T("");
 }
 int			CTaskUpdate_TreeCtrl::GetItemUID(HTREEITEM iItem)const
 {
	 DWORD_PTR ptr = GetItemData(iItem);
	 itemTreePropsStruct* itemProps = (itemTreePropsStruct*)ptr;
	 return itemProps ? itemProps->UniqueID : -1;
 }
HTREEITEM		CTaskUpdate_TreeCtrl::GetItemByUID(int uid)
{
	HTREEITEM hCurrent = GetRootItem();
	while (hCurrent != NULL) 
	{
		if(hCurrent && GetItemUID(hCurrent)==uid)
		return hCurrent;
		hCurrent = GetNextItem(hCurrent, TVGN_NEXT);
	}
	return NULL;
}
// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols
#include "SeavusTaskUpdate_i.h"
#include "dllmain.h"
#include "UpdatesManager.h"


typedef enum { MSP_Unknown=0, MSP_11, MSP_12, MSP_14 } eMSProjectVersion;
extern _ATL_FUNC_INFO OnClickButtonInfo;
extern _ATL_FUNC_INFO OnSimpleEventInfo;
// CConnect


class ATL_NO_VTABLE CConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<AddInDesignerObjects::_IDTExtensibility2, &AddInDesignerObjects::IID__IDTExtensibility2, &AddInDesignerObjects::LIBID_AddInDesignerObjects, 1, 0>,
	public IDispEventSimpleImpl<1, CConnect, &__uuidof(Office::_CommandBarButtonEvents)>,
	public IDispEventSimpleImpl<2, CConnect, &__uuidof(Office::_CommandBarButtonEvents)>,
	public IDispEventSimpleImpl<3, CConnect, &__uuidof(Office::_CommandBarButtonEvents)>,
	public IDispEventSimpleImpl<4, CConnect, &__uuidof(Office::_CommandBarButtonEvents)>,
    public IDispEventSimpleImpl<5, CConnect, &__uuidof(Office::_CommandBarButtonEvents)>,
    public IDispEventSimpleImpl<6, CConnect, &__uuidof(Office::_CommandBarButtonEvents)>,
    public IDispEventSimpleImpl<7, CConnect, &__uuidof(Office::_CommandBarButtonEvents)>,
    public IDispEventSimpleImpl<8, CConnect, &__uuidof(Office::_CommandBarButtonEvents)>,
	public IDispEventSimpleImpl<9,CConnect,&__uuidof(Office::_CommandBarsEvents)>
{
	typedef IDispEventSimpleImpl</*nID =*/ 1,CConnect, &__uuidof(Office::_CommandBarButtonEvents)> CommandButtonMakeCollEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 2,CConnect, &__uuidof(Office::_CommandBarButtonEvents)> CommandButtonGetUpdEvents;
    typedef IDispEventSimpleImpl</*nID =*/ 3,CConnect, &__uuidof(Office::_CommandBarButtonEvents)> CommandButtonOptionsEvents;
    typedef IDispEventSimpleImpl</*nID =*/ 4,CConnect, &__uuidof(Office::_CommandBarButtonEvents)> CommandButtonHelpEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 5,CConnect, &__uuidof(Office::_CommandBarButtonEvents)> CommandMenuMakeCollEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 6,CConnect, &__uuidof(Office::_CommandBarButtonEvents)> CommandMenuGetUpdEvents;
    typedef IDispEventSimpleImpl</*nID =*/ 7,CConnect, &__uuidof(Office::_CommandBarButtonEvents)> CommandMenuOptionsEvents;
    typedef IDispEventSimpleImpl</*nID =*/ 8,CConnect, &__uuidof(Office::_CommandBarButtonEvents)> CommandMenuHelpEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 9,CConnect, &__uuidof(Office::_CommandBarsEvents)> CommandUpdateToolbarEvents;

public:
	CConnect()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_CONNECT)


	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(AddInDesignerObjects::_IDTExtensibility2)
	END_COM_MAP()

	BEGIN_SINK_MAP(CConnect)
		SINK_ENTRY_INFO(1, __uuidof(Office::_CommandBarButtonEvents), /*dispid*/ 0x01, OnClickButtonMakeColl, &OnClickButtonInfo)
		SINK_ENTRY_INFO(2, __uuidof(Office::_CommandBarButtonEvents), /*dispid*/ 0x01, OnClickButtonGetUpd, &OnClickButtonInfo)
        SINK_ENTRY_INFO(3, __uuidof(Office::_CommandBarButtonEvents), /*dispid*/ 0x01, OnClickButtonOptions, &OnClickButtonInfo)
        SINK_ENTRY_INFO(4, __uuidof(Office::_CommandBarButtonEvents), /*dispid*/ 0x01, OnClickButtonHelp, &OnClickButtonInfo)
		SINK_ENTRY_INFO(5, __uuidof(Office::_CommandBarButtonEvents), /*dispid*/ 0x01, OnClickMenuMakeColl, &OnClickButtonInfo)
		SINK_ENTRY_INFO(6, __uuidof(Office::_CommandBarButtonEvents), /*dispid*/ 0x01, OnClickMenuGetUpd, &OnClickButtonInfo)
        SINK_ENTRY_INFO(7, __uuidof(Office::_CommandBarButtonEvents), /*dispid*/ 0x01, OnClickMenuOptions, &OnClickButtonInfo)
        SINK_ENTRY_INFO(8, __uuidof(Office::_CommandBarButtonEvents), /*dispid*/ 0x01, OnClickMenuHelp, &OnClickButtonInfo)
		SINK_ENTRY_INFO(9, __uuidof(Office::_CommandBarsEvents),/*dispinterface*/0x1,OnToolbarUpdate,&OnSimpleEventInfo)
	END_SINK_MAP() 

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct();
	void FinalRelease();
	void __stdcall OnClickButtonMakeColl(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
	void __stdcall OnClickButtonGetUpd(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
    void __stdcall OnClickButtonOptions(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
    void __stdcall OnClickButtonHelp(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
	void __stdcall OnClickMenuMakeColl(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
	void __stdcall OnClickMenuGetUpd(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
    void __stdcall OnClickMenuOptions(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
    void __stdcall OnClickMenuHelp(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
	void __stdcall OnToolbarUpdate();
	// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(IDispatch * Application, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY **custom );
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom );
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom );
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom );

	static CConnect*					pTheConnect;
	CComPtr<IDispatch>					m_pApplication;
	CComPtr<IDispatch>					m_pAddInInstance;
	CComPtr<MSProject::_MSProject>		m_MSPApplication;

	CComPtr <Office::CommandBar>			m_spToolbar;
	CComPtr <Office::_CommandBars>			m_spToolBarCtrl; 
	CComPtr<Office::_CommandBarButton>		m_spButtonMakeColl; 
	CComPtr<Office::_CommandBarButton>		m_spButtonGetUpd; 
    CComPtr<Office::_CommandBarButton>		m_spButtonOptions; 
    CComPtr<Office::_CommandBarButton>		m_spButtonHelp; 
	CComPtr <Office::CommandBarPopup>		m_spMenuCtrl;
	CComPtr<Office::_CommandBarButton>		m_spMenuMakeColl; 
	CComPtr<Office::_CommandBarButton>		m_spMenuGetUpd; 
    CComPtr<Office::_CommandBarButton>		m_spMenuOptions; 
    CComPtr<Office::_CommandBarButton>		m_spMenuHelp; 
	bool m_bConnected;

	void	EnableButton(Office::_CommandBarButton* pButton, bool bEnable=true);
	void	DeleteControls();
	void	EventsUnadvise();
	void	EventsAdvise();
	bool	IsInEditingMode();

    void    ShowMakeCollaborative();
    void    ShowGetUpdates();
    void    ShowOptions();
    void    ShowHelp();

protected:
    template<class T> 
    void showDialog(CUpdatesManager* updatesManager, bool bAskToSave=false);
    template<class T> 
    void showPropertySheet(bool bAskToSave=false);
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)

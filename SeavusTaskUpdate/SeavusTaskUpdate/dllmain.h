// dllmain.h : Declaration of module class.

class CSeavusTaskUpdateModule : public ATL::CAtlDllModuleT< CSeavusTaskUpdateModule >
{
public :
	DECLARE_LIBID(LIBID_SeavusTaskUpdateLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_SEAVUSTASKUPDATE, "{8CB4AD58-C1D9-4D67-A759-8286D3BFA683}")

	HINSTANCE getInstance() const;
	void setInstance(HINSTANCE hInstance);
private:
	HINSTANCE m_hInstance;
};

extern class CSeavusTaskUpdateModule _AtlModule;

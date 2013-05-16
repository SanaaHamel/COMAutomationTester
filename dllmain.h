// dllmain.h : Declaration of module class.

class CCOMAutomationTesterModule : public ATL::CAtlDllModuleT< CCOMAutomationTesterModule >
{
public :
	DECLARE_LIBID(LIBID_COMAutomationTesterLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_COMAUTOMATIONTESTER, "{14859DE3-8A14-4FE2-AA88-F715416477D2}")
};

extern class CCOMAutomationTesterModule _AtlModule;

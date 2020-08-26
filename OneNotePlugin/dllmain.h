// dllmain.h : Declaration of module class.

class CAddinModule : public CAtlDllModuleT< CAddinModule >
{
public :
	DECLARE_LIBID(__uuidof(__KeepOneNoteRunningLib))
};

extern class CAddinModule _AtlModule;

#include "JVMLauncher.h"

static const wchar_t g_kClassNames[] = L"JVMLauncher";

//---------------------------------------------------------------------------//
const WCHAR_T* GetClassNames()
{
	static WCHAR_T* names = 0;
	if (!names)
		::convToShortWchar(&names, g_kClassNames);
	return names;
}

//---------------------------------------------------------------------------//
long GetClassObject(const WCHAR_T* wsName, IComponentBase** pInterface)
{
	if (!*pInterface) {
		*pInterface = new JVMLauncher;
		return (long)*pInterface;
	}
	return 0;
}

long DestroyObject(IComponentBase** pInterface)
{
	if (!*pInterface)
		return -1;

	delete *pInterface;
	*pInterface = nullptr;
	return 0;
}

AppCapabilities SetPlatformCapabilities(const AppCapabilities capabilities) {
	return eAppCapabilitiesLast;
}
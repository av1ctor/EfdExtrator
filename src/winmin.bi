#pragma once

#define null 0

const SPI_GETWORKAREA = &h0030

type HINSTANCE as HINSTANCE__ ptr
type HMODULE as HINSTANCE
type WINBOOL as long
type UINT as ulong
type PVOID as any ptr

type RECT
	left as LONG
	top as LONG
	right as LONG
	bottom as LONG
end type

extern "Windows"
	declare function SetDllDirectory alias "SetDllDirectoryA"(byval lpPathName as zstring ptr) as integer
	declare function LoadLibrary alias "LoadLibraryA"(byval lpLibFileName  as zstring ptr) as HMODULE
	declare function GetProcAddress(byval hModule as HMODULE, byval lpProcName as zstring ptr) as any ptr
	declare sub FreeConsole()
	declare function SystemParametersInfo alias "SystemParametersInfoA"(byval uiAction as UINT, byval uiParam as UINT, byval pvParam as PVOID, byval fWinIni as UINT) as WINBOOL
end extern

#pragma once

#define null 0

type HINSTANCE as HINSTANCE__ ptr
type HMODULE as HINSTANCE

extern "Windows"
	declare function SetDllDirectory alias "SetDllDirectoryA"(byval lpPathName as zstring ptr) as integer
	declare function LoadLibrary alias "LoadLibraryA"(byval lpLibFileName  as zstring ptr) as HMODULE
	declare function GetProcAddress(byval hModule as HMODULE, byval lpProcName as zstring ptr) as any ptr
end extern

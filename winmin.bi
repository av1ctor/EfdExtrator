#pragma once

#define null 0

extern "Windows"
	declare function SetDllDirectory alias "SetDllDirectoryA"(byval lpPathName as zstring ptr) as integer
	declare function LoadLibrary alias "LoadLibraryA"(byval lpLibFileName  as zstring ptr) as long
	declare function GetProcAddress(byval hModule as long, byval lpProcName as zstring ptr) as any ptr
end extern

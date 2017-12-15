#pragma once

#define VB(v) @VarBox(v)

enum VarType
	VT_INT
	VT_UINT
	VT_LNG
	VT_DBL
	VT_STR
end enum

type VarBox
	vtype		as VarType = any
	union
		vi 		as integer
		vui		as uinteger
		vl		as longint
		vd 		as double
		vs 		as zstring ptr
	end union
	
	declare constructor(v as integer)
	declare constructor(v as uinteger)
	declare constructor(v as longint)
	declare constructor(v as double)
	declare constructor(v as const zstring ptr)
	declare destructor()
	
	declare operator cast() as string
end type



#include once "VarBox.bi"

constructor VarBox(v as integer)
	vtype = VT_INT
	vi = v
end constructor

constructor VarBox(v as uinteger)
	vtype = VT_UINT
	vui = v
end constructor

constructor VarBox(v as longint)
	vtype = VT_LNG
	vl = v
end constructor

constructor VarBox(v as double)
	vtype = VT_DBL
	vd = v
end constructor

constructor VarBox(v as const zstring ptr)
	vtype = VT_STR
	vs = allocate(len(*v)+1)
	*vs = *v
end constructor
	
operator VarBox.cast() as string
	select case as const vtype
	case VT_INT
		return str(vi)
	case VT_UINT
		return str(vui)
	case VT_LNG
		return str(vl)
	case VT_DBL
		return str(vd)
	case VT_STR
		return *vs
	end select
end operator

destructor VarBox()
	if vtype = VT_STR then
		deallocate vs
	end if
end destructor
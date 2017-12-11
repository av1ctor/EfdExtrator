#ifndef __HASH_BI__
#define __HASH_BI__

#include once "list.bi"

type HASHITEM
	key			as const zstring ptr			'' shared
	value		as any ptr						'' user data
	prev		as HASHITEM ptr
	next		as HASHITEM ptr
end type

type HASHLIST
	head		as HASHITEM ptr
	tail		as HASHITEM ptr
end type

type THASH
	list		as HASHLIST ptr
	nodes		as integer
	delKey		as boolean
	delVal		as boolean
	allocKey	as boolean
end type

declare sub hashInit _
	( _
		byval hash as THASH ptr, _
		byval nodes as integer, _
		byval delKey as boolean = false, _		'' delete key when hashEnd()/hashDel() is called?
		byval delVal as boolean = false, _		'' deallocate value when hashEnd()/hashDel() is called?
		byval allocKey as boolean = false _		'' allocate and copy key when hashAdd() is called?
	)

declare sub hashEnd(byval hash as THASH ptr)

declare function hashHash _
	( _
		byval key as const zstring ptr _
	) as uinteger

declare function hashLookup _
	( _
		byval hash as THASH ptr, _
		byval key as zstring ptr _
	) as any ptr

declare function hashLookupEx _
	( _
		byval hash as THASH ptr, _
		byval key as const zstring ptr, _
		byval index as uinteger _
	) as any ptr

declare function hashAdd _
	( _
		byval hash as THASH ptr, _
		byval key as const zstring ptr, _
		byval value as any ptr, _
		byval index as uinteger = cuint( -1 ) _
	) as HASHITEM ptr

declare sub hashDel _
	( _
		byval hash as THASH ptr, _
		byval item as HASHITEM ptr, _
		byval index as uinteger _
	)

#endif '' __HASH_BI__
#ifndef __HASH_BI__
#define __HASH_BI__

#include once "list.bi"

type HASHITEM
	key			as const zstring ptr			'' shared if allocKey = false
	value		as any ptr						'' user data
	prev		as HASHITEM ptr
	next		as HASHITEM ptr
end type

type HASHLIST
	head		as HASHITEM ptr
	tail		as HASHITEM ptr
end type

type THash
	declare sub init(nodes as integer, delKey as boolean = false, delVal as boolean = false, allocKey as boolean = false)
	declare sub end_()
	declare function lookup(key as zstring ptr) as any ptr
	declare function lookupEx(key as const zstring ptr, index as uinteger) as any ptr
	declare function add(key as const zstring ptr, value as any ptr, index as uinteger = cuint( -1 )) as HASHITEM ptr
	declare sub del(item as HASHITEM ptr, index as uinteger)

private:	
	declare function hash(key as const zstring ptr) as uinteger

	list		as HASHLIST ptr
	nodes		as integer
	delKey		as boolean
	delVal		as boolean
	allocKey	as boolean
end type

#endif '' __HASH_BI__
#ifndef __HASH_BI__
#define __HASH_BI__

#include once "list.bi"

type HashItem
	key			as const zstring ptr			'' shared if allocKey = false
	value		as any ptr						'' user data
	prev		as HashItem ptr
	next		as HashItem ptr
end type

type HashChain
	head		as HashItem ptr
	tail		as HashItem ptr
end type

type THash
	declare sub init(nodes as integer, delKey as boolean = false, delVal as boolean = false, allocKey as boolean = false)
	declare sub end_()
	declare function lookup(key as const zstring ptr) as any ptr
	declare function lookupEx(key as const zstring ptr, index as uinteger) as any ptr
	declare operator [](key as integer) as any ptr
	declare operator [](key as double) as any ptr
	declare operator [](key as const zstring ptr) as any ptr
	declare function add(key as integer, value as any ptr) as HashItem ptr
	declare function add(key as double, value as any ptr) as HashItem ptr
	declare function add(key as const zstring ptr, value as any ptr) as HashItem ptr
	declare sub del(item as HashItem ptr, index as uinteger)

private:	
	declare function hash(key as const zstring ptr) as uinteger

	chain		as HashChain ptr
	nodes		as integer
	delKey		as boolean
	delVal		as boolean
	allocKey	as boolean
end type

#endif '' __HASH_BI__
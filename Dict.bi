#ifndef __Dict_BI__
#define __Dict_BI__

#include once "list.bi"

type DictItem
	key			as const zstring ptr			'' shared if allocKey = false
	value		as any ptr						'' user data
	prev		as DictItem ptr
	next		as DictItem ptr
end type

type DictChain
	head		as DictItem ptr
	tail		as DictItem ptr
end type

type TDict
	declare sub init(nodes as integer, delKey as boolean = false, delVal as boolean = false, allocKey as boolean = false)
	declare sub end_()
	declare function lookup(key as const zstring ptr) as any ptr
	declare function lookupEx(key as const zstring ptr, index as uinteger) as any ptr
	declare operator [](key as integer) as any ptr
	declare operator [](key as double) as any ptr
	declare operator [](key as const zstring ptr) as any ptr
	declare function add(key as integer, value as any ptr) as DictItem ptr
	declare function add(key as double, value as any ptr) as DictItem ptr
	declare function add(key as const zstring ptr, value as any ptr) as DictItem ptr
	declare sub del(item as DictItem ptr, index as uinteger)

private:	
	declare function hash(key as const zstring ptr) as uinteger

	chain		as DictChain ptr
	nodes		as integer
	delKey		as boolean
	delVal		as boolean
	allocKey	as boolean
end type

#endif '' __Dict_BI__
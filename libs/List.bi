#ifndef __LIST_BI__
#define __LIST_BI__

#define NULL 0

type TListNode
	prev		as TListNode ptr
	next		as TListNode ptr
end type

type TListTb
	next		as TListTb ptr
	nodetb		as any ptr
	nodes		as integer
end type

type onCompareCb as function(key as any ptr, node as any ptr) as boolean
type onCompareExCb as function(key as any ptr, extra as any ptr, node as any ptr) as boolean
type onCompareLongCb as function(key as long, node as any ptr) as boolean
type onCompareDoubleCb as function(key as double, node as any ptr) as boolean
type onCompareStringCb as function(key as zstring ptr, node as any ptr) as boolean

type TList
	declare constructor(nodes as integer, nodeLen as integer, clearNodes as boolean = true)
	declare destructor()
	declare function add() as any ptr
	declare function addOrdAsc(key as any ptr, extra as any ptr, cmpFunc as onCompareExCb) as any ptr
	declare function addOrdAsc(key as zstring ptr, cmpFunc as onCompareStringCb) as any ptr
	declare function addOrdAsc(key as long, cmpFunc as onCompareLongCb) as any ptr
	declare function addOrdAsc(key as double, cmpFunc as onCompareDoubleCb) as any ptr
	declare sub del(node as any ptr)
	declare property head() as any ptr
	declare property tail() as any ptr
	declare property prev(node as any ptr) as any ptr
	declare property next_(node as any ptr) as any ptr

private:
	declare sub allocTB(nodes as integer)
	
	tbhead		as TListTb ptr
	tbtail		as TListTb ptr
	nodes 		as integer
	nodeLen		as integer
	clearNodes	as boolean
	fhead		as TListNode ptr					'' free list
	ahead		as TListNode ptr					'' allocated list
	atail		as TListNode ptr					'' /
end type


#endif '' __LIST_BI__
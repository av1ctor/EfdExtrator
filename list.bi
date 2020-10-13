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

type TList
	declare constructor(nodes as integer, nodeLen as integer, clearNodes as boolean = true)
	declare destructor()
	declare function add() as any ptr
	declare function addOrdAsc(key as any ptr, cmpFunc as function(key as any ptr, node as any ptr) as boolean) as any ptr
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
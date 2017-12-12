#ifndef __LIST_BI__
#define __LIST_BI__

#define NULL 0

type TLISTNODE
	prev	as TLISTNODE ptr
	next	as TLISTNODE ptr
end type

type TLISTTB
	next	as TLISTTB ptr
	nodetb	as any ptr
	nodes	as integer
end type

type TList
	declare sub init(nodes as integer, nodelen as integer)
	declare sub end_()
	declare function add() as any ptr
	declare sub del(node as any ptr)
	declare property head() as any ptr
	declare property tail() as any ptr
	declare property prev(node as any ptr) as any ptr
	declare property next_(node as any ptr) as any ptr

private:
	declare sub allocTB(nodes as integer)
	
	tbhead	as TLISTTB ptr
	tbtail	as TLISTTB ptr
	nodes 	as integer
	nodelen	as integer
	fhead	as TLISTNODE ptr					'' free list
	ahead	as TLISTNODE ptr					'' allocated list
	atail	as TLISTNODE ptr					'' /
end type


#endif '' __LIST_BI__
'' Generic double-linked list for FreeBASIC
'' Copyright 2017 by Andre Victor (av1ctortv[@]gmail.com)
'' Licensed under GNU GPL-2.0-or-above

#include once "list.bi"

'':::::
constructor TList(nodes as integer, nodelen as integer, clearNodes as boolean)

	'' fill ctrl struct
	this.tbhead = NULL
	this.tbtail = NULL
	this.nodes	= 0
	this.nodeLen = nodelen + len( TListNode )
	this.ahead = NULL
	this.atail = NULL
	this.clearNodes = clearNodes

	'' allocate the initial pool
	allocTB( nodes )

end constructor

'':::::
destructor TList()
	'' for each pool, free the mem block and the pool ctrl struct
	var tb = this.tbhead
	do while( tb <> NULL )
		var nxt = tb->next
		deallocate( tb->nodetb )
		deallocate( tb )
		tb = nxt
	loop

	this.tbhead = NULL
	this.tbtail = NULL
	this.nodes	= 0
end destructor

'':::::
sub TList.allocTB(nodes as integer)

	assert(nodes >= 1)

	'' allocate the pool
	dim as TListNode ptr nodetb
	if clearNodes then
		nodetb = cast(TListNode ptr, callocate( nodes * this.nodelen ))
	else
		nodetb = cast(TListNode ptr, allocate( nodes * this.nodelen ))
	end if

	'' and the pool ctrl struct
	var tb = cast(TListTb ptr, allocate( len( TListTb ) ))

	'' add the ctrl struct to pool list
	if( this.tbhead = NULL ) then
		this.tbhead = tb
	end if
	if( this.tbtail <> NULL ) then
		this.tbtail->next = tb
	end if
	this.tbtail = tb

	tb->next = NULL
	tb->nodetb = nodetb
	tb->nodes = nodes

	'' add new nodes to the free list
	this.fhead = nodetb
	this.nodes += nodes

	var prv = cast(TListNode ptr, NULL)
	var node = this.fhead

	for i as integer = 1 to nodes-1
		node->prev	= prv
		node->next	= cast( TListNode ptr, cast( byte ptr, node ) + this.nodelen )

		prv = node
		node = node->next
	next

	node->prev = prv
	node->next = NULL

end sub

'':::::
function TList.add() as any ptr

	'' alloc new node list if there are no free nodes
	if( this.fhead = NULL ) Then
		allocTB( cunsg(this.nodes) \ 4 )
	end if

	'' take from free list
	var node = this.fhead
	this.fhead = node->next

	'' add to used list
	if( this.atail <> NULL ) then
		this.atail->next = node
	else
		this.ahead = node
	end If

	node->prev = this.atail
	node->next = NULL

	this.atail = node

	function = cast( byte ptr, node ) + len( TListNode )

end function

private function long_cmpFunc(key as any ptr, extra as any ptr, node as any ptr) as boolean
	return cast(onCompareLongCb, extra)(*cast(long ptr, key), node)
end function

function TList.addOrdAsc(key as long, cmpFunc as onCompareLongCb) as any ptr
	var pkey = @key
	return addOrdAsc(pkey, cmpFunc, @long_cmpFunc)
end function

private function double_cmpFunc(key as any ptr, extra as any ptr, node as any ptr) as boolean
	return cast(onCompareDoubleCb, extra)(*cast(double ptr, key), node)
end function

function TList.addOrdAsc(key as double, cmpFunc as onCompareDoubleCb) as any ptr
	var pkey = @key
	return addOrdAsc(pkey, cmpFunc, @double_cmpFunc)
end function

private function string_cmpFunc(key as any ptr, extra as any ptr, node as any ptr) as boolean
	return cast(onCompareStringCb, extra)(*cast(zstring ptr ptr, key), node)
end function

function TList.addOrdAsc(key as zstring ptr, cmpFunc as onCompareStringCb) as any ptr
	var pkey = @key
	return addOrdAsc(pkey, cmpFunc, @string_cmpFunc)
end function

'':::::
function TList.addOrdAsc(key as any ptr, extra as any ptr, cmpFunc as onCompareExCb) as any ptr

	'' alloc new node list if there are no free nodes
	if( this.fhead = NULL ) Then
		allocTB( cunsg(this.nodes) \ 4 )
	end if

	'' take from free list
	var node = this.fhead
	this.fhead = node->next
	
	'' add to used list
	if( this.ahead = null ) then
		this.ahead = node
		this.atail = node
		node->prev = null
		node->next = null
	else
		var n = this.ahead
		var p = cast(TListNode ptr, null)
		do
			if( cmpFunc(key, extra, cast(any ptr, cast( byte ptr, n ) + len( TListNode ))) ) then
				node->next = n
				node->prev = n->prev
				n->prev = node
				if p <> null then
					p->next = node
				else
					this.ahead = node
				end if
				exit do
			end if
			
			p = n
			n = n->next
		loop until( n = null )
		
		if( n = null ) then
			this.atail->next = node
			node->prev = this.atail
			node->next = null
			this.atail = node
		end if
	end if

	function = cast( byte ptr, node ) + len( TListNode )

end function


'':::::
sub TList.del(node_ as any ptr)

	if( node_ = NULL ) then
		exit sub
	end if

	var node = cast( TListNode ptr, cast( byte ptr, node_ ) - len( TListNode ) )

	'' remove from used list
	var prv = node->prev
	var nxt = node->next
	if( prv <> NULL ) then
		prv->next = nxt
	else
		this.ahead = nxt
	end If

	if( nxt <> NULL ) then
		nxt->prev = prv
	else
		this.atail = prv
	end If

	'' add to free list
	node->next = this.fhead
	this.fhead = node

	'' node can contain strings descriptors, so, erase it..
	if clearNodes then
		clear( byval node_, 0, this.nodelen - len( TListNode ) )
	end if

end sub

'':::::
property TList.head( ) as any ptr

	if( this.ahead = NULL ) then
		return NULL
	else
		return cast( byte ptr, this.ahead ) + len( TListNode )
	end if

end property

'':::::
property TList.tail() as any ptr

	if( this.atail = NULL ) then
		return NULL
	else
		return cast( byte ptr, this.atail ) + len( TListNode )
	end if

end property

'':::::
property TList.prev(node as any ptr) as any ptr

	assert( node <> NULL )

	var prv = cast( TListNode ptr, cast( byte ptr, node ) - len( TListNode ) )->prev

	if( prv = NULL ) then
		return NULL
	else
		return cast( byte ptr, prv ) + len( TListNode )
	end if

end property

'':::::
property TList.next_(node as any ptr) as any ptr

	assert( node <> NULL )

	var nxt = cast( TListNode ptr, cast( byte ptr, node ) - len( TListNode ) )->next

	if( nxt = NULL ) then
		return NULL
	else
		return cast( byte ptr, nxt ) + len( TListNode )
	end if

end property


'' generic double-linked list

#include once "list.bi"

'':::::
sub TList.init(nodes as integer, nodelen as integer )

	'' fill ctrl struct
	this.tbhead = NULL
	this.tbtail = NULL
	this.nodes	= 0
	this.nodelen = nodelen + len( TLISTNODE )
	this.ahead = NULL
	this.atail = NULL

	'' allocate the initial pool
	allocTB( nodes )

end sub

'':::::
sub TList.end_()
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
end sub

'':::::
sub TList.allocTB(nodes as integer)

	assert(nodes >= 1)

	'' allocate the pool
	var nodetb = cast(TLISTNODE ptr, callocate( nodes * this.nodelen ))

	'' and the pool ctrl struct
	var tb = cast(TLISTTB ptr, allocate( len( TLISTTB ) ))

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

	var prv = cast(TLISTNODE ptr, NULL)
	var node = this.fhead

	for i as integer = 1 to nodes-1
		node->prev	= prv
		node->next	= cast( TLISTNODE ptr, cast( byte ptr, node ) + this.nodelen )

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
	var t = this.atail
	this.atail = node
	if( t <> NULL ) then
		t->next = node
	else
		this.ahead = node
	end If

	node->prev = atail
	node->next = NULL

	function = cast( byte ptr, node ) + len( TLISTNODE )

end function

'':::::
sub TList.del(node_ as any ptr)

	if( node_ = NULL ) then
		exit sub
	end if

	var node = cast( TLISTNODE ptr, cast( byte ptr, node_ ) - len( TLISTNODE ) )

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
	clear( node_, 0, this.nodelen - len( TLISTNODE ) )

end sub

'':::::
property TList.head( ) as any ptr

	if( this.ahead = NULL ) then
		return NULL
	else
		return cast( byte ptr, this.ahead ) + len( TLISTNODE )
	end if

end property

'':::::
property TList.tail() as any ptr

	if( this.atail = NULL ) then
		return NULL
	else
		return cast( byte ptr, this.atail ) + len( TLISTNODE )
	end if

end property

'':::::
property TList.prev(node as any ptr) as any ptr

	assert( node <> NULL )

	var prv = cast( TLISTNODE ptr, cast( byte ptr, node ) - len( TLISTNODE ) )->prev

	if( prv = NULL ) then
		return NULL
	else
		return cast( byte ptr, prv ) + len( TLISTNODE )
	end if

end property

'':::::
property TList.next_(node as any ptr) as any ptr

	assert( node <> NULL )

	var nxt = cast( TLISTNODE ptr, cast( byte ptr, node ) - len( TLISTNODE ) )->next

	if( nxt = NULL ) then
		return NULL
	else
		return cast( byte ptr, nxt ) + len( TLISTNODE )
	end if

end property


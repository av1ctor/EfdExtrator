'' generic hash tables

#include once "hash.bi"

type HASHITEMPOOL
	as integer refcount
	as TList list
end type

dim shared as HASHITEMPOOL itempool

''::::::
private sub lazyInit()
	itempool.refcount += 1
	if (itempool.refcount > 1) then
		exit sub
	end if

	const INITIAL_ITEMS = 8096

	'' allocate the initial item list pool
	itempool.list.init(INITIAL_ITEMS, sizeof(HASHITEM))
end sub

''::::::
private sub lazyEnd()
	itempool.refcount -= 1
	if (itempool.refcount > 0) then
		exit sub
	end if

	itempool.list.end_()
end sub

''::::::
private function hNewItem(list as HASHLIST ptr) as HASHITEM ptr

	'' add a new node
	var item = cast(HASHITEM ptr, itempool.list.add( ))

	'' add it to the internal linked-list
	if( list->tail <> NULL ) then
		list->tail->next = item
	else
		list->head = item
	end if

	item->prev = list->tail
	item->next = NULL

	list->tail = item

	function = item

end function

''::::::
private sub hDelItem(list as HASHLIST ptr, item as HASHITEM ptr)

	''
	if( item = NULL ) Then
		exit sub
	end If

	'' remove from internal linked-list
	var prv  = item->prev
	var nxt  = item->next
	if( prv <> NULL ) then
		prv->next = nxt
	else
		list->head = nxt
	end If

	if( nxt <> NULL ) then
		nxt->prev = prv
	else
		list->tail = prv
	end if

	'' remove node
	itempool.list.del( item )

end sub

''::::::
sub THash.init(nodes as integer, delKey as boolean, delVal as boolean, allocKey as boolean)

	lazyInit()

	'' allocate a fixed list of internal linked-lists
	this.list = callocate( nodes * len( HASHLIST ) )
	this.nodes = nodes
	this.delKey = delKey or allocKey
	this.delVal = delVal
	this.allocKey = allocKey

end sub

''::::::
sub THash.end_()

    var list_ = this.list

	for i as integer = 0 to this.nodes-1
		var item = list_->head
		do while( item <> NULL )
			var nxt = item->next

			if this.delVal then
				if item->value <> null then
					deallocate( item->value )
					item->value = null
				end if
			end if
			if( this.delKey ) then
				deallocate( item->key )
			end if
			item->key = NULL
			hDelItem( list_, item )

			item = nxt
		loop

		list_ += 1
	next

	deallocate( this.list )
	this.list = NULL

	lazyEnd()

end sub

''::::::
function THash.hash(key as const zstring ptr) as uinteger
	dim as uinteger index = 0
	do while (key[0])
		index = key[0] + (index shl 5) - index
		key += 1
	loop
	return index
end function

''::::::
function THash.lookupEx(key as const zstring ptr, index as uinteger ) as any ptr

    index mod= this.nodes

	'' get the start of list
	var item = (@this.list[index])->head
	if( item = NULL ) then
		return NULL
	end if

	'' loop until end of list or if item was found
	do while( item <> NULL )
		if( *item->key = *key ) then
			return item->value
		end if
		item = item->next
	loop
	
	function = null

end function

''::::::
function THash.lookup(key as zstring ptr) as any ptr

    function = lookupEx( key, hash( key ) )

end function

''::::::
function THash.add(key as const zstring ptr, value as any ptr, index as uinteger) as HASHITEM ptr

	'' calc hash?
	if( index = cuint( -1 ) ) then
		index = hash( key )
	end if

    index mod= this.nodes

    '' allocate a new node
    var item = hNewItem( @this.list[index] )

    if( item = NULL ) then
    	return null
	end if

    '' fill node
    if this.allocKey then
		var key2 = cast(zstring ptr, allocate(len(*key)+1))
		*key2 = *key
		key = key2
	end if
	item->key = key
    item->value = value

    function = item
end function

''::::::
sub THash.del(item as HASHITEM ptr, index as uinteger)

	if( item = NULL ) then
		exit sub
	end if

	index mod= this.nodes

	''
	if( this.delKey ) then
		deallocate( item->key )
	end if
	item->key = NULL

	if( this.delVal ) then
		deallocate( item->value )
	end if
	item->value = NULL

	hDelItem( @this.list[index], item )

end sub


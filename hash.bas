'' generic hash tables

#include once "hash.bi"

type HASHITEMPOOL
	as integer refcount
	as TLIST list
end type


declare function 	hashNewItem	( byval list as HASHLIST ptr ) as HASHITEM ptr
declare sub 		hashDelItem	( byval list as HASHLIST ptr, _
								  byval item as HASHITEM ptr )

dim shared as HASHITEMPOOL itempool

''::::::
private sub lazyInit()
	itempool.refcount += 1
	if (itempool.refcount > 1) then
		exit sub
	end if

	const INITIAL_ITEMS = 8096

	'' allocate the initial item list pool
	listInit(@itempool.list, INITIAL_ITEMS, sizeof(HASHITEM), LIST_FLAGS_NOCLEAR)
end sub

''::::::
private sub lazyEnd()
	itempool.refcount -= 1
	if (itempool.refcount > 0) then
		exit sub
	end if

	listEnd(@itempool.list)
end sub

''::::::
sub hashInit _
	( _
		byval hash as THASH ptr, _
		byval nodes as integer, _
		byval delKey as boolean, _
		byval delVal as boolean, _
		byval allocKey as boolean	_
	)

	lazyInit()

	'' allocate a fixed list of internal linked-lists
	hash->list = callocate( nodes * len( HASHLIST ) )
	hash->nodes = nodes
	hash->delKey = delKey or allocKey
	hash->delVal = delVal
	hash->allocKey = allocKey

end sub

''::::::
sub hashEnd(byval hash as THASH ptr)

    var list = hash->list

	for i as integer = 0 to hash->nodes-1
		var item = list->head
		do while( item <> NULL )
			var nxt = item->next

			if hash->delVal then
				if item->value <> null then
					deallocate( item->value )
					item->value = null
				end if
			end if
			if( hash->delKey ) then
				deallocate( item->key )
			end if
			item->key = NULL
			hashDelItem( list, item )

			item = nxt
		loop

		list += 1
	next

	deallocate( hash->list )
	hash->list = NULL

	lazyEnd()

end sub

''::::::
function hashHash(byval key as const zstring ptr) as uinteger
	dim as uinteger index = 0
	do while (key[0])
		index = key[0] + (index shl 5) - index
		key += 1
	loop
	return index
end function

''::::::
function hashLookupEx _
	( _
		byval hash as THASH ptr, _
		byval key as const zstring ptr, _
		byval index as uinteger _
	) as any ptr

    index mod= hash->nodes

	'' get the start of list
	var list = @hash->list[index]
	var item = list->head
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
function hashLookup _
	( _
		byval hash as THASH ptr, _
		byval key as zstring ptr _
	) as any ptr

    function = hashLookupEx( hash, key, hashHash( key ) )

end function

''::::::
private function hashNewItem _
	( _
		byval list as HASHLIST ptr _
	) as HASHITEM ptr

	'' add a new node
	var item = cast(HASHITEM ptr, listNewNode( @itempool.list ))

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
private sub hashDelItem _
	( _
		byval list as HASHLIST ptr, _
		byval item as HASHITEM ptr _
	)

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
	listDelNode( @itempool.list, item )

end sub

''::::::
function hashAdd _
	( _
		byval hash as THASH ptr, _
		byval key as const zstring ptr, _
		byval value as any ptr, _
		byval index as uinteger _
	) as HASHITEM ptr

	'' calc hash?
	if( index = cuint( -1 ) ) then
		index = hashHash( key )
	end if

    index mod= hash->nodes

    '' allocate a new node
    var item = hashNewItem( @hash->list[index] )

    if( item = NULL ) then
    	return null
	end if

    '' fill node
    if hash->allocKey then
		var key2 = cast(zstring ptr, allocate(len(*key)+1))
		*key2 = *key
		key = key2
	end if
	item->key = key
    item->value = value

    function = item
end function

''::::::
sub hashDel _
	( _
		byval hash as THASH ptr, _
		byval item as HASHITEM ptr, _
		byval index as uinteger _
	)

	if( item = NULL ) then
		exit sub
	end if

	index mod= hash->nodes

	'' get start of list
	var list = @hash->list[index]

	''
	if( hash->delKey ) then
		deallocate( item->key )
	end if
	item->key = NULL

	if( hash->delVal ) then
		deallocate( item->value )
	end if
	item->value = NULL

	hashDelItem( list, item )

end sub


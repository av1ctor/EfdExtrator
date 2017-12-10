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


private sub lazyInit()
	itempool.refcount += 1
	if (itempool.refcount > 1) then
		exit sub
	end if

	const INITIAL_ITEMS = 8096

	'' allocate the initial item list pool
	listInit(@itempool.list, INITIAL_ITEMS, sizeof(HASHITEM), LIST_FLAGS_NOCLEAR)
end sub

private sub lazyEnd()
	itempool.refcount -= 1
	if (itempool.refcount > 0) then
		exit sub
	end if

	listEnd(@itempool.list)
end sub

sub hashInit _
	( _
		byval hash as THASH ptr, _
		byval nodes as integer, _
		byval delstr as boolean, _
		byval delvalue as boolean	_
	)

	lazyInit()

	'' allocate a fixed list of internal linked-lists
	hash->list = callocate( nodes * len( HASHLIST ) )
	hash->nodes = nodes
	hash->delstr = delstr
	hash->delval = delvalue

end sub

sub hashEnd(byval hash as THASH ptr)

    dim as integer i = any
    dim as HASHITEM ptr item = any, nxt = any
    dim as HASHLIST ptr list = any

    '' for each item on each list, deallocate it and the name string
    list = hash->list

	for i = 0 to hash->nodes-1
		item = list->head
		do while( item <> NULL )
			nxt = item->next

			if hash->delval then
				if item->data <> null then
					deallocate( item->data )
					item->data = null
				end if
			end if
			if( hash->delstr ) then
				deallocate( item->name )
			end if
			item->name = NULL
			hashDelItem( list, item )

			item = nxt
		loop

		list += 1
	next

	deallocate( hash->list )
	hash->list = NULL

	lazyEnd()

end sub

function hashHash(byval s as const zstring ptr) as uinteger
	dim as uinteger index = 0
	while (s[0])
		index = s[0] + (index shl 5) - index
		s += 1
	wend
	return index
end function

''::::::
function hashLookupEx _
	( _
		byval hash as THASH ptr, _
		byval symbol as const zstring ptr, _
		byval index as uinteger _
	) as any ptr

    dim as HASHITEM ptr item = any
    dim as HASHLIST ptr list = any

    function = NULL

    index mod= hash->nodes

	'' get the start of list
	list = @hash->list[index]
	item = list->head
	if( item = NULL ) then
		exit function
	end if

	'' loop until end of list or if item was found
	do while( item <> NULL )
		if( *item->name = *symbol ) then
			return item->data
		end if
		item = item->next
	loop

end function

''::::::
function hashLookup _
	( _
		byval hash as THASH ptr, _
		byval symbol as zstring ptr _
	) as any ptr

    function = hashLookupEx( hash, symbol, hashHash( symbol ) )

end function

''::::::
private function hashNewItem _
	( _
		byval list as HASHLIST ptr _
	) as HASHITEM ptr

	dim as HASHITEM ptr item = any

	'' add a new node
	item = listNewNode( @itempool.list )

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

	dim as HASHITEM ptr prv = any, nxt = any

	''
	if( item = NULL ) Then
		exit sub
	end If

	'' remove from internal linked-list
	prv  = item->prev
	nxt  = item->next
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
		byval symbol as const zstring ptr, _
		byval userdata as any ptr, _
		byval index as uinteger _
	) as HASHITEM ptr

    dim as HASHITEM ptr item = any

	'' calc hash?
	if( index = cuint( -1 ) ) then
		index = hashHash( symbol )
	end if

    index mod= hash->nodes

    '' allocate a new node
    item = hashNewItem( @hash->list[index] )

    function = item
    if( item = NULL ) then
    	exit function
	end if

    '' fill node
    item->name = symbol
    item->data = userdata

end function

''::::::
sub hashDel _
	( _
		byval hash as THASH ptr, _
		byval item as HASHITEM ptr, _
		byval index as uinteger _
	)

    dim as HASHLIST ptr list = any

	if( item = NULL ) then
		exit sub
	end if

	index mod= hash->nodes

	'' get start of list
	list = @hash->list[index]

	''
	if( hash->delstr ) then
		deallocate( item->name )
	end if
	item->name = NULL

	item->data = NULL

	hashDelItem( list, item )

end sub


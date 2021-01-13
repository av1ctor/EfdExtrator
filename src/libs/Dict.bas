'' Generic Dictionary for FreeBASIC
'' Copyright 2017 by Andre Victor (av1ctortv[@]gmail.com)
'' Licensed under GNU GPL-2.0-or-above

#include once "Dict.bi"

type DictItemPOOL
	refcount		as integer
	list			as TList ptr
end type

dim shared as DictItemPOOL itempool

''::::::
private sub lazyInit()
	itempool.refcount += 1
	if (itempool.refcount > 1) then
		exit sub
	end if

	const INITIAL_ITEMS = 8096

	'' allocate the initial item list pool
	itempool.list = new TList(INITIAL_ITEMS, sizeof(DictItem), false)
end sub

''::::::
private sub lazyEnd()
	itempool.refcount -= 1
	if (itempool.refcount > 0) then
		exit sub
	end if

	delete itempool.list
end sub

''::::::
private function hNewItem(chain_ as DictChain ptr) as DictItem ptr

	'' add a new node
	var item = cast(DictItem ptr, itempool.list->add( ))

	'' add it to the internal linked-list
	if( chain_->tail <> NULL ) then
		chain_->tail->next = item
	else
		chain_->head = item
	end if

	item->prev = chain_->tail
	item->next = NULL

	chain_->tail = item

	function = item

end function

''::::::
private sub hDelItem(chain_ as DictChain ptr, item as DictItem ptr)

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
		chain_->head = nxt
	end If

	if( nxt <> NULL ) then
		nxt->prev = prv
	else
		chain_->tail = prv
	end if

	'' remove node
	itempool.list->del( item )

end sub

''::::::
constructor TDict(nodes as integer, delKey as boolean, delVal as boolean, allocKey as boolean)

	lazyInit()

	'' allocate a fixed list of internal linked-lists
	this.chain = callocate( nodes * len( DictChain ) )
	this.nodes = nodes
	this.delKey = delKey or allocKey
	this.delVal = delVal
	this.allocKey = allocKey

end constructor

''::::::
destructor TDict()

    var list_ = this.chain

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

	deallocate( this.chain )
	this.chain = NULL

	lazyEnd()

end destructor

''::::::
function TDict.hash(key as const zstring ptr) as uinteger
	dim as uinteger index = 0
	do while (key[0])
		index = key[0] + (index shl 5) - index
		key += 1
	loop
	return index
end function

''::::::
function TDict.lookupEx(key as const zstring ptr, index as uinteger ) as any ptr

    index mod= this.nodes

	'' get the start of list
	var item = (@this.chain[index])->head
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
function TDict.lookup(key as integer) as any ptr
	var k = str(key)
	function = lookupEx( strptr(k), hash( strptr(k) ) )
end function

''::::::
function TDict.lookup(key as double) as any ptr
	var k = str(key)
	function = lookupEx( strptr(k), hash( strptr(k) ) )
end function

''::::::
function TDict.lookup(key as const zstring ptr) as any ptr
    function = lookupEx( key, hash( key ) )
end function

''::::::
operator TDict.[](key as integer) as any ptr
	operator = lookup( key )
end operator

''::::::
operator TDict.[](key as double) as any ptr
	operator = lookup( key )
end operator

''::::::
operator TDict.[](key as const zstring ptr) as any ptr
	operator = lookupEx( key, hash( key ) )
end operator

''::::::
function TDict.add(key as const zstring ptr, value as any ptr) as DictItem ptr

	var index = hash( key )

    index mod= this.nodes

    '' allocate a new node
    var item = hNewItem( @this.chain[index] )

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
function TDict.add(key as integer, value as any ptr) as DictItem ptr
	function = add(str(key), value)
end function

''::::::
function TDict.add(key as double, value as any ptr) as DictItem ptr
	function = add(str(key), value)
end function

''::::::
sub TDict.del(item as DictItem ptr, index as uinteger)

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

	hDelItem( @this.chain[index], item )

end sub


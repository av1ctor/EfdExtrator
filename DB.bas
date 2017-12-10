#include once "DB.bi" 
#include once "list.bi" 

''''''''
function TDb.open(fileName as string) as boolean
	
	if sqlite3_open( fileName, @instance ) then 
  		errMsg = *sqlite3_errmsg( instance )
		sqlite3_close( instance ) 
		return false
	end if 
	
	errMsg = ""
	return true
	
end function

''''''''
sub TDb.close()
	if instance <> null then
		sqlite3_close( instance ) 
		instance = null
		errMsg = ""
	end if
end sub

''''''''	
private function callback cdecl _
	( _
		byval rset__ as any ptr, _
		byval argc as integer, _
		byval argv as zstring ptr ptr, _
		byval colName as zstring ptr ptr _
	) as integer
	
	var rset_ = cast(TRSet ptr, rset__)
	
	var row = rset_->newRow()
  
	for i as integer = 0 to argc - 1
		dim as zstring ptr text = null
		if( argv[i] <> 0 ) then
			if *argv[i] <> 0 then 
				text = argv[i]
			end if
		end if
				
		row->newColumn(colName[i], text)
	next 
	
	function = 0
   
end function 
	
''''''''	
function TDb.exec(query as string) as TRSet ptr

	var rset_ = new TRSet
	
	dim as zstring ptr errMsg_ = null
	if sqlite3_exec( instance, query, @callback, rset_, @errMsg_ ) <> SQLITE_OK then 
		delete rset_
		errMsg = *errMsg_
		return null
	end if 
	
	return rset_

end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
constructor TRSet()
	listInit(@rows, 10, len(TRSetRow))
	currRow = null
end constructor	
	
''''''''
destructor TRSet()
	var r = cast(TRSetRow ptr, listGetHead(@rows))
	do while r <> null
		r->destructor
		r = listGetNext(r)
	loop
	
	listEnd(@rows)
	currRow = null
end destructor

''''''''
function TRSet.hasNext() as boolean
	return currRow <> null
end function

''''''''
sub TRSet.next_() 
	if currRow <> null then
		currRow = listGetNext(currRow)
	end if
end sub

''''''''
property TRSet.row() as TRSetRow ptr
	return currRow
end property

''''''''
function TRSet.newRow() as TRSetRow ptr
	var p = listNewNode(@rows)
	var r = new (p) TRSetRow()
	if currRow = null then
		currRow = r
	end if
	return r
end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
constructor TRSetRow()
	hashInit(@columns, 10, true, true)
end constructor	
	
''''''''
destructor TRSetRow()
	hashEnd(@columns)
end destructor

''''''''
sub TRSetRow.newColumn(name_ as zstring ptr, value as zstring ptr)
	if hashLookup( @columns, name_ ) = null then
		var name2 = cast(zstring ptr, allocate(len(*name_)+1))
		*name2 = *name_
		
		dim as zstring ptr value2 = null
		if value <> null then
			value2 = cast(zstring ptr, allocate(len(*value)+1))	
			*value2 = *value
		end if
		
		hashAdd( @columns, name2, value2 )
	end if
end sub

''''''''
operator TRSetRow.[](index as string) as zstring ptr
	return hashLookup( @columns, index )
end operator
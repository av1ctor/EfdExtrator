#include once "DB.bi" 
#include once "list.bi" 

''''''''
function TDb.open(fileName as const zstring ptr) as boolean
	
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
function TDb.exec(query as const zstring ptr) as TRSet ptr

	var rs = new TRSet
	
	dim as zstring ptr errMsg_ = null
	if sqlite3_exec( instance, query, @callback, rs, @errMsg_ ) <> SQLITE_OK then 
		delete rs
		errMsg = *errMsg_
		return null
	end if 
	
	return rs

end function

''''''''	
function TDb.execScalar(query as const zstring ptr) as zstring ptr

	dim as TRSet rs
	
	dim as zstring ptr errMsg_ = null
	if sqlite3_exec( instance, query, @callback, @rs, @errMsg_ ) <> SQLITE_OK then 
		errMsg = *errMsg_
		return null
	end if 
	
	if rs.hasNext then
		var val = (*rs.row)[0]
		if val = null then
			return null
		end if
		
		var val2 = cast(zstring ptr, allocate(len(*val)+1))
		*val2 = *val
		function = val2
	else
		function = null
	end if
	
end function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
constructor TRSet()
	rows.init(10, len(TRSetRow))
	currRow = null
end constructor	
	
''''''''
destructor TRSet()
	var r = cast(TRSetRow ptr, rows.head)
	do while r <> null
		r->destructor
		r = rows.next_(r)
	loop
	
	rows.end_()
	currRow = null
end destructor

''''''''
function TRSet.hasNext() as boolean
	return currRow <> null
end function

''''''''
sub TRSet.next_() 
	if currRow <> null then
		currRow = rows.next_(currRow)
	end if
end sub

''''''''
property TRSet.row() as TRSetRow ptr
	return currRow
end property

''''''''
function TRSet.newRow() as TRSetRow ptr
	var p = rows.add()
	var r = new (p) TRSetRow()
	if currRow = null then
		currRow = r
	end if
	return r
end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
constructor TRSetRow()
	columns.init(16, true, true, true)
	redim colList(0 to 15)
	colCnt = 0
end constructor	
	
''''''''
destructor TRSetRow()
	colCnt = 0
	columns.end_()
end destructor

''''''''
sub TRSetRow.newColumn(name_ as const zstring ptr, value as const zstring ptr)
	if columns.lookup(name_) = null then
		dim as zstring ptr value2 = null
		if value <> null then
			value2 = cast(zstring ptr, allocate(len(*value)+1))	
			*value2 = *value
		end if
		
		columns.add( name_, value2 )
		
		colCnt += 1
		if colCnt-1 > ubound(colList) then
			redim preserve colList(0 to colCnt-1+8)
		end if

		colList(colCnt-1) = value2
	end if
end sub

''''''''
operator TRSetRow.[](index as const zstring ptr) as zstring ptr
	return columns.lookup( index )
end operator

''''''''
operator TRSetRow.[](index as integer) as zstring ptr
	if index <= colCnt-1 then
		return colList(index)
	else
		return null
	end if
end operator
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
function TDb.open() as boolean

	function = open(":memory:")

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
function TDb.getErrorMsg() as const zstring ptr
	function = strptr(errMsg)
end function

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
	else
		errMsg = ""
	end if 
	
	return rs

end function

''''''''	
function TDb.exec(stmt as TDbStmt ptr) as TRSet ptr

	var rs = new TRSet
	
	stmt->reset()
	
	do
		if stmt->step_() <> SQLITE_ROW then
			exit do
		end if
		
		var row = rs->newRow()
		
		var nCols = stmt->colCount()
		for i as integer = 0 to nCols - 1
			row->newColumn( stmt->colName( i ), stmt->colValue( i ) )
		next
	loop
	
	function = rs
	
end function

''''''''	
function TDb.execScalar(query as const zstring ptr) as zstring ptr

	dim as TRSet rs
	
	dim as zstring ptr errMsg_ = null
	if sqlite3_exec( instance, query, @callback, @rs, @errMsg_ ) <> SQLITE_OK then 
		errMsg = *errMsg_
		return null
	else
		errMsg = ""
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

''''''''	
sub TDb.execNonQuery(query as const zstring ptr) 

	var rs = new TRSet
	
	dim as zstring ptr errMsg_ = null
	if sqlite3_exec( instance, query, null, rs, @errMsg_ ) <> SQLITE_OK then 
		errMsg = *errMsg_
	else
		errMsg = ""
	end if 
	
	delete rs

end sub

''''''''	
sub TDb.execNonQuery(stmt as TDbStmt ptr) 

	do
		if stmt->step_() <> SQLITE_ROW then
			exit do
		end if
	loop

end sub
	
''''''''	
function TDb.prepare(query as const zstring ptr) as TDBStmt ptr

	var res = new TDbStmt(this.instance)
	if not res->prepare(query) then
		errMsg = *sqlite3_errmsg(instance)
		delete res
		return null
	else
		errMsg = ""
	end if
	
	function = res

end function

''''''''
function TDb.format cdecl(fmt as string, ...) as string

	dim as string args_v(0 to 9)
	dim as VarType args_t(0 to 9)

	var arg = va_first()
	var a = -1
	
	var res = ""
	
	var i = 0
	do while i < len(fmt)
		if fmt[i] = asc("{") then
			i += 1
			var j = cint(fmt[i] - asc("0"))
			i += 1
			
			if j > a then
				do until a = j
					a += 1
					var v = va_arg(arg, VarBox ptr)
					args_v(a) = *v
					args_t(a) = v->vtype
					arg = va_next(arg, VarBox ptr)
				loop
			end if
			
			if args_t(a) = VT_STR then
				res += "'" + args_v(j) + "'"
			else
				res += args_v(j)
			end if
		else
			res += chr(fmt[i])
		end if
	
		i += 1
	loop

	function = res
	
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
constructor TDbStmt(db as sqlite3 ptr)
	this.db = db
end constructor

''''''''
destructor TDbStmt()
	if stmt <> null then
		sqlite3_finalize(stmt)
	end if
end destructor

''''''''
function TDbStmt.prepare(query as const zstring ptr) as boolean
	function = sqlite3_prepare_v2(db, query, -1, @stmt, null) = SQLITE_OK
end function
	
''''''''	
sub TDbStmt.bind(index as integer, value as integer)
	sqlite3_bind_int(stmt, index, value)
end sub
	
''''''''	
sub TDbStmt.bind(index as integer, value as longint)
	sqlite3_bind_int64(stmt, index, value)
end sub
	
''''''''	
sub TDbStmt.bind(index as integer, value as double)
	sqlite3_bind_double(stmt, index, value)
end sub
	
''''''''	
sub TDbStmt.bind(index as integer, value as const zstring ptr)
	if value = null then
		sqlite3_bind_null(stmt, index)
	else
		sqlite3_bind_text(stmt, index, value, len(*value), null)
	end if
end sub

''''''''	
sub TDbStmt.bind(index as integer, value as const wstring ptr)
	if value = null then
		sqlite3_bind_null(stmt, index)
	else
		sqlite3_bind_text16(stmt, index, value, len(*value), null)
	end if
end sub

''''''''	
sub TDbStmt.bindNull(index as integer)
	sqlite3_bind_null(stmt, index)
end sub

''''''''	
function TDbStmt.step_() as long
	function = sqlite3_step(stmt)
end function

''''''''
sub TDbStmt.reset()
	sqlite3_reset(stmt)
end sub

''''''''
sub TDbStmt.clear_()
	sqlite3_clear_bindings(stmt)
end sub

''''''''
function TDbStmt.colCount() as integer
	function = sqlite3_column_count(stmt)
end function

''''''''
function TDbStmt.colName(index as integer) as const zstring ptr
	function = sqlite3_column_name(stmt, index)
end function

''''''''
function TDbStmt.colValue(index as integer) as const zstring ptr
	function = sqlite3_column_text(stmt, index)
end function
'' SQLite Helper Library for FreeBASIC
'' Copyright 2017 by Andre Victor (av1ctortv[@]gmail.com)
'' Licensed under GNU GPL-2.0-or-above

#include once "SQLite.bi" 

''''''''
function SQLite.open(fileName as const zstring ptr) as boolean
	
	if sqlite3_open( fileName, @instance ) then 
  		errMsg = *sqlite3_errmsg( instance )
		sqlite3_close( instance ) 
		return false
	end if 
	
	errMsg = ""
	return true
	
end function

''''''''
function SQLite.open() as boolean

	function = open(":memory:")

end function

''''''''
sub SQLite.close()
	if instance <> null then
		sqlite3_close( instance ) 
		instance = null
		errMsg = ""
	end if
end sub

''''''''
function SQLite.getErrorMsg() as const zstring ptr
	function = strptr(errMsg)
end function

''''''''
private function callback cdecl _
	( _
		byval dset as any ptr, _
		byval argc as long, _
		byval argv as zstring ptr ptr, _
		byval colName as zstring ptr ptr _
	) as long
	
	var ds = cast(SQLiteDataSet ptr, dset)
	
	var row = ds->newRow(argc)
  
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
function SQLite.exec(query as const zstring ptr) as SQLiteDataSet ptr

	var ds = new SQLiteDataSet
	
	dim as zstring ptr errMsg_ = null
	if sqlite3_exec( instance, query, @callback, ds, @errMsg_ ) <> SQLITE_OK then 
		delete ds
		errMsg = *errMsg_
		sqlite3_free(errMsg_)
		return null
	else
		errMsg = ""
	end if 
	
	return ds

end function

''''''''	
function SQLite.exec(stmt as SQLiteStmt ptr) as SQLiteDataSet ptr

	var ds = new SQLiteDataSet
	
	stmt->reset()
	
	do
		if stmt->step_() <> SQLITE_ROW then
			exit do
		end if
		
		var nCols = stmt->colCount()
		var row = ds->newRow(nCols)
		
		for i as integer = 0 to nCols - 1
			row->newColumn( stmt->colName( i ), stmt->colValue( i ) )
		next
	loop
	
	function = ds
	
end function

''''''''	
function SQLite.execScalar(query as const zstring ptr) as zstring ptr

	dim as SQLiteDataSet ds
	
	dim as zstring ptr errMsg_ = null
	if sqlite3_exec( instance, query, @callback, @ds, @errMsg_ ) <> SQLITE_OK then 
		errMsg = *errMsg_
		sqlite3_free(errMsg_)
		return null
	else
		errMsg = ""
	end if 
	
	if ds.hasNext then
		var val = (*ds.row)[0]
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
function SQLite.execNonQuery(query as const zstring ptr) as boolean

	var ds = new SQLiteDataSet
	
	dim as zstring ptr errMsg_ = null
	if sqlite3_exec( instance, query, null, ds, @errMsg_ ) <> SQLITE_OK then 
		errMsg = *errMsg_
		sqlite3_free(errMsg_)
		function = false
	else
		errMsg = ""
		function = true
	end if 
	
	delete ds

end function

''''''''	
function SQLite.execNonQuery(stmt as SQLiteStmt ptr) as boolean

	do
		if stmt->step_() <> SQLITE_ROW then
			exit do
		end if
	loop
	
	function = true

end function
	
''''''''	
function SQLite.prepare(query as const zstring ptr) as SQLiteStmt ptr

	var res = new SQLiteStmt(this.instance)
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
function SQLite.lastId() as long
	function = sqlite3_last_insert_rowid(instance)
end function

''''''''
/'
function SQLite.format cdecl(fmt as string, ...) as string

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
'/

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
constructor SQLiteDataSet()
	rows = new TList(10, len(SQLiteDataSetRow))
	currRow = null
end constructor	
	
''''''''
destructor SQLiteDataSet()
	var r = cast(SQLiteDataSetRow ptr, rows->head)
	do while r <> null
		r->destructor		'' NOTA: não user delete, porque foi criado com placement new
		r = rows->next_(r)
	loop
	
	delete rows
	currRow = null
end destructor

''''''''
function SQLiteDataSet.hasNext() as boolean
	return currRow <> null
end function

''''''''
sub SQLiteDataSet.next_() 
	if currRow <> null then
		currRow = rows->next_(currRow)
	end if
end sub

''''''''
property SQLiteDataSet.row() as SQLiteDataSetRow ptr
	return currRow
end property

''''''''
function SQLiteDataSet.newRow(cols as integer) as SQLiteDataSetRow ptr
	var p = rows->add()
	var r = new (p) SQLiteDataSetRow(cols)
	if currRow = null then
		currRow = r
	end if
	return r
end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
constructor SQLiteDataSetRow(numCols as integer)
	if numCols = 0 then
		numCols = 16
	end if
	dict = new TDict(numCols, true, true, true)
	redim cols(0 to numCols-1)
	cnt = 0
end constructor	
	
''''''''
destructor SQLiteDataSetRow()
	cnt = 0
	delete dict
end destructor

''''''''
sub SQLiteDataSetRow.newColumn(name_ as const zstring ptr, value as const zstring ptr)
	if dict->lookup(name_) = null then
		dim as zstring ptr value2 = null
		if value <> null then
			value2 = cast(zstring ptr, allocate(len(*value)+1))	
			*value2 = *value
		end if
		
		var node = dict->add( name_, value2 )
		
		cnt += 1
		if cnt-1 > ubound(cols) then
			redim preserve cols(0 to cnt-1+8)
		end if

		cols(cnt-1).name = cast(zstring ptr, node->key)
		cols(cnt-1).value = value2
	end if
end sub

''''''''
operator SQLiteDataSetRow.[](index as const zstring ptr) as zstring ptr
	return dict->lookup( index )
end operator

''''''''
operator SQLiteDataSetRow.[](index as integer) as zstring ptr
	if index >= 0 and index <= cnt-1 then
		return cols(index).value
	else
		return null
	end if
end operator

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
constructor SQLiteStmt(db as sqlite3 ptr)
	this.db = db
end constructor

''''''''
destructor SQLiteStmt()
	if stmt <> null then
		sqlite3_finalize(stmt)
	end if
end destructor

''''''''
function SQLiteStmt.prepare(query as const zstring ptr) as boolean
	function = sqlite3_prepare_v2(db, query, -1, @stmt, null) = SQLITE_OK
end function
	
''''''''	
sub SQLiteStmt.bind(index as integer, value as integer)
	sqlite3_bind_int(stmt, index, value)
end sub
	
''''''''	
sub SQLiteStmt.bind(index as integer, value as longint)
	sqlite3_bind_int64(stmt, index, value)
end sub
	
''''''''	
sub SQLiteStmt.bind(index as integer, value as double)
	sqlite3_bind_double(stmt, index, value)
end sub
	
''''''''	
sub SQLiteStmt.bind(index as integer, value as const zstring ptr)
	
	'' NOTE: the value string can't be freed or modified until exec() is called!
	
	if value = null then
		sqlite3_bind_null(stmt, index)
	else
		sqlite3_bind_text(stmt, index, value, len(*value), null)
	end if
end sub

''''''''	
sub SQLiteStmt.bind(index as integer, value as const wstring ptr)

	'' NOTE: the value string can't be freed or modified until exec() is called!
	
	if value = null then
		sqlite3_bind_null(stmt, index)
	else
		sqlite3_bind_text16(stmt, index, value, len(*value), null)
	end if
end sub

''''''''	
sub SQLiteStmt.bindNull(index as integer)
	sqlite3_bind_null(stmt, index)
end sub

''''''''	
function SQLiteStmt.step_() as long
	function = sqlite3_step(stmt)
end function

''''''''
sub SQLiteStmt.reset()
	sqlite3_reset(stmt)
end sub

''''''''
sub SQLiteStmt.clear_()
	sqlite3_clear_bindings(stmt)
end sub

''''''''
function SQLiteStmt.colCount() as integer
	function = sqlite3_column_count(stmt)
end function

''''''''
function SQLiteStmt.colName(index as integer) as const zstring ptr
	function = sqlite3_column_name(stmt, index)
end function

''''''''
function SQLiteStmt.colValue(index as integer) as const zstring ptr
	function = sqlite3_column_text(stmt, index)
end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
private function luacb_db_new cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	var db = new SQLite()
	lua_pushlightuserdata(L, db)
	
	function = 1
	
end function

''''''''
private function luacb_db_del cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var db = cast(SQLite ptr, lua_touserdata(L, 1))
		delete db
	end if
	
	function = 0
	
end function

''''''''
private function luacb_db_open cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args < 1 or args > 2 then
		lua_pushboolean(L, false)
	else
		var db = cast(SQLite ptr, lua_touserdata(L, 1))
		if args > 1 then
			var fname = lua_tostring(L, 2)
			lua_pushboolean(L, db->open(fname))
		else
			lua_pushboolean(L, db->open())
		end if
	end if
	
	function = 1
	
end function

''''''''
private function luacb_db_close cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var db = cast(SQLite ptr, lua_touserdata(L, 1))
		db->close()
	end if
	
	function = 0
	
end function

''''''''
private function luacb_db_execNonQuery cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var db = cast(SQLite ptr, lua_touserdata(L, 1))
		if lua_isstring(L, 2) then
			var query = lua_tostring(L, 2)
			if not db->execNonQuery(query) then
				print "SQL error: "; *db->getErrorMsg(); " at query: "; *query
			end if
		else
			var query = cast(SQLiteStmt ptr, lua_touserdata(L, 2))
			if not db->execNonQuery(query) then
				print "SQL error: "; *db->getErrorMsg()
			end if
		end if
	end if
	
	function = 0
	
end function

''''''''
private function luacb_db_exec cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var db = cast(SQLite ptr, lua_touserdata(L, 1))
		
		dim as SQLiteDataSet ptr ds = null
		if lua_isstring(L, 2) then
			var query = lua_tostring(L, 2)
			ds = db->exec(query)
			if ds = null then
				print "SQL error: "; *db->getErrorMsg(); " at query: "; *query
			end if
		else
			var query = cast(SQLiteStmt ptr, lua_touserdata(L, 2))
			ds = db->exec(query)
		end if
		
		if ds = null then
			lua_pushnil(L)
		else
			lua_pushlightuserdata(L, ds)
		end if
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_db_execScalarInt cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var db = cast(SQLite ptr, lua_touserdata(L, 1))
		
		dim as zstring ptr res = null
		if lua_isstring(L, 2) then
			var query = lua_tostring(L, 2)
			res = db->execScalar(query)
			if res = null then
				print "SQL error: "; *db->getErrorMsg(); " at query: "; *query
			end if
		end if
		
		if res = null then
			lua_pushnil(L)
		else
			lua_pushinteger(L, vallng(*res))
			deallocate res
		end if
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_db_execScalarDbl cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var db = cast(SQLite ptr, lua_touserdata(L, 1))
		
		dim as zstring ptr res = null
		if lua_isstring(L, 2) then
			var query = lua_tostring(L, 2)
			res = db->execScalar(query)
			if res = null then
				print "SQL error: "; *db->getErrorMsg(); " at query: "; *query
			end if
		end if
		
		if res = null then
			lua_pushnil(L)
		else
			lua_pushnumber(L, val(*res))
			deallocate res
		end if
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_db_prepare cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var db = cast(SQLite ptr, lua_touserdata(L, 1))
		var query = lua_tostring(L, 2)
		var stmt = db->prepare(query)
		if stmt <> null then
			lua_pushlightuserdata(L, stmt)
		else
			lua_pushnil(L)
			print "SQL error: "; *db->getErrorMsg(); " at query: "; *query
		end if
	else
		lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_ds_hasNext cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var ds = cast(SQLiteDataSet ptr, lua_touserdata(L, 1))
		
		lua_pushboolean(L, ds->hasNext())
	else
		lua_pushboolean(L, false)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_ds_next cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var ds = cast(SQLiteDataSet ptr, lua_touserdata(L, 1))
		
		ds->next_()
	end if
	
	function = 0
	
end function

''''''''
private function luacb_ds_del cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var ds = cast(SQLiteDataSet ptr, lua_touserdata(L, 1))
		
		delete ds
	end if
	
	function = 0
	
end function

''''''''
private function luacb_ds_row_getColValue cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var ds = cast(SQLiteDataSet ptr, lua_touserdata(L, 1))
		var colName = lua_tostring(L, 2)

		lua_pushstring(L, (*ds->row)[colName])
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_ds_row cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var ds = cast(SQLiteDataSet ptr, lua_touserdata(L, 1))

		var row = ds->currRow
		lua_createtable(L, row->cnt, 0)
		for i as integer = 0 to row->cnt-1
			lua_pushstring(L, row->cols(i).name)
			lua_pushstring(L, row->cols(i).value)
			lua_settable(L, -3)
		next 
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
static sub SQLite.exportAPI(L as lua_State ptr)
	
	lua_register(L, "db_new", @luacb_db_new)
	lua_register(L, "db_del", @luacb_db_del)
	lua_register(L, "db_open", @luacb_db_open)
	lua_register(L, "db_close", @luacb_db_close)
	lua_register(L, "db_execNonQuery", @luacb_db_execNonQuery)
	lua_register(L, "db_exec", @luacb_db_exec)
	lua_register(L, "db_execScalarInt", @luacb_db_execScalarInt)
	lua_register(L, "db_execScalarDbl", @luacb_db_execScalarDbl)
	lua_register(L, "db_prepare", @luacb_db_prepare)
	
	lua_register(L, "ds_hasNext", @luacb_ds_hasNext)
	lua_register(L, "ds_next", @luacb_ds_next)
	lua_register(L, "ds_row_getColValue", @luacb_ds_row_getColValue)
	lua_register(L, "ds_del", @luacb_ds_del)
	lua_register(L, "ds_row", @luacb_ds_row)
	
end sub
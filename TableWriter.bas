
#include once "TableWriter.bi"
#include "libiconv.bi"

''
constructor TableWriter()
	tables = new TableCollection
	fnum = 0
	xlsxWorkbook = null
	
	colType2Str(CT_STRING) 		= "String"
	colType2Str(CT_STRING_UTF8)	= "String"
	colType2Str(CT_NUMBER) 		= "Number"
	colType2Str(CT_INTNUMBER)   = "Number"
	colType2Str(CT_DATE) 		= "DateTime"
	colType2Str(CT_MONEY) 		= "Number"
	colType2Str(CT_PERCENT)		= "Number"

	colType2Sql(CT_STRING) 		= "text"
	colType2Sql(CT_STRING_UTF8)	= "text"
	colType2Sql(CT_NUMBER) 		= "real"
	colType2Sql(CT_INTNUMBER)   = "integer"
	colType2Sql(CT_DATE) 		= "date"
	colType2Sql(CT_MONEY) 		= "decimal(10,2)"
	colType2Sql(CT_PERCENT)		= "decimal(10,2)"

	colWidth(CT_STRING) 		= LXW_DEF_COL_WIDTH + 8
	colWidth(CT_STRING_UTF8)	= LXW_DEF_COL_WIDTH + 8
	colWidth(CT_NUMBER) 		= LXW_DEF_COL_WIDTH + 0
	colWidth(CT_INTNUMBER)   	= LXW_DEF_COL_WIDTH + 0
	colWidth(CT_DATE) 			= LXW_DEF_COL_WIDTH + 2
	colWidth(CT_MONEY) 			= LXW_DEF_COL_WIDTH + 6
	colWidth(CT_PERCENT)		= LXW_DEF_COL_WIDTH + 0

	cd = iconv_open("UTF-8", "ISO_8859-1")
end constructor

''
destructor TableWriter()
	if tables <> null then
		delete tables
		tables = null
	end if
	
	if fnum <> 0 then
		..close #fnum
		fnum = 0
	end if
	
	if xlsxWorkbook <> null then
		workbook_close(xlsxWorkbook)
	end if
	
	if db <> null then
		db->close()
		delete db
	end if
	
end destructor

''
function TableWriter.addTable(name_ as string) as TableTable ptr

	function = tables->addTable(name_)

end function

''
function TableWriter.create(fileName as string, ftype as FileType) as boolean
	
	this.ftype = ftype
	this.fileName = fileName
	
	select case ftype
	case FT_XML
		fnum = FreeFile

		var res = open(fileName + ".xml" for output as #fnum) = 0

		' header
		if res then 
			print #fnum, !"<?xml version=\"1.0\" encoding=\"iso-8859-1\"?>"
			print #fnum, !"<?mso-application progid=\"Excel.Sheet\"?>"
			print #fnum, !"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">"
		else
			fnum = 0
		end if
		function = res

	case FT_CSV
		fnum = 0
		function = true

	case FT_XLSX
		xlsxWorkbook = workbook_new(fileName + ".xlsx")
		
		for i as integer = 0 to CT__LEN__-1
			xlsxFormats(i) = workbook_add_format(xlsxWorkbook)
		next
		
		format_set_num_format(xlsxFormats(CT_MONEY), !"\"R$\" #,##0.00")
		format_set_num_format(xlsxFormats(CT_DATE), "dd/mm/yyyy")
		format_set_num_format(xlsxFormats(CT_INTNUMBER), "0")
		format_set_num_format(xlsxFormats(CT_NUMBER), "0.00")
		format_set_num_format(xlsxFormats(CT_PERCENT), "0.00%")
		function = true
		
	case FT_SQLITE
		kill fileName + ".db"
		db = new TDb
		db->open(fileName + ".db")
		db->execNonQuery("PRAGMA JOURNAL_MODE=OFF")
		db->execNonQuery("PRAGMA SYNCHRONOUS=0")
		db->execNonQuery("PRAGMA LOCKING_MODE=EXCLUSIVE")
		
	case else
		fnum = 0
		function = true
		
	end select

end function

private function escapeContent(src as string) as string
	for i as integer = 0 to len(src) - 1
		select case as const src[i]
		case asc("&")
			src[i] = asc("e")
		case asc("<")
			src[i] = asc("_")
		end select
	next
	function = src
end function

private function nameToSql(src as string) as string
	for i as integer = 0 to len(src) - 1
		select case as const src[i]
		case asc(" "), asc(".")
			src[i] = asc("_")
		end select
	next
	function = src
end function

private function latin2UTF8(src as zstring ptr, cd as iconv_t) as string
	var chars = len(*src)
	var dst = cast(zstring ptr, callocate(chars*2+1))
	var srcp = src
	var srcleft = chars
	var dstp = dst
	var dstleft = chars*2
	iconv(cd, @srcp, @srcleft, @dstp, @dstleft)
	*cast(byte ptr, dstp) = 0
	function = *dst
	deallocate dst
end function

''
function TableWriter.flush(onProgress as OnProgressCB, onError as OnErrorCB) as boolean

	var p = 1

	select case ftype
	case FT_XML
		print #fnum, !"<Styles>"
	end select

	var totalRows = 0
   
	'' para cada tabela..
	var table = tables->tableListHead
	do while table <> null
		
		if table->nRows > 1 then
			totalRows += table->nRows
		
			select case ftype
			case FT_XML
				'' para cada coluna..
				if table->colListHead <> null then
					var ct = table->colListHead
					var i = 1
					do while ct <> null
						print #fnum, !"<Style ss:ID=\"colStyle_" & p & "_" & i & !"\">"
						select case as const ct->type_
						case CT_DATE
							print #fnum, !"<NumberFormat ss:Format=\"Short Date\"/>"
						case CT_INTNUMBER
							print #fnum, !"<NumberFormat ss:Format=\"0\"/>"
						case CT_NUMBER
							print #fnum, !"<NumberFormat ss:Format=\"0.00\"/>"
						case CT_PERCENT
							print #fnum, !"<NumberFormat ss:Format=\"0.00%\"/>"
						case CT_MONEY
							print #fnum, !"<NumberFormat ss:Format=\"_-&quot;R$&quot;\\ * #,##0.00_-;\\-&quot;R$&quot;\\ * #,##0.00_-;_-&quot;R$&quot;\\ * &quot;-&quot;??_-;_-@_-\"/>"
						end select
						print #fnum, !"</Style>"
						ct = ct->next_
						i += 1
					loop
				end if
			end select
		end if
      
		table = table->next_
		p += 1
	loop

	select case ftype
	case FT_XML
		print #fnum, !"</Styles>"
	end select
      
	' para cada tabela..
	p = 1
	var curRow = 0
	table = tables->tableListHead
	do while table <> null
		
		if table->nRows > 1 then
			dim as lxw_worksheet ptr xlsXWorksheet
			dim as TDbStmt ptr stmt = null
			dim as integer totCols = 0
	
			select case ftype
			case FT_XML
				print #fnum, !"<Worksheet ss:Name=\"" + table->name + !"\">"
				print #fnum, !"<Table>"

				if table->colListHead <> null then
					var ct = table->colListHead
					var i = 1
					 do while ct <> null
						print #fnum, !"<Column ss:Index=\"" & i & !"\" ss:StyleID=\"colStyle_" & p & "_" & i & !"\" ss:AutoFitWidth=\"1\" />"
						ct = ct->next_
						i += 1
					 loop
				end if
				
			case FT_CSV
				fnum = FreeFile

				var csvName = fileName + "_" + table->name + ".csv"
				var res = open(csvName for output as #fnum) = 0
				if not res then
					onError(wstr("Erro: não foi possível criar arquivo " + csvName))
					return false
				end if

			case FT_XLSX
				xlsXWorksheet = workbook_add_worksheet(xlsxWorkbook, table->name)
				
				'' para cada coluna..
				if table->colListHead <> null then
					var ct = table->colListHead
					var colNum = 0
					do while ct <> null
						var wdt = iif(ct->width_ = 0, colWidth(ct->type_), ct->width_)
						worksheet_set_column(xlsXWorksheet, colNum, colNum, wdt, xlsxFormats(ct->type_))
						colNum += 1
						ct = ct->next_
					loop
				end if
				
			case FT_SQLITE
				var tblName = "'" & table->name & "'"
				
				var createTable = "create table " & tblName & "("
				var insertInto = "insert into " & tblName & "("
				
				var row = table->rowListHead
				var cell = row->cellListHead
				var ct = table->colListHead
				var colNum = 0
				do while cell <> null
					do while cell->num > colNum
						colNum += 1
						if ct <> null then
							ct = ct->next_
						end if
					loop
						
					var tp = iif(ct <> null, ct->type_, CT_STRING)
					var colName = nameToSql(cell->content)
					
					createTable &= "'" & colName & "' " & colType2Sql(tp) & " null,"
					insertInto &= "'" & colName & "',"
					
					colNum += 1
					if ct <> null then
						ct = ct->next_
					end if
					cell = cell->next_
				loop

				totCols = colNum
				
				if totCols > 0 then
					createTable = left(createTable, len(createTable)-1) & ")"
					if db->execNonQuery(createTable) = false then
						onError("Ao criar tabela: " & createTable)
						return false
					end if
				
					insertInto = left(insertInto, len(insertInto)-1) & ") values ("
					for i as integer = 1 to totCols-1
						insertInto &= "?,"
					next
					insertInto &= "?)"
					stmt = db->prepare(insertInto)
					if stmt = null then
						onError("Ao criar statement: " & insertInto)
						return false
					end if
				end if
				
			end select

			'' para cada linha..
			if table->rowListHead <> null then
				var rowNum = 0
				var row = table->rowListHead
				do while row <> null
					curRow += 1
					if onProgress <> null then
						if not onProgress(null, curRow / totalRows) then
							exit do
						end if
					end if
					
					if row->asIs then
						select case ftype
						case FT_XML
							print #fnum, !"<Row>"
							var cell = row->cellListHead
							var colNum = 0
							do while cell <> null
								do while cell->num > colNum
									print #fnum, "<Cell><Data></Data></Cell>"
									colNum += 1
								loop

								print #fnum, !"<Cell><Data ss:Type=\"String\">" + cell->content + "</Data></Cell>"
								cell = cell->next_
								colNum += 1
							loop
							print #fnum, !"</Row>"
						
						case FT_CSV
							var cell = row->cellListHead
							var colNum = 0
							do while cell <> null
								do while cell->num > colNum
									print #fnum, ";";
									colNum += 1
								loop
								print #fnum, cell->content + ";";
								colNum += 1
								cell = cell->next_
							loop
							print #fnum, chr(13, 10);
						
						case FT_XLSX
							var cell = row->cellListHead
							var colNum = 0
							do while cell <> null
								do while cell->num > colNum
									worksheet_write_string(xlsXWorksheet, rowNum, colNum, "", NULL)
									colNum += 1
								loop
								worksheet_write_string(xlsXWorksheet, rowNum, colNum, cell->content, NULL)
								colNum += 1
								cell = cell->next_
							loop
						end select
					
					else
						select case ftype
						case FT_XML
							print #fnum, !"<Row>"
							'' para cada celula da linha..
							var cell = row->cellListHead
							var ct = table->colListHead
							var colNum = 0
							do while cell <> null
								do while cell->num > colNum
									print #fnum, "<Cell><Data></Data></Cell>"
									colNum += 1
									if ct <> null then
										ct = ct->next_
									end if
								loop
								
								var content = cell->content
								select case ct->type_
								case CT_STRING, CT_STRING_UTF8
									content = escapeContent(content)
								end select
								print #fnum, !"<Cell><Data ss:Type=\"" + colType2Str(iif(ct <> null, ct->type_, CT_STRING)) + !"\">" + content + "</Data></Cell>"
								cell = cell->next_
								colNum += 1
								if ct <> null then
									ct = ct->next_
								end if
							loop
						
							print #fnum, !"</Row>"

						case FT_CSV
							'' para cada celula da linha..
							var cell = row->cellListHead
							var ct = table->colListHead
							var colNum = 0
							do while cell <> null
								do while cell->num > colNum
									print #fnum, ";";
									colNum += 1
									if ct <> null then
										ct = ct->next_
									end if
								loop
								print #fnum, cell->content + ";";
								cell = cell->next_
								colNum += 1
								if ct <> null then
									ct = ct->next_
								end if
							loop
							print #fnum, chr(13, 10);

						case FT_XLSX
							'' para cada celula da linha..
							var cell = row->cellListHead
							var ct = table->colListHead
							var colNum = 0
							do while cell <> null
								do while cell->num > colNum
									worksheet_write_string(xlsXWorksheet, rowNum, colNum, "", NULL)
									colNum += 1
									if ct <> null then
										ct = ct->next_
									end if
								loop
							
								if cell->width_ > 1 then
									worksheet_merge_range(xlsXWorksheet, rowNum, colNum, rowNum, colNum+cell->width_-1, cell->content, NULL)
									colNum += cell->width_
								else
									select case as const iif(ct <> null, ct->type_, CT_STRING)
									case CT_STRING
										worksheet_write_string(xlsXWorksheet, rowNum, colNum, latin2UTF8(cell->content, cd), NULL)
									case CT_STRING_UTF8
										worksheet_write_string(xlsXWorksheet, rowNum, colNum, cell->content, NULL)
									case CT_NUMBER, CT_MONEY, CT_PERCENT
										worksheet_write_number(xlsXWorksheet, rowNum, colNum, cdbl(cell->content), NULL)
									case CT_INTNUMBER
										worksheet_write_number(xlsXWorksheet, rowNum, colNum, clngint(cell->content), NULL)
									case CT_DATE
										var value = cell->content
										dim as lxw_datetime date = (valint(left(value, 4)), valint(mid(value, 6, 2)), valint(mid(value, 9, 2)))
										worksheet_write_datetime(xlsXWorksheet, rowNum, colNum, @date, NULL)
									end select
									
									colNum += 1
								end if
								
								cell = cell->next_
								if ct <> null then
									ct = ct->next_
								end if
							loop
						
						case FT_SQLITE
							if rowNum > 0 then
								'' para cada celula da linha..
								var cell = row->cellListHead
								var ct = table->colListHead
								var colNum = 0
								var sqlCol = 1
								dim bindContents(1 to totCols) as string
								stmt->reset()
								do while cell <> null
									do while cell->num > colNum
										stmt->bind(sqlCol, null)
										colNum += 1
										sqlCol += 1
										if ct <> null then
											ct = ct->next_
										end if
									loop
								
									select case as const iif(ct <> null, ct->type_, CT_STRING)
									case CT_STRING
										bindContents(sqlCol) = latin2UTF8(cell->content, cd)
										stmt->bind(sqlCol, bindContents(sqlCol))
									case CT_STRING_UTF8
										stmt->bind(sqlCol, cell->content)
									case CT_NUMBER, CT_MONEY, CT_PERCENT
										stmt->bind(sqlCol, cdbl(cell->content))
									case CT_INTNUMBER
										stmt->bind(sqlCol, clngint(cell->content))
									case CT_DATE
										var value = cell->content
										bindContents(sqlCol) = left(value, 4) & "-" & mid(value, 6, 2) & "-" & mid(value, 9, 2)
										stmt->bind(sqlCol, bindContents(sqlCol))
									end select
										
									colNum += 1
									sqlCol += 1
									
									cell = cell->next_
									if ct <> null then
										ct = ct->next_
									end if
								loop
								
								db->execNonQuery(stmt)
							end if
						end select
					end if
					
					rowNum += 1
					row = row->next_
				loop
			end if
			
			select case ftype
			case FT_XML
				print #fnum, !"</Table>"
				print #fnum, !"</Worksheet>"
			case FT_CSV
				..close #fnum
			case FT_SQLITE
				if stmt <> null then
					delete stmt
					stmt = null
				end if
			end select
		end if
		
		table = table->next_
		p += 1
	loop
	
	function = true
	
end function

''
sub TableWriter.close
	if fnum <> 0 then
		'' footer
		print #fnum, !"</Workbook>"

		..close #fnum 
		fnum = 0
	end if
	
	if xlsxWorkbook then
		workbook_close(xlsxWorkbook)
		xlsxWorkbook = null
	end if
end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
constructor TableColumn(type_ as ColumnType, width_ as integer, size as integer)
	this.type_ = type_
	this.width_ = width_
	this.size = size
end constructor

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
constructor TableTable(name as string)
	this.name = name
	curRow = 0
	nRows = 0
	redim rows(0 to 9999)
end constructor

''
destructor TableTable()

	do while colListHead <> null
		var next_ = colListHead->next_
		delete colListHead
		colListHead = next_
	loop

	do while rowListHead <> null
		var next_ = rowListHead->next_
		delete rowListHead
		rowListHead = next_
	loop
	
end destructor

''
function TableTable.addColumn(type_ as ColumnType, width_ as integer, size as integer) as TableColumn ptr

	var ct = new TableColumn(type_, width_, size)
	
	if colListHead = null then
		colListHead = ct
		colListTail = ct
	else
		colListTail->next_ = ct
		colListTail = ct
	end if
	
	function = ct

end function

''
function TableTable.addRow(asIs as boolean, num as integer) as TableRow ptr

	if num >= 0 then
		curRow = num
	end if
	
	if curRow > ubound(rows) then
		redim preserve rows(0 to curRow + 10000)
	end if
	
	var row = rows(curRow)
	if row = null then
		row = new TableRow(curRow, asIs)
		rows(curRow) = row
		
		if rowListHead = null then
			rowListHead = row
			rowListTail = row
		else
			rowListTail->next_ = row
			rowListTail = row
		end if
		
		nRows += 1
	end if
	
	curRow += 1
	
	function = row
	
end function

''
sub TableTable.setRow(num as integer)
	curRow = num
end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
destructor TableCollection
	do while tableListHead <> null
		var next_ = tableListHead->next_
		delete tableListHead
		tableListHead = next_
	loop
end destructor

''
function TableCollection.addTable(name_ as string) as TableTable ptr

	var table = new TableTable(name_)
	
	if tableListHead = null then
		tableListHead = table
		tableListTail = table
	else
		tableListTail->next_ = table
		tableListTail = table
	end if
	
	function = table

end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
constructor TableRow(num as integer, asIs as boolean)
	this.num = num
	this.asIs = asIs
end constructor

''
function TableRow.addCell(content as const zstring ptr, width_ as integer, num as integer) as TableCell ptr

	var cell = new TableCell(content, num)
	cell->width_ = width_
	
	if cellListHead = null then
		cellListHead = cell
		cellListTail = cell
	else
		cellListTail->next_ = cell
		cellListTail = cell
	end if
	
	function = cell

end function

''
function TableRow.addCell(content as integer, num as integer) as TableCell ptr

	function = AddCell(str(content), num)

end function

''
function TableRow.addCell(content as longint, num as integer) as TableCell ptr

	function = AddCell(str(content), num)

end function

''
function TableRow.addCell(content as double, num as integer) as TableCell ptr

	function = AddCell(str(content), num)

end function

''
destructor TableRow()

	do while cellListHead <> null
		var next_ = cellListHead->next_
		delete cellListHead
		cellListHead = next_
	loop

end destructor


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
constructor TableCell(content as const zstring ptr, num as integer)
	this.num = num
	this.content = *content
end constructor


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
private function luacb_ew_new cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	var ew = new TableWriter()
	lua_pushlightuserdata(L, ew)
	
	function = 1
	
end function

''''''''
private function luacb_ew_del cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var ew = cast(TableWriter ptr, lua_touserdata(L, 1))
		delete ew
	end if
	
	function = 0
	
end function

''''''''
private function luacb_ew_create cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 3 then
		var ew = cast(TableWriter ptr, lua_touserdata(L, 1))
		var fName = cast(zstring ptr, lua_tostring(L, 2))
		var isCSV = lua_tointeger(L, 3)
		
		lua_pushboolean(L, ew->create(*fName, isCSV))
	else
		lua_pushboolean(L, false)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_ew_close cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var ew = cast(TableWriter ptr, lua_touserdata(L, 1))
		
		ew->close()
	end if
	
	function = 1
	
end function

''''''''
private function luacb_ew_addTable cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var ew = cast(TableWriter ptr, lua_touserdata(L, 1))
		var wsName = cast(zstring ptr, lua_tostring(L, 2))
		
		lua_pushlightuserdata(L, ew->addTable(*wsName))
	else
		lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_ws_addRow cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 3 then
		var ws = cast(TableTable ptr, lua_touserdata(L, 1))
		var num = lua_tointeger(L, 2)
		var asIs = lua_tointeger(L, 3)
		
		lua_pushlightuserdata(L, ws->addRow(num, asIs))
	else
		lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_ws_addColumn cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var ws = cast(TableTable ptr, lua_touserdata(L, 1))
		var type_ = lua_tointeger(L, 2)
		
		ws->addColumn(type_)
	end if
	
	function = 0
	
end function

''''''''
private function luacb_er_addCell cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var er = cast(TableRow ptr, lua_touserdata(L, 1))
		
		dim as TableCell ptr ec = null
		if lua_isstring(L, 2) then
			ec = er->addCell(lua_tostring(L, 2))
		else
			ec = er->addCell(lua_tonumber(L, 2))
		end if
		
		lua_pushlightuserdata(L, ec)
	else
		lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
#define lua_defGlobal(L, varName, value) lua_pushnumber(L, cint(value)): lua_setglobal(L, varName)

''''''''
static sub TableWriter.exportAPI(L as lua_State ptr)
	
	lua_defGlobal(L, "CT_STRING", CT_STRING)
	lua_defGlobal(L, "CT_STRING_UTF8", CT_STRING_UTF8)
	lua_defGlobal(L, "CT_NUMBER", CT_NUMBER)
	lua_defGlobal(L, "CT_INTNUMBER", CT_INTNUMBER)
	lua_defGlobal(L, "CT_DATE", CT_DATE)
	lua_defGlobal(L, "CT_MONEY", CT_MONEY)
	lua_defGlobal(L, "CT_PERCENT", CT_PERCENT)
	
	lua_register(L, "ew_new", @luacb_ew_new)
	lua_register(L, "ew_del", @luacb_ew_del)
	lua_register(L, "ew_create", @luacb_ew_create)
	lua_register(L, "ew_close", @luacb_ew_close)
	lua_register(L, "ew_addTable", @luacb_ew_addTable)
	
	lua_register(L, "ws_addRow", @luacb_ws_addRow)
	lua_register(L, "ws_addColumn", @luacb_ws_addColumn)
	
	lua_register(L, "er_addCell", @luacb_er_addCell)
	
end sub
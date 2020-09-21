
#include once "ExcelWriter.bi"
#include "libiconv.bi"

''
constructor ExcelWriter()
	workbook = new ExcelWorkbook
	fnum = 0
	xlsxWorkbook = null
	
	CellType2String(CT_STRING) 		= "String"
	CellType2String(CT_NUMBER) 		= "Number"
	CellType2String(CT_INTNUMBER)   = "Number"
	CellType2String(CT_DATE) 		= "DateTime"
	CellType2String(CT_MONEY) 		= "Number"

	cd = iconv_open("UTF-8", "ISO_8859-1")
end constructor

''
destructor ExcelWriter()
	if workbook <> null then
		delete workbook
		workbook = null
	end if
	
	if fnum <> 0 then
		..close #fnum
		fnum = 0
	end if
	
	if xlsxWorkbook then
		workbook_close(xlsxWorkbook)
	end if
	
end destructor

''
function ExcelWriter.AddWorksheet(name_ as string) as ExcelWorksheet ptr

	function = workbook->AddWorksheet(name_)

end function

''
function ExcelWriter.Create(fileName as string, ftype as FileType) as boolean
	
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
		
		for i as integer = 0 to __CT_LEN__-1
			xlsxFormats(i) = workbook_add_format(xlsxWorkbook)
		next
		
		format_set_num_format(xlsxFormats(CT_MONEY), !"\"R$\" #,##0.00")
		format_set_num_format(xlsxFormats(CT_DATE), "dd/mm/yyyy")
		format_set_num_format(xlsxFormats(CT_INTNUMBER), "0")
		function = true
		
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
function ExcelWriter.Flush(showProgress as ProgressCB) as boolean

	var p = 1

	select case ftype
	case FT_XML
		print #fnum, !"<Styles>"
	end select

	var totalRows = 0
   
	'' para cada planilha..
	var sheet = workbook->worksheetListHead
	do while sheet <> null
   
		totalRows += sheet->nRows
	
		select case ftype
		case FT_XML
			'' para cada cell type..
			if sheet->cellTypeListHead <> null then
				var ct = sheet->cellTypeListHead
				var i = 1
				do while ct <> null
					print #fnum, !"<Style ss:ID=\"colStyle_" & p & "_" & i & !"\">"
					select case as const ct->type_
					case CT_DATE
						print #fnum, !"<NumberFormat ss:Format=\"Short Date\"/>"
					case CT_INTNUMBER
						print #fnum, !"<NumberFormat ss:Format=\"0\"/>"
					case CT_MONEY
						print #fnum, !"<NumberFormat ss:Format=\"_-&quot;R$&quot;\\ * #,##0.00_-;\\-&quot;R$&quot;\\ * #,##0.00_-;_-&quot;R$&quot;\\ * &quot;-&quot;??_-;_-@_-\"/>"
					end select
					print #fnum, !"</Style>"
					ct = ct->next_
					i += 1
				loop
			end if
		end select
      
		sheet = sheet->next_
		p += 1
	loop

	select case ftype
	case FT_XML
		print #fnum, !"</Styles>"
	end select
      
	' para cada planilha..
	p = 1
	var curRow = 0
	sheet = workbook->worksheetListHead
	do while sheet <> null
		dim as lxw_worksheet ptr xlsXWorksheet
	
		select case ftype
		case FT_XML
			print #fnum, !"<Worksheet ss:Name=\"" + sheet->name + !"\">"
			print #fnum, !"<Table>"

			if sheet->cellTypeListHead <> null then
				var ct = sheet->cellTypeListHead
				var i = 1
				 do while ct <> null
					print #fnum, !"<Column ss:Index=\"" & i & !"\" ss:StyleID=\"colStyle_" & p & "_" & i & !"\" ss:AutoFitWidth=\"1\" />"
					ct = ct->next_
					i += 1
				 loop
			end if
			
			'' para cada cell type..
			if sheet->cellTypeListHead <> null then
				print #fnum, !"<Row>"
				var ct = sheet->cellTypeListHead
				do while ct <> null
					print #fnum, !"<Cell><Data ss:Type=\"String\">" + ct->name + "</Data></Cell>"
					ct = ct->next_
				loop
				print #fnum, !"</Row>"
			end if
			
		case FT_CSV
			fnum = FreeFile

			var csvName = fileName + "_" + sheet->name + ".csv"
			var res = open(csvName for output as #fnum) = 0
			if not res then
				print wstr("Erro: não foi possível criar arquivo " + csvName)
				return false
			end if

			'' para cada cell type..
			if sheet->cellTypeListHead <> null then
				var ct = sheet->cellTypeListHead
				do while ct <> null
					print #fnum, ct->name + ";";
					ct = ct->next_
				loop
				print #fnum, chr(13, 10);
			end if
			
		case FT_XLSX
			xlsXWorksheet = workbook_add_worksheet(xlsxWorkbook, sheet->name)

			'' para cada cell type..
			if sheet->cellTypeListHead <> null then
				var ct = sheet->cellTypeListHead
				var col = 0
				do while ct <> null
					var colName = chr(asc("A") + col)
					worksheet_set_column(xlsXWorksheet, LXW_MAKE_COLS(colName + ":" + colName), LXW_DEF_COL_WIDTH, xlsxFormats(ct->type_))
					worksheet_write_string(xlsXWorksheet, 0, col, ct->name, NULL)
					col += 1
					ct = ct->next_
				loop
			end if
			
		end select

		'' para cada linha..
		if sheet->rowListHead <> null then
			var rowNum = 1
			var row = sheet->rowListHead
			do while row <> null
				
				curRow += 1
				if showProgress <> null then
					showProgress(null, curRow / totalRows)
				end if
				
				select case ftype
				case FT_XML
					print #fnum, !"<Row>"
					'' para cada cÃ©lula da linha..
					var cell = row->cellListHead
					var ct = sheet->cellTypeListHead
					do while cell <> null
						var content = cell->content
						if ct->type_ = CT_STRING then
							content = escapeContent(content)
						end if
						print #fnum, !"<Cell><Data ss:Type=\"" + CellType2String(iif(ct <> null, ct->type_, CT_STRING)) + !"\">" + content + "</Data></Cell>"
						cell = cell->next_
						if ct <> null then
							ct = ct->next_
						end if
					loop
				
					print #fnum, !"</Row>"

				case FT_CSV
					'' para cada cÃ©lula da linha..
					var cell = row->cellListHead
					var ct = sheet->cellTypeListHead
					do while cell <> null
						print #fnum, cell->content + ";";
						cell = cell->next_
						if ct <> null then
							ct = ct->next_
						end if
					loop
					print #fnum, chr(13, 10);

				case FT_XLSX
					'' para cada cÃ©lula da linha..
					var cell = row->cellListHead
					var ct = sheet->cellTypeListHead
					var colNum = 0
					do while cell <> null
						select case as const iif(ct <> null, ct->type_, CT_STRING)
						case CT_STRING
							worksheet_write_string(xlsXWorksheet, rowNum, colNum, latin2UTF8(cell->content, cd), NULL)
						case CT_NUMBER, _
							 CT_MONEY
							 worksheet_write_number(xlsXWorksheet, rowNum, colNum, cdbl(cell->content), NULL)
						case CT_INTNUMBER
							 worksheet_write_number(xlsXWorksheet, rowNum, colNum, clngint(cell->content), NULL)
						case CT_DATE
							var value = cell->content
							dim as lxw_datetime date = (valint(left(value, 4)), valint(mid(value, 6, 2)), valint(mid(value, 9, 2)))
							worksheet_write_datetime(xlsXWorksheet, rowNum, colNum, @date, NULL)
						end select
						
						colNum += 1
						cell = cell->next_
						if ct <> null then
							ct = ct->next_
						end if
					loop
				
				end select
				
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
		end select
		
		sheet = sheet->next_
		p += 1
	loop
	
	function = true
	
end function

''
sub ExcelWriter.close
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
constructor ExcelCellType(type_ as CellType, name_ as string)
	this.type_ = type_
	this.name = name_
end constructor

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
constructor ExcelWorksheet(name as string)
	this.name = name
	nRows = 0
end constructor

''
destructor ExcelWorksheet()

	do while cellTypeListHead <> null
		var next_ = cellTypeListHead->next_
		delete cellTypeListHead
		cellTypeListHead = next_
	loop

	do while rowListHead <> null
		var next_ = rowListHead->next_
		delete rowListHead
		rowListHead = next_
	loop
	
end destructor

''
function ExcelWorksheet.AddCellType(type_ as CellType, name_ as string) as ExcelCellType ptr

	var ct = new ExcelCellType(type_, name_)
	
	if cellTypeListHead = null then
		cellTypeListHead = ct
		cellTypeListTail = ct
	else
		cellTypeListTail->next_ = ct
		cellTypeListTail = ct
	end if
	
	function = ct

end function

''
function ExcelWorksheet.AddRow() as ExcelRow ptr

	var row = new ExcelRow()
	
	if rowListHead = null then
		rowListHead = row
		rowListTail = row
	else
		rowListTail->next_ = row
		rowListTail = row
	end if
	
	nRows += 1
	
	function = row

end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
destructor ExcelWorkbook
	do while worksheetListHead <> null
		var next_ = worksheetListHead->next_
		delete worksheetListHead
		worksheetListHead = next_
	loop
end destructor

''
function ExcelWorkbook.AddWorksheet(name_ as string) as ExcelWorksheet ptr

	var sheet = new ExcelWorksheet(name_)
	
	if worksheetListHead = null then
		worksheetListHead = sheet
		worksheetListTail = sheet
	else
		worksheetListTail->next_ = sheet
		worksheetListTail = sheet
	end if
	
	function = sheet

end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
function ExcelRow.AddCell(content as const zstring ptr) as ExcelCell ptr

	var cell = new ExcelCell(content)
	
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
function ExcelRow.AddCell(content as integer) as ExcelCell ptr

	function = AddCell(str(content))

end function

''
function ExcelRow.AddCell(content as longint) as ExcelCell ptr

	function = AddCell(str(content))

end function

''
function ExcelRow.AddCell(content as double) as ExcelCell ptr

	function = AddCell(str(content))

end function

''
destructor ExcelRow()

	do while cellListHead <> null
		var next_ = cellListHead->next_
		delete cellListHead
		cellListHead = next_
	loop

end destructor


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
constructor ExcelCell(content as const zstring ptr)
	this.content = *content
end constructor


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
private function luacb_ew_new cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	var ew = new ExcelWriter()
	lua_pushlightuserdata(L, ew)
	
	function = 1
	
end function

''''''''
private function luacb_ew_del cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var ew = cast(ExcelWriter ptr, lua_touserdata(L, 1))
		delete ew
	end if
	
	function = 0
	
end function

''''''''
private function luacb_ew_create cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 3 then
		var ew = cast(ExcelWriter ptr, lua_touserdata(L, 1))
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
		var ew = cast(ExcelWriter ptr, lua_touserdata(L, 1))
		
		ew->close()
	end if
	
	function = 1
	
end function

''''''''
private function luacb_ew_addWorksheet cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var ew = cast(ExcelWriter ptr, lua_touserdata(L, 1))
		var wsName = cast(zstring ptr, lua_tostring(L, 2))
		
		lua_pushlightuserdata(L, ew->addWorksheet(*wsName))
	else
		lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_ws_addRow cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var ws = cast(ExcelWorksheet ptr, lua_touserdata(L, 1))
		
		lua_pushlightuserdata(L, ws->addRow())
	else
		lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_ws_addCellType cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 3 then
		var ws = cast(ExcelWorksheet ptr, lua_touserdata(L, 1))
		var type_ = lua_tointeger(L, 2)
		var ctname = cast(zstring ptr, lua_tostring(L, 3))
		
		ws->addCellType(type_, *ctname)
	end if
	
	function = 0
	
end function

''''''''
private function luacb_er_addCell cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 2 then
		var er = cast(ExcelRow ptr, lua_touserdata(L, 1))
		
		dim as ExcelCell ptr ec = null
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
static sub ExcelWriter.exportAPI(L as lua_State ptr)
	
	lua_defGlobal(L, "CT_STRING", CT_STRING)
	lua_defGlobal(L, "CT_NUMBER", CT_NUMBER)
	lua_defGlobal(L, "CT_INTNUMBER", CT_INTNUMBER)
	lua_defGlobal(L, "CT_DATE", CT_DATE)
	lua_defGlobal(L, "CT_MONEY", CT_MONEY)
	
	lua_register(L, "ew_new", @luacb_ew_new)
	lua_register(L, "ew_del", @luacb_ew_del)
	lua_register(L, "ew_create", @luacb_ew_create)
	lua_register(L, "ew_close", @luacb_ew_close)
	lua_register(L, "ew_addWorksheet", @luacb_ew_addWorksheet)
	
	lua_register(L, "ws_addRow", @luacb_ws_addRow)
	lua_register(L, "ws_addCellType", @luacb_ws_addCellType)
	
	lua_register(L, "er_addCell", @luacb_er_addCell)
	
end sub
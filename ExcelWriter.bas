
#include once "ExcelWriter.bi"

''
constructor ExcelWriter()
	workbook = new ExcelWorkbook
	fnum = 0
	
	CellType2String(CT_STRING) 		= "String"
	CellType2String(CT_NUMBER) 		= "Number"
	CellType2String(CT_INTNUMBER)   = "Number"
	CellType2String(CT_DATE) 		= "DateTime"
	CellType2String(CT_MONEY) 		= "Number"
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
	
end destructor

''
function ExcelWriter.AddWorksheet(name_ as string) as ExcelWorksheet ptr

	function = workbook->AddWorksheet(name_)

end function

''
function ExcelWriter.Create(fileName as string) as boolean
	fnum = FreeFile

	var res = open(fileName for output as #fnum) = 0

	' header
	if res then 
		print #fnum, !"<?xml version=\"1.0\" encoding=\"iso-8859-1\"?>"
		print #fnum, !"<?mso-application progid=\"Excel.Sheet\"?>"
		print #fnum, !"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">"
	else
		fnum = 0
	end if

	function = res

end function

''
function ExcelWriter.Flush(showProgress as ProgressCB) as boolean

	var sheet = workbook->worksheetListHead
	var p = 1

	print #fnum, !"<Styles>"

	var totalRows = 0
   
	'' para cada planilha..
	do while sheet <> null
   
		totalRows += sheet->nRows
	
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
      
		sheet = sheet->next_
		p += 1
	loop

	print #fnum, !"</Styles>"
      
	' para cada planilha..
	sheet = workbook->worksheetListHead
	p = 1
	var curRow = 0
	do while sheet <> null
	
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

		'' para cada linha..
		if sheet->rowListHead <> null then
			var row = sheet->rowListHead
			do while row <> null
				
				curRow += 1
				if showProgress <> null then
					showProgress(null, curRow / totalRows)
				end if
				
				print #fnum, !"<Row>"
				'' para cada célula da linha..
				var cell = row->cellListHead
				var ct = sheet->cellTypeListHead
				do while cell <> null
					print #fnum, !"<Cell><Data ss:Type=\"" + CellType2String(iif(ct <> null, ct->type_, CT_STRING)) + !"\">" + cell->content + "</Data></Cell>"
					cell = cell->next_
					if ct <> null then
						ct = ct->next_
					end if
				loop
				
				row = row->next_
				print #fnum, !"</Row>"
			loop
		end if
		
		print #fnum, !"</Table>"
		print #fnum, !"</Worksheet>"
		
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
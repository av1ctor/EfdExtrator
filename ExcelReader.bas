#include once "ExcelReader.bi"
#include once "crt/time.bi"

'''''
private function utf8toLatin(cd as iconv_t, src as zstring ptr) as string
	var chars = len(*src)
	var dst = cast(zstring ptr, callocate(chars+1))
	var srcp = src
	var srcleft = chars
	var dstp = dst
	var dstleft = chars
	iconv(cd, @srcp, @srcleft, @dstp, @dstleft)
	*cast(byte ptr, dstp) = 0
	function = *dst
	deallocate dst
end function

'''''
constructor ExcelReader()
	cd = iconv_open("ISO_8859-1", "UTF-8")
end constructor

'''''
destructor ExcelReader()
	if xreader then
		xlsxioread_close(xreader)
		xreader = null
	end if
	
	iconv_close(cd)
end destructor

'''''
function ExcelReader.open(fileName as zstring ptr) as boolean
	xreader = xlsxioread_open(fileName)
	function = (xreader <> NULL)
end function

'''''
function ExcelReader.setSheet(sheetName as zstring ptr) as boolean
	
	sheet = xlsxioread_sheet_open(xreader, sheetName, XLSXIOREAD_SKIP_EMPTY_ROWS)
	function = (sheet <> NULL)
	
end function

'''''
function ExcelReader.nextRow() as boolean
	function = xlsxioread_sheet_next_row(sheet)
end function

'''''
function ExcelReader.read(toLatin as boolean) as string

	var value = xlsxioread_sheet_next_cell(sheet)
	if (value = null) then
		return ""
	end if
	
	if toLatin then
		function = utf8toLatin(cd, value)
	else
		function = *value
	end if
	
	deallocate value

end function

'''''
function ExcelReader.readDbl() as double

	var value = xlsxioread_sheet_next_cell(sheet)
	if (value = null) then
		return 0
	end if
	
	function = cdbl(*value)
	
	deallocate value

end function

'''''
function ExcelReader.readInt() as longint

	var value = xlsxioread_sheet_next_cell(sheet)
	if (value = null) then
		return 0
	end if
	
	function = clngint(*value)
	
	deallocate value

end function

'''''
function ExcelReader.readDate(fmt as zstring ptr) as string

	var value = xlsxioread_sheet_next_cell(sheet)
	if (value = null) then
		return ""
	end if
	
	var date = cast(time_t, (cdbl(*value) - 25569) * 86400) '' Unix timestamp
	
	deallocate value

	dim as zstring * 64 buff
	strftime(@buff, 64, fmt, localtime(@date))
	
	function = buff
	
end function

'''''
sub ExcelReader.skip() 

	var value = xlsxioread_sheet_next_cell(sheet)
	if (value <> null) then
		deallocate value
	end if
	
end sub
	

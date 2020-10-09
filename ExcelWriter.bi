
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#include once "libiconv.bi"
#include once "xlsxwriter.bi"
#define NULL 0

enum CellType
	CT_STRING
	CT_STRING_UTF8
	CT_NUMBER
	CT_INTNUMBER
	CT_PERCENT
	CT_DATE	
	CT_MONEY
	__CT_LEN__
end enum

type ExcelCellType
	type_				   	as CellType = CT_STRING
	next_				   	as ExcelCellType ptr = null
	
	declare constructor(type_ as CellType)
end type

type ExcelCell
	num						as integer
	content			   		as string
	width_					as integer
	next_				   	as ExcelCell ptr = null
	
	declare constructor(content as const zstring ptr, num as integer = -1)
end type

type ExcelRow
	asIs					as boolean = false
	num						as integer
	cellListHead	   		as ExcelCell ptr = null
	cellListTail	   		as ExcelCell ptr = null
	next_				   	as ExcelRow ptr = null
	
	declare constructor(num as integer, asIs as boolean = false)
	declare destructor
	declare function AddCell(content as const zstring ptr, width_ as integer = 1, num as integer = -1) as ExcelCell ptr
	declare function AddCell(content as integer, num as integer = -1) as ExcelCell ptr
	declare function AddCell(content as longint, num as integer = -1) as ExcelCell ptr
	declare function AddCell(content as double, num as integer = -1) as ExcelCell ptr
end type

type ExcelWorksheet
	name					as string
	cellTypeListHead		as ExcelCellType ptr = null
	cellTypeListTail		as ExcelCellType ptr = null
	rowListHead				as ExcelRow ptr = null
	rowListTail				as ExcelRow ptr = null
	rows(any)				as ExcelRow ptr
	curRow					as integer
	next_					as ExcelWorksheet ptr = null
	
	declare constructor(name as string)
	declare destructor
	declare function AddCellType(type_ as CellType) as ExcelCellType ptr
	declare function AddRow(asIs as boolean = false, num as integer = -1) as ExcelRow ptr
	declare sub setRow(num as integer)
end type

type ExcelWorkbook
	worksheetListHead		as ExcelWorksheet ptr = null
	worksheetListTail		as ExcelWorksheet ptr = null
	
	declare destructor
	declare function AddWorksheet(name as string) as ExcelWorksheet ptr
end type

type ProgressCB as sub(stage as const wstring ptr, preComplete as double)

enum FileType
	FT_XLSX
	FT_XML
	FT_CSV
	FT_NULL
end enum

type ExcelWriter
	declare constructor
	declare destructor
	declare function AddWorksheet(name as string) as ExcelWorksheet ptr
	declare function create(fileName as string, ftype as FileType = FT_XLSX) as boolean
	declare function flush(showProgress as ProgressCB) as boolean
	declare sub close
	declare static sub exportAPI(L as lua_State ptr)
	
private:
	ftype					as FileType
	fileName				as string
	fnum				   	as integer = 0
	xlsxWorkbook			as lxw_workbook ptr
	xlsxFormats(0 to __CT_LEN__-1) as lxw_format ptr
	cd						as iconv_t
	
	workbook 				as ExcelWorkbook ptr = null
	CellType2String(0 to __CT_LEN__-1) as string
	
end type
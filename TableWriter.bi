
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#include once "libiconv.bi"
#include once "xlsxwriter.bi"
#define NULL 0

enum ColumnType
	CT_STRING
	CT_STRING_UTF8
	CT_NUMBER
	CT_INTNUMBER
	CT_PERCENT
	CT_DATE	
	CT_MONEY
	__CT_LEN__
end enum

type TableColumn
	type_				   	as ColumnType = CT_STRING
	width_					as integer
	next_				   	as TableColumn ptr = null
	
	declare constructor(type_ as ColumnType, width_ as integer = 0)
end type

type TableCell
	num						as integer
	content			   		as string
	width_					as integer
	next_				   	as TableCell ptr = null
	
	declare constructor(content as const zstring ptr, num as integer = -1)
end type

type TableRow
	asIs					as boolean = false
	num						as integer
	cellListHead	   		as TableCell ptr = null
	cellListTail	   		as TableCell ptr = null
	next_				   	as TableRow ptr = null
	
	declare constructor(num as integer, asIs as boolean = false)
	declare destructor
	declare function addCell(content as const zstring ptr, width_ as integer = 1, num as integer = -1) as TableCell ptr
	declare function addCell(content as integer, num as integer = -1) as TableCell ptr
	declare function addCell(content as longint, num as integer = -1) as TableCell ptr
	declare function addCell(content as double, num as integer = -1) as TableCell ptr
end type

type TableTable
	name					as string
	colListHead				as TableColumn ptr = null
	colListTail				as TableColumn ptr = null
	rowListHead				as TableRow ptr = null
	rowListTail				as TableRow ptr = null
	rows(any)				as TableRow ptr
	curRow					as integer
	nRows					as integer
	next_					as TableTable ptr = null
	
	declare constructor(name as string)
	declare destructor
	declare function addColumn(type_ as ColumnType, width_ as integer = 0) as TableColumn ptr
	declare function addRow(asIs as boolean = false, num as integer = -1) as TableRow ptr
	declare sub setRow(num as integer)
end type

type TableCollection
	tableListHead		as TableTable ptr = null
	tableListTail		as TableTable ptr = null
	
	declare destructor
	declare function addTable(name as string) as TableTable ptr
end type

type OnProgressCB as function(stage as const zstring ptr, perComplete as double) as boolean

enum FileType
	FT_XLSX
	FT_XML
	FT_CSV
	FT_SQLITE
	FT_ACCESS
	FT_NULL
end enum

type TableWriter
	declare constructor
	declare destructor
	declare function addTable(name as string) as TableTable ptr
	declare function create(fileName as string, ftype as FileType = FT_XLSX) as boolean
	declare function flush(onProgress as OnProgressCB) as boolean
	declare sub close
	declare static sub exportAPI(L as lua_State ptr)
	
private:
	ftype as FileType
	fileName as string
	fnum as integer = 0
	xlsxWorkbook as lxw_workbook ptr
	xlsxFormats(0 to __CT_LEN__-1) as lxw_format ptr
	cd as iconv_t
	
	tables as TableCollection ptr = null
	colType2Str(0 to __CT_LEN__-1) as string
	colWidth(0 to __CT_LEN__-1) as integer
end type
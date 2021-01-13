
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#include once "libiconv.bi"
#include once "libs/xlsxwriter.bi"
#include once "libs/SQLite.bi"
#define NULL 0

const MAX_COLUMNS = 128

enum ColumnType
	CT_STRING
	CT_STRING_UTF8
	CT_NUMBER
	CT_INTNUMBER
	CT_PERCENT
	CT_DATE	
	CT_MONEY
	CT__LEN__
end enum

type TableColumn
	type_ as ColumnType = CT_STRING
	width_ as integer
	size as integer
	next_ as TableColumn ptr = null
	
	declare constructor(type_ as ColumnType, width_ as integer = 0, size as integer = 0)
end type

type TableCell
	num as integer
	content as string
	width_ as integer
	next_ as TableCell ptr
	
	declare constructor(content as const zstring ptr, num as integer = -1)
end type

type TableRow
	asIs as boolean
	num as integer
	cellListHead as TableCell ptr
	cellListTail as TableCell ptr
	next_ as TableRow ptr
	
	declare constructor(num as integer, asIs as boolean = false)
	declare destructor
	declare function addCell(content as const zstring ptr, width_ as integer = 1, num as integer = -1) as TableCell ptr
	declare function addCell(content as integer, num as integer = -1) as TableCell ptr
	declare function addCell(content as longint, num as integer = -1) as TableCell ptr
	declare function addCell(content as double, num as integer = -1) as TableCell ptr
end type

type TableTable
	name as string
	colListHead as TableColumn ptr
	colListTail as TableColumn ptr
	rowListHead as TableRow ptr
	rowListTail as TableRow ptr
	rows(any) as TableRow ptr
	curRow as integer
	nRows as integer
	next_ as TableTable ptr
	
	declare constructor(name as string)
	declare destructor
	declare function addColumn(type_ as ColumnType, width_ as integer = 0, size as integer = 0) as TableColumn ptr
	declare function addRow(asIs as boolean = false, num as integer = -1) as TableRow ptr
	declare sub setRow(num as integer)
end type

type TableCollection
	tableListHead as TableTable ptr
	tableListTail as TableTable ptr
	
	declare destructor
	declare function addTable(name as string) as TableTable ptr
end type

type OnProgressCB as function(stage as const zstring ptr, perComplete as double) as boolean
type OnErrorCB as sub(msg as const zstring ptr)

enum FileType
	FT_XLSX
	FT_XML
	FT_CSV
	FT_SQLITE
	FT_ACCESS
	FT_NULL
end enum

type TOdbc_ as TOdbc

type TableWriter
	declare constructor()
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as TableWriter ptr
	declare destructor
	declare function addTable(name as string) as TableTable ptr
	declare function create(fileName as string, ftype as FileType = FT_XLSX) as boolean
	declare function flush() as boolean
	declare sub close
	declare static sub exportAPI(L as lua_State ptr)
	
private:
	onProgress as OnProgressCB
	onError as OnErrorCB
	
	ftype as FileType
	fileName as string
	fnum as integer = 0
	xlsxWorkbook as lxw_workbook ptr
	xlsxFormats(0 to CT__LEN__-1) as lxw_format ptr
	db as SQLite ptr
	odbc as TOdbc_ ptr
	
	cdLatinToUtf8 as iconv_t
	cdUtf8ToLatin as iconv_t
	
	tables as TableCollection ptr = null
	colType2Str(0 to CT__LEN__-1) as string
	colType2Sql(0 to CT__LEN__-1) as string
	colWidth(0 to CT__LEN__-1) as integer
end type
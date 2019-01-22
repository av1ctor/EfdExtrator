
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#define NULL 0

enum CellType
	CT_STRING
	CT_NUMBER
	CT_INTNUMBER
	CT_DATE	
	CT_MONEY
	__CT_LEN__
end enum

type ExcelCellType
	type_				   	as CellType = CT_STRING
	name				   	as string
	next_				   	as ExcelCellType ptr = null
	
	declare constructor(type_ as CellType, name as string)
end type

type ExcelCell
	content			   		as string
	next_				   	as ExcelCell ptr = null
	
	declare constructor(content as const zstring ptr)
end type

type ExcelRow
	cellListHead	   		as ExcelCell ptr = null
	cellListTail	   		as ExcelCell ptr = null
	next_				   	as ExcelRow ptr = null
	
	declare destructor
	declare function AddCell(content as const zstring ptr) as ExcelCell ptr
	declare function AddCell(content as integer) as ExcelCell ptr
	declare function AddCell(content as longint) as ExcelCell ptr
	declare function AddCell(content as double) as ExcelCell ptr
end type

type ExcelWorksheet
	name					as string
	cellTypeListHead		as ExcelCellType ptr = null
	cellTypeListTail		as ExcelCellType ptr = null
	rowListHead				as ExcelRow ptr = null
	rowListTail				as ExcelRow ptr = null
	nRows					as integer
	next_					as ExcelWorksheet ptr = null
	
	declare constructor(name as string)
	declare destructor
	declare function AddCellType(type_ as CellType, name as string) as ExcelCellType ptr
	declare function AddRow() as ExcelRow ptr
end type

type ExcelWorkbook
	worksheetListHead		as ExcelWorksheet ptr = null
	worksheetListTail		as ExcelWorksheet ptr = null
	
	declare destructor
	declare function AddWorksheet(name as string) as ExcelWorksheet ptr
end type

type ProgressCB as sub(stage as const wstring ptr, preComplete as double)

type ExcelWriter
	declare constructor
	declare destructor
	declare function AddWorksheet(name as string) as ExcelWorksheet ptr
	declare function create(fileName as string, generateCSV as boolean) as boolean
	declare function flush(showProgress as ProgressCB) as boolean
	declare sub close
	declare static sub exportAPI(L as lua_State ptr)
	
private:
	isCSV					as boolean = false
	fileName				as string
	fnum				   	as integer = 0
	workbook 				as ExcelWorkbook ptr = null
	CellType2String(0 to __CT_LEN__-1) as string
end type
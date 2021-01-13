#pragma once

#include once "crt/stdlib.bi"
#include once "crt/stdint.bi"
#include once "crt/time.bi"

#ifdef XML_UNICODE
	#include once "crt/wchar.bi"
	#inclib "xlsxio_readw"
	#inclib "expatw-1"
	type XLSXIOCHAR as wstring
#else
	#inclib "xlsxio_read"
	#inclib "expat-1"
	type XLSXIOCHAR as zstring
#endif
#inclib "zip"

extern "C"

declare sub xlsxioread_get_version(byval pmajor as long ptr, byval pminor as long ptr, byval pmicro as long ptr)
declare function xlsxioread_get_version_string() as const XLSXIOCHAR ptr
type xlsxioreader as xlsxio_read_struct ptr
declare function xlsxioread_open(byval filename as const zstring ptr) as xlsxioreader
declare function xlsxioread_open_filehandle(byval filehandle as long) as xlsxioreader
declare function xlsxioread_open_memory(byval data as any ptr, byval datalen as ulongint, byval freedata as long) as xlsxioreader
declare sub xlsxioread_close(byval handle as xlsxioreader)
type xlsxioread_list_sheets_callback_fn as function(byval name as const XLSXIOCHAR ptr, byval callbackdata as any ptr) as long
declare sub xlsxioread_list_sheets(byval handle as xlsxioreader, byval callback as xlsxioread_list_sheets_callback_fn, byval callbackdata as any ptr)

const XLSXIOREAD_SKIP_NONE = 0
const XLSXIOREAD_SKIP_EMPTY_ROWS = &h01
const XLSXIOREAD_SKIP_EMPTY_CELLS = &h02
const XLSXIOREAD_SKIP_ALL_EMPTY = XLSXIOREAD_SKIP_EMPTY_ROWS or XLSXIOREAD_SKIP_EMPTY_CELLS
const XLSXIOREAD_SKIP_EXTRA_CELLS = &h04
type xlsxioread_process_cell_callback_fn as function(byval row as uinteger, byval col as uinteger, byval value as const XLSXIOCHAR ptr, byval callbackdata as any ptr) as long
type xlsxioread_process_row_callback_fn as function(byval row as uinteger, byval maxcol as uinteger, byval callbackdata as any ptr) as long
declare function xlsxioread_process(byval handle as xlsxioreader, byval sheetname as const XLSXIOCHAR ptr, byval flags as ulong, byval cell_callback as xlsxioread_process_cell_callback_fn, byval row_callback as xlsxioread_process_row_callback_fn, byval callbackdata as any ptr) as long
type xlsxioreadersheetlist as xlsxio_read_sheetlist_struct ptr

declare function xlsxioread_sheetlist_open(byval handle as xlsxioreader) as xlsxioreadersheetlist
declare sub xlsxioread_sheetlist_close(byval sheetlisthandle as xlsxioreadersheetlist)
declare function xlsxioread_sheetlist_next(byval sheetlisthandle as xlsxioreadersheetlist) as const XLSXIOCHAR ptr
type xlsxioreadersheet as xlsxio_read_sheet_struct ptr
declare function xlsxioread_sheet_open(byval handle as xlsxioreader, byval sheetname as const XLSXIOCHAR ptr, byval flags as ulong) as xlsxioreadersheet
declare sub xlsxioread_sheet_close(byval sheethandle as xlsxioreadersheet)
declare function xlsxioread_sheet_next_row(byval sheethandle as xlsxioreadersheet) as long
declare function xlsxioread_sheet_next_cell(byval sheethandle as xlsxioreadersheet) as XLSXIOCHAR ptr
declare function xlsxioread_sheet_next_cell_string(byval sheethandle as xlsxioreadersheet, byval pvalue as XLSXIOCHAR ptr ptr) as long
declare function xlsxioread_sheet_next_cell_int(byval sheethandle as xlsxioreadersheet, byval pvalue as longint ptr) as long
declare function xlsxioread_sheet_next_cell_float(byval sheethandle as xlsxioreadersheet, byval pvalue as double ptr) as long
declare function xlsxioread_sheet_next_cell_datetime(byval sheethandle as xlsxioreadersheet, byval pvalue as time_t ptr) as long

end extern

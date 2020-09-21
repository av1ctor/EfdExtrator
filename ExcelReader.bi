#include once "xlsxio_read.bi"
#include once "libiconv.bi"

type ExcelReader
public:
	declare constructor ()
	declare destructor ()
	declare function open(fileName as zstring ptr) as boolean
	declare function setSheet(sheetName as zstring ptr) as boolean
	declare function nextRow() as boolean
	declare function read(toLatin as boolean = false) as string
	declare function readDbl() as double
	declare function readInt() as longint
	declare function readDate(fmt as zstring ptr = @"%Y%m%dT%H:%M:%S.000") as string
	declare sub skip() 

private:
	cd			as iconv_t
	xreader		as xlsxioreader
	sheet		as xlsxioreadersheet
end type
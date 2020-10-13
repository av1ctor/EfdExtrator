#include once "libiconv.bi"

dim shared cdUtf8ToUtf16Le as iconv_t
dim shared cdLatinToUtf16Le as iconv_t

private sub init() constructor
	cdLatinToUtf16Le = iconv_open("UTF-16LE", "ISO_8859-1")
	cdUtf8ToUtf16Le = iconv_open("UTF-16LE", "UTF-8")
end sub

private sub shutdown() destructor
	iconv_close(cdUtf8ToUtf16Le)
	iconv_close(cdLatinToUtf16Le)
end sub

'''''
function latinToUtf16le(src as const zstring ptr) as wstring ptr
	var bytes = len(*src)
	var dst = allocate((bytes+1) * len(wstring))
	var srcp = src
	var srcleft = bytes
	var dstp = dst
	var dstleft = bytes*2
	iconv(cdLatinToUtf16Le, @srcp, @srcleft, @dstp, @dstleft)
	*cast(wstring ptr, dstp) = 0
	function = dst
end function

'''''
function utf8ToUtf16le(src as const zstring ptr) as wstring ptr
	var bytes = len(*src)
	var dst = allocate((bytes+1) * len(wstring))
	var srcp = src
	var srcleft = bytes
	var dstp = dst
	var dstleft = bytes*2
	iconv(cdUtf8ToUtf16Le, @srcp, @srcleft, @dstp, @dstleft)
	*cast(wstring ptr, dstp) = 0
	function = dst
end function

''''''''
function dupstr(s as const zstring ptr) as zstring ptr
	dim as zstring ptr d = allocate(len(*s)+1)
	*d = *s
	function = d
end function

''''''''
function splitstr(Text as string, Delim as string, Ret() as string) as long

	var items = 10
	redim RetVal(0 to items-1) as integer
	
	var x = 0
	var p = 0
	do 
		x = InStr(x + 1, Text, Delim)
		if( x > 0 ) then
			if( p >= items ) then
				items += 10
				redim preserve RetVal(0 to items-1)
			end if
			RetVal(p) = x
		end if
		p += 1
	loop until x = 0
	
	var cnt = p - 1
	if( cnt = 0 ) then
		redim Ret(0 to 0)
		ret(0) = text
		return 1
	end if
	
	redim Ret(0 to cnt)
	Ret(0) = Left(Text, RetVal(0) - 1 )
	p = 1
	do until p = cnt
		Ret(p) = mid(Text, RetVal(p - 1) + 1, RetVal(p) - RetVal(p - 1) - 1 )
		p += 1
	loop
	Ret(cnt) = mid(Text, RetVal(cnt - 1) + 1)
	
	return cnt+1
   
end function


'''''''
function loadstrings(fromFile as string, toArray() as string) as boolean
	
	var fnum = FreeFile
	if open(fromFile for input as #fnum) <> 0 then
		return false
	end if

	var items = 10
	redim toArray(0 to items-1)
	
	var i = 0
	do while not eof(fnum)
		if( i >= items ) then
			items += 10
			redim preserve toArray(0 to items-1)
		end if
		
		line input #fnum, toArray(i)
		if len(toArray(i)) = 0 then
			exit do
		end if
		i += 1
	loop
	
	redim preserve toArray(0 to i-1)
	
	close #fnum
	
	return true
end function

function strreplace _
	( _
		byref text as string, _
		byref a as string, _
		byref b as string _
	) as string

	var result = text

	var alen = len(a)
	var blen = len(b)

	var i = 0
	do
		'' Does result contain an occurence of a?
		i = instr(i + 1, result, a)
		if i = 0 then
			exit do
		end if

		'' Cut out a and insert b in its place
		'' result  =  front  +  b  +  back
		var keep = right(result, len(result) - ((i - 1) + alen))
		result = left(result, i - 1)
		result += b
		result += keep

		i += blen - 1
	loop

	function = result
end function


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
function loadstrings(fromFile as string, toArray() as string) as integer
	
	var fnum = FreeFile
	if open(fromFile for input as #fnum) <> 0 then
		return 0
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
	
	close #fnum
	
	if i > 0 then
		redim preserve toArray(0 to i-1)
	else
		erase toArray
	end if
	
	return i
	
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

function strCalcWrapPoints(text as zstring ptr, points() as integer) as integer
	var items = 20
	redim points(0 to items-1)
	var j = 0
	for i as integer = 0 to len(*text)-1
		select case cast(ubyte ptr, text)[i]
		case asc(" ")
			points(j) = i+1
			j += 1
		end select
		
		if j >= items then
			items += items \ 2
			redim preserve points(0 to items-1)
		end if
	next

	return j
end function

	dim shared helvetica_charWidth(0 to 255) as single = {_
		0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,_		'00  ,01  ,02  ,03  ,04  ,05  ,06  ,07
		0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,_		'08  ,09  ,10  ,11  ,12  ,13  ,14  ,15
		0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,_		'16  ,17  ,18  ,19  ,20  ,21  ,22  ,23
		0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,_		'24  ,25  ,26  ,27  ,28  ,29  ,30  ,31
		0.50,0.25,0.25,1.00,1.00,1.00,1.00,0.25,_ 		'    ,!   ,"   ,#   ,$   ,%   ,&   ,'   
		0.25,0.25,0.25,0.25,0.25,0.25,0.25,0.50,_		'(   ,)   ,*   ,+   ,,   ,-   ,.   ,/   
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'0   ,1   ,2   ,3   ,4   ,5   ,6   ,7
		1.00,1.00,0.25,0.25,1.00,1.00,1.00,0.50,_		'8   ,9   ,:   ,;   ,<   ,=   ,>   ,?
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'@   ,A   ,B   ,C   ,D   ,E   ,F   ,G
		1.00,0.25,0.75,1.00,1.00,1.00,1.00,1.00,_		'H   ,I   ,J   ,K   ,L   ,M   ,N   ,O
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'P   ,Q   ,R   ,S   ,T   ,U   ,V   ,W
		1.00,1.00,1.00,0.25,0.25,0.25,0.50,0.50,_		'X   ,Y   ,Z   ,[   ,\   ,]   ,^   ,_
		0.25,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'`   ,a   ,b   ,c   ,d   ,e   ,f   ,g
		1.00,0.25,0.25,1.00,0.25,1.25,1.00,1.00,_		'h   ,i   ,j   ,k   ,l   ,m   ,n   ,o
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'p   ,q   ,r   ,s   ,t   ,u   ,v   ,w
		1.00,1.00,1.00,0.25,0.25,0.25,0.50,0.00,_		'x   ,y   ,z   ,{   ,|   ,}   ,~   , 
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'ß   ,à   ,á   ,â   ,ã   ,ä   ,v   ,æ
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_       'ç   ,è   ,é   ,ê   ,ë   ,ì   ,í   ,î 
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'ï   ,ð   ,ñ   ,ò   ,ó   ,ô   ,õ   ,ö
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'÷   ,ø   ,ù   ,ú   ,û   ,û   ,ü   ,ý
		1.00,1.00,0.00,0.00,0.00,0.00,0.00,0.00 _		'þ   ,ÿ   ,192 ,193 ,mrk1,mrk2,mrk...
	}

function ttfLen(src as const zstring ptr) as double
	var lgt = 0.0
	for i as integer = 0 to len(*src) - 1
		lgt += helvetica_charWidth(cast(ubyte ptr, src)[i])
	next
	return lgt
end function

function ttfSubstr(src as const zstring ptr, byref start as single, maxWidth as single) as string
	var res = ""
	var i = 0
	var width_ = 0.0
	do 
		var c = cast(ubyte ptr, src)[i]
		if c = 0 then
			exit do
		end if
		i += 1
		
		var cw = helvetica_charWidth(c)
		if width_+cw > start+maxWidth then
			start = width_
			exit do
		end if
		
		if width_ >= start then
			res += chr(c)
		end if
		
		width_ += cw
	loop
	return res
end function

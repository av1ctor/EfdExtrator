
#include once "bfile.bi"

#define CHAR2BYTE(c) (c - asc("0"))

constructor bfile()
end constructor  

''''''''
function bfile.abrir(arquivo as string) as Boolean
	fnum = 0
	fpos = 0
	blen = 0
	bptr = NULL
	fnum = FreeFile
   
	function = open(arquivo for binary access read as #fnum) = 0
   
end function

''''''''
function bfile.criar(arquivo as string) as Boolean
	fnum = 0
	fpos = 0
	blen = 0
	bptr = NULL
	fnum = FreeFile
   
	function = open(arquivo for binary access write as #fnum) = 0
   
end function

''''''''
sub bfile.fechar()
	close #fnum
	fnum = 0
	blen = 0
end sub

''''''''
function bfile.tamanho() as longint
   
	function = lof(fnum)

end function

''''''''
function bfile.lerTudo( ) as string

	seek #fnum, 1
	var res = string( tamanho(), asc(" "))
	get #fnum, , res 

	blen = 0
	bptr = null
	
	function = res

end function

''''''''
function bfile.ler(destino() as byte, lgt as integer) as integer
	var i = 0
	do while temProximo() and cbool(i < lgt)
		destino(i) = char1
		i += 1
	loop
	
	function = i
end function

''''''''
function bfile.ler(destino as byte ptr, lgt as integer) as integer
	var i = 0
	do while temProximo() and cbool(i < lgt)
		*destino = char1
		destino += 1
		i += 1
	loop
	
	function = i
end function

''''''''
sub bfile.escrever(texto as string) 

	put #fnum, , texto

end sub

''''''''
sub bfile.escrever(src() as byte) 
	
	put #fnum, , src()
	
end sub

''''''''
function bfile.posicao() as longint

	function = seek(fnum) - blen

end function

''''''''
function bfile.temProximo as Boolean
	if blen > 0 then
		return true
	elseif eof(fnum) = 0 then
		return true
	else
		return false
	end if
end function

''''''''
property bfile.char1() as uByte
      
	property = peek1

	bptr += 1
	blen -= 1
   
end property

''''''''
property bfile.peek1() as uByte

	if blen = 0 then
		fpos = seek( fnum )
		if get( #fnum, , buffer ) = 0 then
			blen = seek( fnum ) - fpos
			bptr = @buffer
		end if
	end if

	property = *bptr
   
end property


''''''''
function bfile.nchar(caracteres as Integer, preenchimento as byte) as string

	var res = string(caracteres, preenchimento)

	for i as integer = 0 to caracteres-1
		res[i] = char1
	next

	function = res
  
end function

''''''''
function bfile.varchar(separador as uInteger) as string

	var res = ""
	var c1 = " "

	do
		c1[0] = char1
		if c1[0] = separador then
			exit do
		end if
		res += c1
	loop

	function = res
  
end function

''''''''
function bfile.varint(separador as uInteger) as longint

	dim as longint res = 0

	do
		var c1 = char1
		if c1 = separador then
			exit do
		end if
		res = res * 10 + CHAR2BYTE(c1)
	loop

	function = res
  
end function

''''''''
function bfile.vardbl(separador as uInteger, decimalSep as uinteger) as double

	dim as longint intp = 0

	dim as integer c1
	do
		c1 = char1
		if c1 = separador then
			exit do
		elseif c1 = decimalSep then
			exit do
		end if
		intp = intp * 10 + CHAR2BYTE(c1)
	loop

	if c1 = decimalSep then
		dim as integer decp = 0
		dim as integer decdiv = 1
		do
			c1 = char1
			if c1 = separador then
				exit do
			end if
				
			decp = decp * 10 + CHAR2BYTE(c1)
			decdiv = decdiv * 10
		loop

		function = cdbl(intp) + (decp / decdiv)
	else
		function = cdbl(intp)
	end if
  
end function

''''''''
property bfile.int1() as integer
   
	property = CHAR2BYTE(char1)
   
end property


''''''''
property bfile.char2() as string
   
	res2[0] = char1
	res2[1] = char1

	property = res2
   
end property

''''''''
property bfile.int2() as integer
   
	property = cint(CHAR2BYTE(char1)) * 10 + CHAR2BYTE(char1)
   
end property

''''''''
property bfile.char4() as string
   
	res4[0] = char1
	res4[1] = char1
	res4[2] = char1
	res4[3] = char1

	property = res4
   
end property

''''''''
property bfile.int4() as integer
   
	property = cint(CHAR2BYTE(char1)) * 1000 + cint(CHAR2BYTE(char1)) * 100 + cint(CHAR2BYTE(char1)) * 10 + CHAR2BYTE(char1)
   
end property

''''''''
property bfile.char6() as string
   
	for i as integer = 0 to 5
		res6[i] = char1
	next

	property = res6
   
end property

''''''''
property bfile.int6() as integer
   
	for i as integer = 0 to 5
		res6[i] = char1
	next

	property = valint(res6)
   
end property

''''''''
property bfile.char8() as string
   
	for i as integer = 0 to 7	
		res8[i] = char1
	next

	property = res8
   
end property

''''''''
property bfile.int9() as integer
   
	for i as integer = 0 to 8
		res9[i] = char1
	next

	property = valint(res9)
   
end property

''''''''
property bfile.char13() as string
   
	for i as integer = 0 to 12
		res13[i] = char1
	next

	property = res13
   
end property

''''''''
property bfile.dbl13_2() as double
   
	for i as integer = 0 to 10
		res14[i] = char1
	next

	res14[11] = asc(".")
	res14[12] = char1
	res14[13] = char1
	
	property = val(res14)
   
end property

''''''''
property bfile.dbl5_2() as double
   
	res6[0] = char1
	res6[1] = char1
	res6[2] = char1
	res6[3] = asc(".")
	res6[4] = char1
	res6[5] = char1

	property = val(res6)
   
end property

''''''''
property bfile.dbl4_2() as double
   
	res5[0] = char1
	res5[1] = char1
	res5[2] = asc(".")
	res5[3] = char1
	res5[4] = char1

	property = val(res5)
   
end property

''''''''
property bfile.dbl11_3() as double
   
	for i as integer = 0 to 7
		res12[i] = char1
	next

	res12[08] = asc(".")
	res12[09] = char1
	res12[10] = char1
	res12[11] = char1

	property = val(res12)
   
end property

''''''''
property bfile.dbl12_2() as double
   
	for i as integer = 0 to 9
		res13[i] = char1
	next

	res13[10] = asc(".")
	res13[11] = char1
	res13[12] = char1

	property = val(res13)
   
end property

''''''''
property bfile.dbl13_3() as double
   
	for i as integer = 0 to 9
		res14[i] = char1
	next

	res14[10] = asc(".")
	res14[11] = char1
	res14[12] = char1
	res14[13] = char1

	property = val(res14)
   
end property

''''''''
property bfile.char14() as string
   
	for i as integer = 0 to 13
		res14[i] = char1
	next

	property = res14
   
end property

''''''''
property bfile.lng14() as longint
   
	for i as integer = 0 to 13
		res14[i] = char1
	next

	property = vallng(res14)
   
end property

''''''''
property bfile.char22() as string
   
	for i as integer = 0 to 21
		res22[i] = char1
	next

	property = res22
   
end property

''''''''
function bfile.charcsv(separador as uInteger, qualificador as uInteger) as string

	var res = ""
	var c1 = " "
   
	'' qualificador?
	if peek1 = qualificador then
		'' pular qualificador
		char1

		do
			c1[0] = char1
			if c1[0] = qualificador then
				'' dois qualificadores, um seguido do outro? considerar como parte do texto
				if peek1 = qualificador then
					char1
				else
					exit do
				end if
			end if
			res += c1
		loop
	end if
   
	select case peek1
	'' separador? pular
	case separador
		char1		
	'' final de linha? deixar
	case 13, 10
	
	'' se não for o separador, e não for final de linha, concatenar ao texto até encontar o separador
	case else
		do
			select case peek1
			case separador
				char1
				exit do
			case 13, 10
				exit do
			end select
			
			c1[0] = char1
			res += c1
		loop
	end select
   
	function = res
  
end function

''''''''
function bfile.intCsv(separador as uInteger, qualificador as uInteger) as longint

	'' qualificador?
	if peek1 = qualificador then
		'' pular qualificador
		char1
		
		function = varint(qualificador)
		
		'' separador? pular
		if peek1 = separador then
			char1		
		end if
	
	'' sem qualificador.. 
	else
		function = valint(charcsv(separador, qualificador))
	end if


end function

''''''''
function bfile.dblCsv(separador as uInteger, qualificador as uInteger) as double


	'' qualificador?
	if peek1 = qualificador then
		'' pular qualificador
		char1
		function = vardbl(qualificador)
	  
		'' separador? pular
		if peek1 = separador then
			char1		
		end if

	'' sem qualificador.. 
	else
		function = val(charcsv(separador, qualificador))
	end if
  
end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''
private function luacb_bf_char1 cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var bf = cast(bfile ptr, lua_touserdata(L, 1))

		lua_pushinteger(L, cuint(bf->char1))
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_bf_int1 cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var bf = cast(bfile ptr, lua_touserdata(L, 1))

		lua_pushinteger(L, bf->int1)
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_bf_int2 cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var bf = cast(bfile ptr, lua_touserdata(L, 1))

		lua_pushinteger(L, bf->int2)
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
private function luacb_bf_varchar cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args < 1 or args > 2 then
		lua_pushnil(L)
	else
		var bf = cast(bfile ptr, lua_touserdata(L, 1))
		var separador = iif(args = 2, cuint(lua_tointeger(L, 2)), asc("|"))

		lua_pushstring(L, bf->varchar(separador))
	end if
	
	function = 1
	
end function

''''''''
private function luacb_bf_vardbl cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args < 1 or args > 2 then
		lua_pushnil(L)
	else
		var bf = cast(bfile ptr, lua_touserdata(L, 1))
		var separador = iif(args = 2, cuint(lua_tointeger(L, 2)), asc("|"))

		lua_pushnumber(L, bf->vardbl(separador))
	end if
	
	function = 1
	
end function

''''''''
private function luacb_bf_varint cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args < 1 or args > 2 then
		lua_pushnil(L)
	else
		var bf = cast(bfile ptr, lua_touserdata(L, 1))
		var separador = iif(args = 2, cuint(lua_tointeger(L, 2)), asc("|"))

		lua_pushinteger(L, bf->varint(separador))
	end if
	
	function = 1
	
end function

''''''''
static sub bfile.exportAPI(L as lua_State ptr)
	
	lua_register(L, "bf_char1", @luacb_bf_char1)
	lua_register(L, "bf_int1", @luacb_bf_int1)
	lua_register(L, "bf_int2", @luacb_bf_int2)
	lua_register(L, "bf_varchar", @luacb_bf_varchar)
	lua_register(L, "bf_varint", @luacb_bf_varint)
	lua_register(L, "bf_vardbl", @luacb_bf_vardbl)
	
end sub
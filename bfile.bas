
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
sub bfile.fechar()
	close #fnum
	fnum = 0
	blen = 0
end sub

function bfile.tamanho() as longint
   
	function = lof(fnum)

end function

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
property bfile.char1() as Byte
      
	property = peek1

	bptr += 1
	blen -= 1
   
end property

''''''''
property bfile.peek1() as Byte

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
function bfile.varchar(separador as Integer) as string

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
function bfile.varint(separador as Integer) as longint

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
function bfile.vardbl(separador as Integer) as double

	dim as longint intp = 0

	dim as integer c1
	do
		c1 = char1
		if c1 = separador then
			exit do
		elseif c1 = asc(",") then
			exit do
		end if
		intp = intp * 10 + CHAR2BYTE(c1)
	loop

	if c1 = asc(",") then
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
property bfile.dbl4_2() as double
   
	for i as integer = 0 to 1
		res5[i] = char1
	next

	res5[2] = asc(".")
	res5[3] = char1
	res5[4] = char1

	property = val(res5)
   
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
function bfile.charcsv(separador as Integer, qualificador as Integer) as string

	var res = ""
	var c1 = " "
   
	'' pular qualificador
	if peek1 = qualificador then
		char1		
	end if
   
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
   	
	select case peek1 
	'' separador? pular
	case separador
		char1		
	'' final de linha? deixar
	case 13, 10
	
	'' se não for o separador, e não for final de linha, concatenar ao texto até encontar o separador
	case else
		do
			c1[0] = char1
			select case c1[0] 
			case separador, 13, 10
				exit do
			end select
			res += c1
		loop
	end select
   
	function = res
  
end function

''''''''
function bfile.intCsv(separador as Integer, qualificador as Integer) as longint

	'' pular qualificador
	if peek1 = qualificador then
		char1		
	end if

	function = varint(qualificador)
	
	'' separador? pular
	if peek1 = separador then
		char1		
	end if

end function

''''''''
function bfile.dblCsv(separador as Integer, qualificador as Integer) as double

	'' pular qualificador
	if peek1 = qualificador then
		char1		
	end if

	function = vardbl(qualificador)
  
	'' separador? pular
	if peek1 = separador then
		char1		
	end if
  
end function

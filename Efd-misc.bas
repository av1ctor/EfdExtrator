#include once "Efd.bi"

dim shared as string ufCod2Sigla(11 to 53)
dim shared as TDict ufSigla2CodDict
dim shared as string codSituacao2Str(0 to __TipoSituacao__LEN__-1)

private sub tablesCtor constructor
	ufCod2Sigla(11)="RO"
	ufCod2Sigla(12)="AC"
	ufCod2Sigla(13)="AM"
	ufCod2Sigla(14)="RR"
	ufCod2Sigla(15)="PA"
	ufCod2Sigla(16)="AP"
	ufCod2Sigla(17)="TO"
	ufCod2Sigla(21)="MA"
	ufCod2Sigla(22)="PI"
	ufCod2Sigla(23)="CE"
	ufCod2Sigla(24)="RN"
	ufCod2Sigla(25)="PB"
	ufCod2Sigla(26)="PE"
	ufCod2Sigla(27)="AL"
	ufCod2Sigla(28)="SE"
	ufCod2Sigla(29)="BA"
	ufCod2Sigla(31)="MG"
	ufCod2Sigla(32)="ES"
	ufCod2Sigla(33)="RJ"
	ufCod2Sigla(35)="SP"
	ufCod2Sigla(41)="PR"
	ufCod2Sigla(42)="SC"
	ufCod2Sigla(43)="RS"
	ufCod2Sigla(50)="MS"
	ufCod2Sigla(51)="MT"
	ufCod2Sigla(52)="GO"
	ufCod2Sigla(53)="DF"
	
	''
	ufSigla2CodDict.init(30)
	for i as integer = 11 to 53
		if len(ufCod2Sigla(i)) > 0 then
			var valor = new VarBox(i)
			ufSigla2CodDict.add(ufCod2Sigla(i), valor)
		end if
	next

	var valor = new VarBox(99)
	ufSigla2CodDict.add("EX", valor)
	
	''
	codSituacao2Str(REGULAR) 			= "REG"
	codSituacao2Str(EXTEMPORANEO) 		= "EXTEMP"
	codSituacao2Str(CANCELADO) 			= "CANC"
	codSituacao2Str(CANCELADO_EXT) 		= "CANC EXTEMP"
	codSituacao2Str(DENEGADO) 			= "DENEG"
	codSituacao2Str(INUTILIZADO) 		= "INUT"
	codSituacao2Str(COMPLEMENTAR) 		= "COMPL"
	codSituacao2Str(COMPLEMENTAR_EXT) 	= "COMPL EXTEMP"
	codSituacao2Str(REGIME_ESPECIAL) 	= "REG ESP"
	codSituacao2Str(SUBSTITUIDO) 		= "SUBST"
end sub

''''''''
function ddMmYyyy2YyyyMmDd(s as const zstring ptr) as string
	
	var res = "19000101"
	
	if len(*s) > 0 then
		res[0] = s[4]
		res[1] = s[5]
		res[2] = s[6]
		res[3] = s[7]
		res[4] = s[2]
		res[5] = s[3]
		res[6] = s[0]
		res[7] = s[1]
	end if
	
	function = res
	
end function

''''''''
function yyyyMmDd2Datetime(s as const zstring ptr) as string 
	''         0123456789
	var res = "1900-01-01T00:00:00.000"
	
	if len(*s) > 0 then
		res[0] = s[0]
		res[1] = s[1]
		res[2] = s[2]
		res[3] = s[3]
		res[5] = s[4]
		res[6] = s[5]
		res[8] = s[6]
		res[9] = s[7]
	end if
	
	function = res
end function

''''''''
function YyyyMmDd2DatetimeBR(s as const zstring ptr) as string 
	''         0123456789
	var res = "01/01/1900"
	
	if len(*s) > 0 then
		res[0] = s[6]
		res[1] = s[7]
		res[3] = s[4]
		res[4] = s[5]
		res[6] = s[0]
		res[7] = s[1]
		res[8] = s[2]
		res[9] = s[3]
	end if
	
	function = res
end function

''''''''
function STR2IE(ie as string) as string
	var ie2 = right(string(12,"0") + ie, 12)
	function = left(ie2,3) + "." + mid(ie2,4,3) + "." + mid(ie2,4+3,3) + "." + right(ie2,3)
end function


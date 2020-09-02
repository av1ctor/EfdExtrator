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

'''''
function UF_SIGLA2COD(s as zstring ptr) as integer
	
	if s = null then
		return 0
	end if
	
	if len(s) = 0 then
		return 0
	end if
	
	var cod = cast(VarBox ptr, ufSigla2CodDict[s])
	if cod = null then
		return 0
	end if
	
	function = cast(integer, *cod)

end function

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
	
	if res = "01/01/1900" then
		res = ""
	end if
	
	function = res
end function

''''''''
function STR2IE(ie as string) as string
	var ie2 = right(string(12,"0") + ie, 12)
	function = left(ie2,3) + "." + mid(ie2,4,3) + "." + mid(ie2,4+3,3) + "." + right(ie2,3)
end function

''''''''
function EFd.codMunicipio2Nome(cod as integer) as string
	
	var nome = cast(zstring ptr, municipDict[cod])
	if nome <> null then
		return *nome
	end if
	
	var nomedb = dbConfig->execScalar("select Nome || ' - ' || uf nome from Municipio where Codigo = " & cod)
	if nomedb = null then
		return ""
	end if
	
	municipDict.add(cod, nomedb)
	
	function = *nomedb
end function

''''''''
function tipoItem2Str(tipo as TipoItemId) as string
	select case as const tipo
	case TI_Mercadoria_para_Revenda
		return "Mercadoria para Revenda"
	case TI_Materia_Prima
		return "Materia Prima"
	case TI_Embalagem
		return "Embalagem"
	case TI_Produto_em_Processo
		return "Produto em Processo"
	case TI_Produto_Acabado
		return "Produto Acabado"
	case TI_Subproduto
		return "Subproduto"
	case TI_Produto_Intermediario
		return "Produto Intermediario"
	case TI_Material_de_Uso_e_Consumo
		return "Material de Uso e Consumo"
	case TI_Ativo_Imobilizado
		return "Ativo Imobilizado"
	case TI_Servicos
		return "Servicos"
	case TI_Outros_insumos
		return "Outros_insumos"
	case else
		return "Outras"
	end select
end function

''''''''
function dupstr(s as const zstring ptr) as zstring ptr
	dim as zstring ptr d = allocate(len(*s)+1)
	*d = *s
	function = d
end function

''''''''
sub splitstr(Text as string, Delim as string, Ret() as string)

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
		return
	end if
	
	redim Ret(0 to cnt)
	Ret(0) = Left(Text, RetVal(0) - 1 )
	p = 1
	do until p = cnt
		Ret(p) = mid(Text, RetVal(p - 1) + 1, RetVal(p) - RetVal(p - 1) - 1 )
		p += 1
	loop
	Ret(cnt) = mid(Text, RetVal(cnt - 1) + 1)
   
end sub

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
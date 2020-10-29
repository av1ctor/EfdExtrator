#include once "Efd.bi"

dim shared as string ufCod2Sigla(11 to 53)
dim shared as TDict ptr ufSigla2CodDict
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
	ufSigla2CodDict = new TDict(30)
	for i as integer = 11 to 53
		if len(ufCod2Sigla(i)) > 0 then
			var valor = new VarBox(i)
			ufSigla2CodDict->add(ufCod2Sigla(i), valor)
		end if
	next

	var valor = new VarBox(99)
	ufSigla2CodDict->add("EX", valor)
	
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
function ISREGULAR(sit as TipoSituacao) as boolean
	select case as const sit
	case REGULAR, EXTEMPORANEO, COMPLEMENTAR, COMPLEMENTAR_EXT, REGIME_ESPECIAL, SUBSTITUIDO
		return true
	case else
		return false
	end select
end function

'''''
function UF_SIGLA2COD(s as zstring ptr) as integer
	
	if s = null then
		return 0
	end if
	
	if len(s) = 0 then
		return 0
	end if
	
	var cod = cast(VarBox ptr, ufSigla2CodDict->lookup(s))
	if cod = null then
		return 0
	end if
	
	function = cast(integer, *cod)

end function

''''''''
'' ddmmyyyy to yyyymmdd
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
'' yyyy-mm-dd to yyyymmdd
function yyyyMmDd2YyyyMmDd(s as const zstring ptr) as string
	
	var res = "19000101"
	
	if len(*s) > 0 then
		res[0] = s[0]
		res[1] = s[1]
		res[2] = s[2]
		res[3] = s[3]
		res[4] = s[5]
		res[5] = s[6]
		res[6] = s[8]
		res[7] = s[9]
	end if
	
	function = res
	
end function

''''''''
'' yyyymmdd to yyyy-mm-ddT00:00:00.000
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
'' yyyymmdd to dd/mm/yyyy
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
'' d[d]/m[m]/yyyy to yyyymmddT00:00:00.000
function csvDate2YYYYMMDD(s as zstring ptr) as string 
	''         01234567
	var res = "00000000T00:00:00.000"
	
	var p = 0
	if s[0+1] = asc("/") then
		res[7] = s[0]
		p += 1+1
	else
		res[6] = s[0]
		res[7] = s[1]
		p += 2+1
	end if

	if s[p+1] = asc("/") then
		res[5] = s[p]
		p += 1+1
	else
		res[4] = s[p]
		res[5] = s[p+1]
		p += 2+1
	end if
	
	res[0] = s[p]
	res[1] = s[p+1]
	res[2] = s[p+2]
	res[3] = s[p+3]
	
	function = res
end function

''''''''
function STR2IE(ie as string) as string
	var ie2 = right(string(12,"0") + ie, 12)
	function = left(ie2,3) + "." + mid(ie2,4,3) + "." + mid(ie2,4+3,3) + "." + right(ie2,3)
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
sub pularLinha(bf as bfile) 

	'ler até \r
	do
		var c = bf.char1
		
		if c = 13 or c = 10 then
			exit do
		end if
	loop

	'pular \n
	if bf.peek1 = 10 then
		bf.char1 
	end if
	
end sub

''''''''
function lerLinha(bf as bfile) as string

	var res = ""
	var c = " "
	
	'ler até \r
	do
		c[0] = bf.char1
		if c[0] = 13 or c[0] = 10 then
			exit do
		end if
		
		res += c
	loop
	
	'pular \n
	if bf.peek1 = 10 then
		bf.char1 
	end if

	function = res
	
end function

''''''''
function filtrarPorCnpj(cnpj as const zstring ptr, listaCnpj() as string) as boolean
	
	for i as integer = 0 to ubound(listaCnpj)
		if(*cnpj = listaCnpj(i)) then
			return true
		end if
	next
	
	function = false
	
end function

''''''''
function filtrarPorChave(chave as const zstring ptr, listaChaves() as string) as boolean
	
	for i as integer = 0 to ubound(listaChaves)
		if(*chave = listaChaves(i)) then
			return true
		end if
	next
	
	function = false
	
end function

''''''''
function codMunicipio2Nome(cod as integer, municipDict as TDict ptr, configDb as TDb ptr) as string
	
	var nome = cast(zstring ptr, municipDict->lookup(cod))
	if nome <> null then
		return *nome
	end if
	
	var nomedb = configDb->execScalar("select Nome || ' - ' || uf nome from Municipio where Codigo = " & cod)
	if nomedb = null then
		return ""
	end if
	
	municipDict->add(cod, nomedb)
	
	function = *nomedb
end function



''''''''
sub lua_setarGlobal overload (lua as lua_State ptr, varName as const zstring ptr, value as integer)
	lua_pushnumber(lua, value)
	lua_setglobal(lua, varName)
end sub

''''''''
sub lua_setarGlobal overload (lua as lua_State ptr, varName as const zstring ptr, value as any ptr)
	lua_pushlightuserdata(lua, value)
	lua_setglobal(lua, varName)
end sub


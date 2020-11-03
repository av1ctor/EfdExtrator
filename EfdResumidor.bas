#include once "Efd.bi"
#include once "EfdResumidor.bi"
#include once "TableWriter.bi"
#include once "vbcompat.bi"
#include once "DB.bi"
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#include once "trycatch.bi"

''''''''
constructor EfdResumidor(opcoes as OpcoesExtracao ptr, tableExp as EfdTabelaExport ptr)
	this.opcoes = opcoes
	this.tableExp = tableExp
end constructor

''''''''
function EfdResumidor.withDBs(db as TDb ptr) as EfdResumidor ptr
	this.db = db
	return @this
end function

''''''''
function EfdResumidor.withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdResumidor ptr
	this.onProgress = onProgress
	this.onError = onError
	return @this
end function

''''''''
function EfdResumidor.withLua(lua as lua_State ptr) as EfdResumidor ptr
	this.lua = lua
	return @this
end function

''''''''
sub EfdResumidor.resumoAddHeaderCfopLRE(ws as TableTable ptr)
	var num = 0
	if opcoes->formatoDeSaida = FT_XLSX then
		var row = ws->AddRow(false, num)
		row->addCell("Resumo por CFOP", 9)
		num += 1
	end if
	
	if opcoes->formatoDeSaida = FT_XLSX then
		ws->addColumn(CT_INTNUMBER)
		ws->addColumn(CT_STRING_UTF8, 45)
		ws->addColumn(CT_STRING_UTF8, 15)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_PERCENT)
		ws->addColumn(CT_PERCENT)
		ws->addColumn(CT_MONEY)
		
		var row = ws->addRow(true, num)
		row->addCell("CFOP")
		row->addCell("Descricao")
		row->addCell("Operacao")
		row->addCell("Vl Oper")
		row->addCell("BC ICMS")
		row->addCell("Vl ICMS")
		row->addCell("RedBC ICMS")
		row->addCell("Aliq ICMS")
		row->addCell("Vl IPI")
	end if
end sub

''''''''
sub EfdResumidor.resumoAddHeaderCstLRE(ws as TableTable ptr)
	var num = 0
	if opcoes->formatoDeSaida = FT_XLSX then
		var row = ws->AddRow(false, num)
		row->addCell("Resumo por CST", 9, 10)
		num += 1
	end if
	
	if opcoes->formatoDeSaida = FT_XLSX then
		ws->addColumn(CT_STRING_UTF8, 4)
		ws->addColumn(CT_INTNUMBER)
		ws->addColumn(CT_STRING_UTF8, 45)
		ws->addColumn(CT_STRING_UTF8, 30)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_PERCENT)
		ws->addColumn(CT_PERCENT)
		ws->addColumn(CT_MONEY)
		
		var row = ws->addRow(true, num)
		row->addCell("CST", 1, 10)
		row->addCell("Origem")
		row->addCell("Tributacao")
		row->addCell("Vl Oper")
		row->addCell("BC ICMS")
		row->addCell("Vl ICMS")
		row->addCell("RedBC ICMS")
		row->addCell("Aliq ICMS")
		row->addCell("Vl IPI")
	end if
end sub

''''''''
sub EfdResumidor.resumoAddHeaderCstCfopLRE(ws as TableTable ptr)
	var num = 0
	if opcoes->formatoDeSaida = FT_XLSX then
		var row = ws->AddRow(false, num)
		row->addCell("Resumo por CST e CFOP", 12, 20)
		num += 1
		ws->addColumn(CT_STRING_UTF8, 4)
	end if
	
	ws->addColumn(CT_INTNUMBER)
	ws->addColumn(CT_STRING_UTF8, 45)
	ws->addColumn(CT_STRING_UTF8, 30)
	ws->addColumn(CT_INTNUMBER)
	ws->addColumn(CT_STRING_UTF8, 45)
	ws->addColumn(CT_STRING_UTF8, 15)
	ws->addColumn(CT_MONEY)
	ws->addColumn(CT_MONEY)
	ws->addColumn(CT_MONEY)
	ws->addColumn(CT_PERCENT)
	ws->addColumn(CT_PERCENT)
	ws->addColumn(CT_MONEY)
	
	var row = ws->addRow(true, num)
	if opcoes->formatoDeSaida = FT_XLSX then
		row->addCell("CST", 1, 20)
	else
		row->addCell("CST")
	end if
	row->addCell("Origem")
	row->addCell("Tributacao")
	row->addCell("CFOP")
	row->addCell("Descricao")
	row->addCell("Operacao")
	row->addCell("Vl Oper")
	row->addCell("BC ICMS")
	row->addCell("Vl ICMS")
	row->addCell("RedBC ICMS")
	row->addCell("Aliq ICMS")
	row->addCell("Vl IPI")
end sub

''''''''
sub EfdResumidor.resumoAddHeaderCfopLRS(ws as TableTable ptr)
	var num = 0
	if opcoes->formatoDeSaida = FT_XLSX then
		var row = ws->AddRow(false, num)
		row->addCell("Resumo por CFOP", 12)
		num += 1
	end if

	if opcoes->formatoDeSaida = FT_XLSX then
		ws->addColumn(CT_INTNUMBER)
		ws->addColumn(CT_STRING_UTF8, 45)
		ws->addColumn(CT_STRING_UTF8, 15)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_PERCENT)
		ws->addColumn(CT_PERCENT)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_PERCENT)
		ws->addColumn(CT_MONEY)
		
		var row = ws->addRow(true, num)
		row->addCell("CFOP")
		row->addCell("Descricao")
		row->addCell("Operacao")
		row->addCell("Vl Oper")
		row->addCell("BC ICMS")
		row->addCell("Vl ICMS")
		row->addCell("RedBC ICMS")
		row->addCell("Aliq ICMS")
		row->addCell("BC ICMS ST")
		row->addCell("Vl ICMS ST")
		row->addCell("Aliq ICMS ST")
		row->addCell("Vl IPI")
	end if
end sub

''''''''
sub EfdResumidor.resumoAddHeaderCstLRS(ws as TableTable ptr)
	var num = 0
	if opcoes->formatoDeSaida = FT_XLSX then
		var row = ws->AddRow(false, num)
		row->addCell("Resumo por CST", 12, 13)
		num += 1
	end if
	
	if opcoes->formatoDeSaida = FT_XLSX then
		ws->addColumn(CT_STRING_UTF8, 4)
		ws->addColumn(CT_INTNUMBER)
		ws->addColumn(CT_STRING_UTF8, 45)
		ws->addColumn(CT_STRING_UTF8, 30)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_PERCENT)
		ws->addColumn(CT_PERCENT)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_MONEY)
		ws->addColumn(CT_PERCENT)
		ws->addColumn(CT_MONEY)
		
		var row = ws->addRow(true, num)
		row->addCell("CST", 1, 13)
		row->addCell("Origem")
		row->addCell("Tributacao")
		row->addCell("Vl Oper")
		row->addCell("BC ICMS")
		row->addCell("Vl ICMS")
		row->addCell("RedBC ICMS")
		row->addCell("Aliq ICMS")
		row->addCell("BC ICMS ST")
		row->addCell("Vl ICMS ST")
		row->addCell("Aliq ICMS ST")
		row->addCell("Vl IPI")
	end if
end sub

''''''''
sub EfdResumidor.resumoAddHeaderCstCfopLRS(ws as TableTable ptr)
	var num = 0
	if opcoes->formatoDeSaida = FT_XLSX then
		var row = ws->AddRow(false, num)
		row->addCell("Resumo por CST e CFOP", 15, 26)
		num += 1
		ws->addColumn(CT_STRING_UTF8, 4)
	end if

	ws->addColumn(CT_INTNUMBER)
	ws->addColumn(CT_STRING_UTF8, 45)
	ws->addColumn(CT_STRING_UTF8, 30)
	ws->addColumn(CT_INTNUMBER)
	ws->addColumn(CT_STRING_UTF8, 45)
	ws->addColumn(CT_STRING_UTF8, 15)
	ws->addColumn(CT_MONEY)
	ws->addColumn(CT_MONEY)
	ws->addColumn(CT_MONEY)
	ws->addColumn(CT_PERCENT)
	ws->addColumn(CT_PERCENT)
	ws->addColumn(CT_MONEY)
	ws->addColumn(CT_MONEY)
	ws->addColumn(CT_PERCENT)
	ws->addColumn(CT_MONEY)
	
	var row = ws->addRow(true, num)
	if opcoes->formatoDeSaida = FT_XLSX then
		row->addCell("CST", 1, 26)
	else
		row->addCell("CST")
	end if
	row->addCell("Origem")
	row->addCell("Tributacao")
	row->addCell("CFOP")
	row->addCell("Descricao")
	row->addCell("Operacao")
	row->addCell("Vl Oper")
	row->addCell("BC ICMS")
	row->addCell("Vl ICMS")
	row->addCell("RedBC ICMS")
	row->addCell("Aliq ICMS")
	row->addCell("BC ICMS ST")
	row->addCell("Vl ICMS ST")
	row->addCell("Aliq ICMS ST")
	row->addCell("Vl IPI")
end sub

''''''''
private sub resumoAddRowLRE(xrow as TableRow ptr, byref drow as TDataSetRow, opcoes as OpcoesExtracao ptr, tipo as TipoResumo)
	select case tipo
	case TR_CFOP
		if opcoes->formatoDeSaida = FT_XLSX then
			xrow->addCell(drow["cfop"])
			xrow->addCell(drow["descricao"])
			xrow->addCell(drow["operacao"])
			xrow->addCell(drow["vlOper"])
			xrow->addCell(drow["bcIcms"])
			xrow->addCell(drow["vlIcms"])
			xrow->addCell(drow["redBcIcms"])
			xrow->addCell(drow["aliqIcms"])
			xrow->addCell(drow["vlIpi"])
		end if
	case TR_CST
		if opcoes->formatoDeSaida = FT_XLSX then
			xrow->addCell(drow["cst"], 1, 10)
			xrow->addCell(drow["origem"])
			xrow->addCell(drow["tributacao"])
			xrow->addCell(drow["vlOper"])
			xrow->addCell(drow["bcIcms"])
			xrow->addCell(drow["vlIcms"])
			xrow->addCell(drow["redBcIcms"])
			xrow->addCell(drow["aliqIcms"])
			xrow->addCell(drow["vlIpi"])
		end if
	case TR_CST_CFOP
		if opcoes->formatoDeSaida = FT_XLSX then
			xrow->addCell(drow["cst"], 1, 20)
		else
			xrow->addCell(drow["cst"])
		end if
		xrow->addCell(drow["origem"])
		xrow->addCell(drow["tributacao"])
		xrow->addCell(drow["cfop"])
		xrow->addCell(drow["descricao"])
		xrow->addCell(drow["operacao"])
		xrow->addCell(drow["vlOper"])
		xrow->addCell(drow["bcIcms"])
		xrow->addCell(drow["vlIcms"])
		xrow->addCell(drow["redBcIcms"])
		xrow->addCell(drow["aliqIcms"])
		xrow->addCell(drow["vlIpi"])
	end select
end sub

''''''''
private sub resumoAddRowLRS(xrow as TableRow ptr, byref drow as TDataSetRow, opcoes as OpcoesExtracao ptr, tipo as TipoResumo)
	select case tipo
	case TR_CFOP
		if opcoes->formatoDeSaida = FT_XLSX then
			xrow->addCell(drow["cfop"])
			xrow->addCell(drow["descricao"])
			xrow->addCell(drow["operacao"])
			xrow->addCell(drow["vlOper"])
			xrow->addCell(drow["bcIcms"])
			xrow->addCell(drow["vlIcms"])
			xrow->addCell(drow["redBcIcms"])
			xrow->addCell(drow["aliqIcms"])
			xrow->addCell(drow["bcIcmsST"])
			xrow->addCell(drow["vlIcmsST"])
			xrow->addCell(drow["aliqIcmsST"])
			xrow->addCell(drow["vlIpi"])
		end if
	case TR_CST
		if opcoes->formatoDeSaida = FT_XLSX then
			xrow->addCell(drow["cst"], 1, 13)
			xrow->addCell(drow["origem"])
			xrow->addCell(drow["tributacao"])
			xrow->addCell(drow["vlOper"])
			xrow->addCell(drow["bcIcms"])
			xrow->addCell(drow["vlIcms"])
			xrow->addCell(drow["redBcIcms"])
			xrow->addCell(drow["aliqIcms"])
			xrow->addCell(drow["bcIcmsST"])
			xrow->addCell(drow["vlIcmsST"])
			xrow->addCell(drow["aliqIcmsST"])
			xrow->addCell(drow["vlIpi"])
		end if
	case TR_CST_CFOP
		if opcoes->formatoDeSaida = FT_XLSX then
			xrow->addCell(drow["cst"], 1, 26)
		else
			xrow->addCell(drow["cst"])
		end if
		xrow->addCell(drow["origem"])
		xrow->addCell(drow["tributacao"])
		xrow->addCell(drow["cfop"])
		xrow->addCell(drow["descricao"])
		xrow->addCell(drow["operacao"])
		xrow->addCell(drow["vlOper"])
		xrow->addCell(drow["bcIcms"])
		xrow->addCell(drow["vlIcms"])
		xrow->addCell(drow["redBcIcms"])
		xrow->addCell(drow["aliqIcms"])
		xrow->addCell(drow["bcIcmsST"])
		xrow->addCell(drow["vlIcmsST"])
		xrow->addCell(drow["aliqIcmsST"])
		xrow->addCell(drow["vlIpi"])
	end select
end sub

''''''''
private function luacb_efd_plan_resumos_AddRow cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 5 then
		var ws = cast(TableTable ptr, lua_touserdata(L, 1))
		var ds = cast(TDataSet ptr, lua_touserdata(L, 2))
		var opcoes = cast(OpcoesExtracao ptr, lua_touserdata(L, 3))
		var tipo = lua_tointeger(L, 4)
		var livro = lua_tointeger(L, 5)

		if livro = TL_SAIDAS then
			resumoAddRowLRS(ws->AddRow(), *ds->row, opcoes, tipo)
		else
			resumoAddRowLRE(ws->AddRow(), *ds->row, opcoes, tipo)
		end if
	end if
	
	function = 0
	
end function

''''''''
private function luacb_efd_plan_resumos_Reset cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var ws = cast(TableTable ptr, lua_touserdata(L, 1))

		ws->setRow(2)
	end if
	
	function = 0
	
end function

''''''''
sub EfdResumidor.executar(safiFornecidoMask as long) 

	'' configurar lua
	lua_register(lua, "efd_plan_resumos_AddRow", @luacb_efd_plan_resumos_AddRow)
	lua_register(lua, "efd_plan_resumos_Reset", @luacb_efd_plan_resumos_Reset)
	
	luaL_dofile(lua, ExePath + "\scripts\resumos.lua")
	
	lua_pushnumber(lua, safiFornecidoMask)
	lua_setglobal(lua, "dfeFornecidoMask")
	
	''
	criarResumosLRE()
	criarResumosLRS()
	
end sub

''''''''
sub EfdResumidor.criarResumosLRE()

	onProgress(!"\tResumos das entradas", 0)
	
	var resumosLRE = tableExp->getPlanilha("resumos LRE")
	
	' CFOP
	resumoAddHeaderCfopLRE(resumosLRE)
	try
		lua_getglobal(lua, "LRE_criarResumoCFOP")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRE)
		lua_pushlightuserdata(lua, opcoes)
		lua_call(lua, 3, 0)
	catch
		onError("no script lua!")
	endtry
	
	if not onProgress(null, 0.33) then
		exit sub
	end if

	' CST
	resumoAddHeaderCstLRE(resumosLRE)
	try
		lua_getglobal(lua, "LRE_criarResumoCST")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRE)
		lua_pushlightuserdata(lua, opcoes)
		lua_call(lua, 3, 0)
	catch
		onError("no script lua!")
	endtry

	if not onProgress(null, 0.66) then
		exit sub
	end if

	' CST e CFOP
	resumoAddHeaderCstCfopLRE(resumosLRE)
	try
		lua_getglobal(lua, "LRE_criarResumoCstCfop")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRE)
		lua_pushlightuserdata(lua, opcoes)
		lua_call(lua, 3, 0)
	catch
		onError("no script lua!")
	endtry
	
	onProgress(null, 1)

end sub

''''''''
sub EfdResumidor.criarResumosLRS()
	
	onProgress(!"\tResumos das saídas", 0)
	
	var resumosLRS = tableExp->getPlanilha("resumos LRS")

	' CFOP
	resumoAddHeaderCfopLRS(resumosLRS)
	try
		lua_getglobal(lua, "LRS_criarResumoCFOP")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRS)
		lua_pushlightuserdata(lua, opcoes)
		lua_call(lua, 3, 0)
	catch
		onError("no script lua!")
	endtry
	
	if not onProgress(null, 0.33) then
		exit sub
	end if

	' CST
	resumoAddHeaderCstLRS(resumosLRS)
	try
		lua_getglobal(lua, "LRS_criarResumoCST")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRS)
		lua_pushlightuserdata(lua, opcoes)
		lua_call(lua, 3, 0)
	catch
		onError("no script lua!")
	endtry
	
	if not onProgress(null, 0.66) then
		exit sub
	end if

	' CST
	resumoAddHeaderCstCfopLRS(resumosLRS)
	try
		lua_getglobal(lua, "LRS_criarResumoCstCfop")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRS)
		lua_pushlightuserdata(lua, opcoes)
		lua_call(lua, 3, 0)
	catch
		onError("no script lua!")
	endtry
	
	onProgress(null, 1)
end sub



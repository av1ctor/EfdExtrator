#include once "efd.bi"
#include once "ExcelWriter.bi"
#include once "vbcompat.bi"
#include once "DB.bi"
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#include once "trycatch.bi"

''''''''
private sub resumoAddHeaderCfopLRE(ws as ExcelWorksheet ptr)
	var row = ws->AddRow(false, 0)
	row->addCell("Resumo por CFOP", 9)
	
	row = ws->addRow(true)
	row->addCell("CFOP")
	row->addCell("Descricao")
	row->addCell("Operacao")
	row->addCell("Vl Oper")
	row->addCell("BC ICMS")
	row->addCell("Vl ICMS")
	row->addCell("RedBC ICMS")
	row->addCell("Aliq ICMS")
	row->addCell("Vl IPI")
	
	ws->AddCellType(CT_INTNUMBER)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_MONEY)
end sub

''''''''
private sub resumoAddHeaderCstLRE(ws as ExcelWorksheet ptr)
	var row = ws->AddRow(false, 0)
	row->addCell("Resumo por CST", 9, 10)
	
	row = ws->addRow(true)
	row->addCell("CST", 1, 10)
	row->addCell("Origem")
	row->addCell("Tributacao")
	row->addCell("Vl Oper")
	row->addCell("BC ICMS")
	row->addCell("Vl ICMS")
	row->addCell("RedBC ICMS")
	row->addCell("Aliq ICMS")
	row->addCell("Vl IPI")
	
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_INTNUMBER)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_MONEY)
end sub

''''''''
private sub resumoAddHeaderCstCfopLRE(ws as ExcelWorksheet ptr)
	var row = ws->AddRow(false, 0)
	row->addCell("Resumo por CST e CFOP", 12, 20)
	
	row = ws->addRow(true)
	row->addCell("CST", 1, 20)
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
	
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_INTNUMBER)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_INTNUMBER)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_MONEY)
end sub

''''''''
private sub resumoAddHeaderCfopLRS(ws as ExcelWorksheet ptr)
	var row = ws->AddRow(false, 0)
	row->addCell("Resumo por CFOP", 12)

	row = ws->addRow(true)
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
	
	ws->AddCellType(CT_INTNUMBER)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_MONEY)
end sub

''''''''
private sub resumoAddHeaderCstLRS(ws as ExcelWorksheet ptr)
	var row = ws->AddRow(false, 0)
	row->addCell("Resumo por CST", 12, 13)

	row = ws->addRow(true)
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
	
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_INTNUMBER)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_MONEY)
end sub

''''''''
private sub resumoAddHeaderCstCfopLRS(ws as ExcelWorksheet ptr)
	var row = ws->AddRow(false, 0)
	row->addCell("Resumo por CST e CFOP", 15, 26)

	row = ws->addRow(true)
	row->addCell("CST", 1, 26)
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
	
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_INTNUMBER)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_INTNUMBER)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_STRING_UTF8)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_PERCENT)
	ws->AddCellType(CT_MONEY)
end sub

''''''''
private sub resumoAddRowLRE(xrow as ExcelRow ptr, byref drow as TDataSetRow, tipo as TipoResumo)
	select case tipo
	case TR_CFOP
		xrow->addCell(drow["cfop"])
		xrow->addCell(drow["descricao"])
		xrow->addCell(drow["operacao"])
		xrow->addCell(drow["vlOper"])
		xrow->addCell(drow["bcIcms"])
		xrow->addCell(drow["vlIcms"])
		xrow->addCell(drow["redBcIcms"])
		xrow->addCell(drow["aliqIcms"])
		xrow->addCell(drow["vlIpi"])
	case TR_CST
		xrow->addCell(drow["cst"], 1, 10)
		xrow->addCell(drow["origem"])
		xrow->addCell(drow["tributacao"])
		xrow->addCell(drow["vlOper"])
		xrow->addCell(drow["bcIcms"])
		xrow->addCell(drow["vlIcms"])
		xrow->addCell(drow["redBcIcms"])
		xrow->addCell(drow["aliqIcms"])
		xrow->addCell(drow["vlIpi"])
	case TR_CST_CFOP
		xrow->addCell(drow["cst"], 1, 20)
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
private sub resumoAddRowLRS(xrow as ExcelRow ptr, byref drow as TDataSetRow, tipo as TipoResumo)
	select case tipo
	case TR_CFOP
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
	case TR_CST
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
	case TR_CST_CFOP
		xrow->addCell(drow["cst"], 1, 26)
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
	
	if args = 4 then
		var ws = cast(ExcelWorksheet ptr, lua_touserdata(L, 1))
		var ds = cast(TDataSet ptr, lua_touserdata(L, 2))
		var tipo = lua_tointeger(L, 3)
		var livro = lua_tointeger(L, 4)

		if livro = TL_SAIDAS then
			resumoAddRowLRS(ws->AddRow(), *ds->row, tipo)
		else
			resumoAddRowLRE(ws->AddRow(), *ds->row, tipo)
		end if
	end if
	
	function = 0
	
end function

''''''''
private function luacb_efd_plan_resumos_Reset cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 1 then
		var ws = cast(ExcelWorksheet ptr, lua_touserdata(L, 1))

		ws->setRow(2)
	end if
	
	function = 0
	
end function

''''''''
sub Efd.criarResumos(mostrarProgresso as ProgressoCB) 

	'' configurar lua
	lua_register(lua, "efd_plan_resumos_AddRow", @luacb_efd_plan_resumos_AddRow)
	lua_register(lua, "efd_plan_resumos_Reset", @luacb_efd_plan_resumos_Reset)
	
	luaL_dofile(lua, ExePath + "\scripts\resumos.lua")
	
	''
	var safiFornecidoMask = iif(nfeDestSafiFornecido, MASK_SAFI_NFE_DEST_FORNECIDO, 0)
	safiFornecidoMask or= iif(nfeEmitSafiFornecido, MASK_SAFI_NFE_EMIT_FORNECIDO, 0)
	safiFornecidoMask or= iif(itemNFeSafiFornecido, MASK_SAFI_ITEM_NFE_FORNECIDO, 0)
	safiFornecidoMask or= iif(cteSafiFornecido, MASK_SAFI_CTE_FORNECIDO, 0)
	
	lua_pushnumber(lua, safiFornecidoMask)
	lua_setglobal(lua, "dfeFornecidoMask")
	
	''
	criarResumosLRE(mostrarProgresso)
	criarResumosLRS(mostrarProgresso)
	
end sub

''''''''
sub Efd.criarResumosLRE(mostrarProgresso as ProgressoCB)

	
	mostrarProgresso(wstr(!"\tResumos das entradas"), 0)
	
	' CFOP
	resumoAddHeaderCfopLRE(resumosLRE)
	try
		lua_getglobal(lua, "LRE_criarResumoCFOP")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRE)
		lua_call(lua, 2, 0)
	catch
		print "Erro no script lua!"
	endtry
	
	mostrarProgresso(null, 0.33)

	' CST
	resumoAddHeaderCstLRE(resumosLRE)
	try
		lua_getglobal(lua, "LRE_criarResumoCST")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRE)
		lua_call(lua, 2, 0)
	catch
		print "Erro no script lua!"
	endtry

	mostrarProgresso(null, 0.66)

	' CST e CFOP
	resumoAddHeaderCstCfopLRE(resumosLRE)
	try
		lua_getglobal(lua, "LRE_criarResumoCstCfop")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRE)
		lua_call(lua, 2, 0)
	catch
		print "Erro no script lua!"
	endtry
	
	mostrarProgresso(null, 1)

end sub

''''''''
sub Efd.criarResumosLRS(mostrarProgresso as ProgressoCB)
	
	mostrarProgresso(wstr(!"\tResumos das saídas"), 0)

	' CFOP
	resumoAddHeaderCfopLRS(resumosLRS)
	try
		lua_getglobal(lua, "LRS_criarResumoCFOP")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRS)
		lua_call(lua, 2, 0)
	catch
		print "Erro no script lua!"
	endtry
	
	mostrarProgresso(null, 0.33)

	' CST
	resumoAddHeaderCstLRS(resumosLRS)
	try
		lua_getglobal(lua, "LRS_criarResumoCST")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRS)
		lua_call(lua, 2, 0)
	catch
		print "Erro no script lua!"
	endtry
	
	mostrarProgresso(null, 0.66)

	' CST
	resumoAddHeaderCstCfopLRS(resumosLRS)
	try
		lua_getglobal(lua, "LRS_criarResumoCstCfop")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, resumosLRS)
		lua_call(lua, 2, 0)
	catch
		print "Erro no script lua!"
	endtry
	
	mostrarProgresso(null, 1)
end sub



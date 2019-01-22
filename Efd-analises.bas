
#include once "efd.bi"
#include once "ExcelWriter.bi"
#include once "vbcompat.bi"
#include once "DB.bi"
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#include once "trycatch.bi"

''''''''
private sub inconsistenciaAddHeader(ws as ExcelWorksheet ptr)
	ws->AddCellType(CT_STRING, "Chave")
	ws->AddCellType(CT_DATE, "Data")
	ws->AddCellType(CT_INTNUMBER, "Modelo")
	ws->AddCellType(CT_STRING, "Serie")
	ws->AddCellType(CT_INTNUMBER, "Numero")
	ws->AddCellType(CT_MONEY, "Valor Operacao")
	ws->AddCellType(CT_INTNUMBER, "Tipo Inconsistencia")
	ws->AddCellType(CT_STRING, "Descricao Inconsistencia")
end sub

''''''''
private sub inconsistenciaAddRow(xrow as ExcelRow ptr, byref drow as TDataSetRow, tIncons as TipoInconsistencia, descricao as const zstring ptr)
	xrow->addCell(drow["chave"])
	xrow->addCell(yyyyMmDd2Datetime(drow["dataEmit"]))
	xrow->addCell(drow["modelo"])
	xrow->addCell(drow["serie"])
	xrow->addCell(drow["numero"])
	xrow->addCell(drow["valorOp"])
	xrow->addCell(tIncons)
	xrow->addCell(*descricao)
end sub

''''''''
private function luacb_efd_plan_inconsistencias_AddRow cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	if args = 4 then
		var ws = cast(ExcelWorksheet ptr, lua_touserdata(L, 1))
		var ds = cast(TDataSet ptr, lua_touserdata(L, 2))
		var tipo = lua_tointeger(L, 3)
		var descricao = lua_tostring(L, 4)

		inconsistenciaAddRow(ws->AddRow(), *ds->row, tipo, descricao)
	end if
	
	function = 0
	
end function

''''''''
sub Efd.analisar(mostrarProgresso as ProgressoCB) 

	'' configurar lua
	lua_register(lua, "efd_plan_inconsistencias_AddRow", @luacb_efd_plan_inconsistencias_AddRow)
	
	luaL_dofile(lua, ExePath + "\scripts\analises.lua")
	
	''
	var safiFornecidoMask = iif(nfeDestSafiFornecido, MASK_SAFI_NFE_DEST_FORNECIDO, 0)
	safiFornecidoMask or= iif(nfeEmitSafiFornecido, MASK_SAFI_NFE_EMIT_FORNECIDO, 0)
	safiFornecidoMask or= iif(itemNFeSafiFornecido, MASK_SAFI_ITEM_NFE_FORNECIDO, 0)
	safiFornecidoMask or= iif(cteSafiFornecido, MASK_SAFI_CTE_FORNECIDO, 0)
	
	lua_pushnumber(lua, safiFornecidoMask)
	lua_setglobal(lua, "dfeFornecidoMask")
	
	''
	analisarInconsistenciasLRE(mostrarProgresso)
	analisarInconsistenciasLRS(mostrarProgresso)
	
end sub

''''''''
sub Efd.analisarInconsistenciasLRE(mostrarProgresso as ProgressoCB)

	inconsistenciaAddHeader(inconsistenciasLRE)
	
	mostrarProgresso(wstr(!"\tInconsistências nas entradas"), 0)
	
	try
		lua_getglobal(lua, "analisarInconsistenciasLRE")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, inconsistenciasLRE)
		lua_call(lua, 2, 0)
	catch
		print "Erro no script lua!"
	endtry
	
	mostrarProgresso(null, 1)

end sub

''''''''
sub Efd.analisarInconsistenciasLRS(mostrarProgresso as ProgressoCB)
	
	inconsistenciaAddHeader(inconsistenciasLRS)
	
	mostrarProgresso(wstr(!"\tInconsistências nas saídas"), 0)

	try
		lua_getglobal(lua, "analisarInconsistenciasLRS")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, inconsistenciasLRS)
		lua_call(lua, 2, 0)
	catch
		print "Erro no script lua!"
	endtry
	
	mostrarProgresso(null, 1)
end sub



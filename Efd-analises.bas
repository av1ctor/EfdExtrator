
#include once "efd.bi"
#include once "ExcelWriter.bi"
#include once "vbcompat.bi"
#include once "DB.bi"
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#include once "trycatch.bi"

''''''''
private sub inconsistenciaAddHeader(ws as ExcelWorksheet ptr)
	var row = ws->addRow(true)
	row->addCell("Chave")
	row->addCell("Data")
	row->addCell("CNPJ")
	row->addCell("UF")
	row->addCell("Modelo")
	row->addCell("Serie")
	row->addCell("Numero")
	row->addCell("Valor Operacao")
	row->addCell("Tipo Inconsistencia")
	row->addCell("Descricao Inconsistencia")
	
	ws->AddCellType(CT_STRING, 45)
	ws->AddCellType(CT_DATE)
	ws->AddCellType(CT_STRING)
	ws->AddCellType(CT_STRING, 4)
	ws->AddCellType(CT_STRING, 4)
	ws->AddCellType(CT_STRING, 6)
	ws->AddCellType(CT_INTNUMBER)
	ws->AddCellType(CT_MONEY)
	ws->AddCellType(CT_INTNUMBER)
	ws->AddCellType(CT_STRING, 60)
end sub

''''''''
private sub inconsistenciaAddRow(xrow as ExcelRow ptr, byref drow as TDataSetRow, tIncons as TipoInconsistencia, descricao as const zstring ptr)
	xrow->addCell(drow["chave"])
	xrow->addCell(yyyyMmDd2Datetime(drow["dataEmit"]))
	xrow->addCell(drow["cnpj"])
	xrow->addCell(drow["uf"])
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
sub Efd.analisar() 

	'' configurar lua
	lua_register(lua, "efd_plan_inconsistencias_AddRow", @luacb_efd_plan_inconsistencias_AddRow)
	
	luaL_dofile(lua, ExePath + "\scripts\analises.lua")
	
	''
	var safiFornecidoMask = iif(nfeDestSafiFornecido, MASK_BO_NFe_DEST_FORNECIDO, 0)
	safiFornecidoMask or= iif(nfeEmitSafiFornecido, MASK_BO_NFe_EMIT_FORNECIDO, 0)
	safiFornecidoMask or= iif(itemNFeSafiFornecido, MASK_BO_ITEM_NFE_FORNECIDO, 0)
	safiFornecidoMask or= iif(cteSafiFornecido, MASK_BO_CTe_FORNECIDO, 0)
	
	lua_pushnumber(lua, safiFornecidoMask)
	lua_setglobal(lua, "dfeFornecidoMask")
	
	''
	analisarInconsistenciasLRE()
	analisarInconsistenciasLRS()
	
end sub

''''''''
sub Efd.analisarInconsistenciasLRE()

	inconsistenciaAddHeader(inconsistenciasLRE)
	
	onProgress(!"\tInconsistências nas entradas", 0)
	
	try
		lua_getglobal(lua, "LRE_analisarInconsistencias")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, inconsistenciasLRE)
		lua_call(lua, 2, 0)
	catch
		onError("Erro no script lua!")
	endtry
	
	onProgress(null, 1)

end sub

''''''''
sub Efd.analisarInconsistenciasLRS()
	
	inconsistenciaAddHeader(inconsistenciasLRS)
	
	onProgress(!"\tInconsistências nas saídas", 0)

	try
		lua_getglobal(lua, "LRS_analisarInconsistencias")
		lua_pushlightuserdata(lua, db)
		lua_pushlightuserdata(lua, inconsistenciasLRS)
		lua_call(lua, 2, 0)
	catch
		onError("Erro no script lua!")
	endtry
	
	onProgress(null, 1)
end sub



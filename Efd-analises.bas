
#include once "efd.bi"
#include once "ExcelWriter.bi"
#include once "vbcompat.bi"
#include once "DB.bi"
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"

''''''''
private sub inconsistenciaAddHeader(ws as ExcelWorksheet ptr)
	ws->AddCellType(CT_STRING, "Chave")
	ws->AddCellType(CT_DATE, "Data")
	ws->AddCellType(CT_INTNUMBER, "Modelo")
	ws->AddCellType(CT_INTNUMBER, "Serie")
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
private sub lua_setarGlobal(lua as lua_State ptr, varName as const zstring ptr, value as integer)
	lua_pushnumber(lua, value)
	lua_setglobal(lua, varName)
end sub

''''''''
private function luacb_inconsistencia_AddRow cdecl(byval L as lua_State ptr) as long
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
	lua_setarGlobal(lua, "TI_ESCRIT_FALTA", TI_ESCRIT_FALTA)
	lua_setarGlobal(lua, "TI_ESCRIT_FANTASMA", TI_ESCRIT_FANTASMA)
	lua_setarGlobal(lua, "TI_ALIQ", TI_ALIQ)
	lua_setarGlobal(lua, "TI_DUP", TI_DUP)
	lua_setarGlobal(lua, "TI_DIF", TI_DIF)
	
	lua_register(lua, "inconsistencia_AddRow", @luacb_inconsistencia_AddRow)
	
	luaL_dofile(lua, ExePath + "\scripts\analises.lua")
	
	''
	if not (nfeDestSafiFornecido or nfeEmitSafiFornecido or itemNFeSafiFornecido or cteSafiFornecido) then
		print wstr(!"\tN�o ser� possivel realizar an�lises e cruzamentos porque os relat�rios Infoview BO do SAFI n�o foram fornecidos")
	else
		analisarInconsistenciasLRE(mostrarProgresso)
		analisarInconsistenciasLRS(mostrarProgresso)
	end if
	
end sub

''''''''
sub Efd.analisarInconsistenciasLRE(mostrarProgresso as ProgressoCB)

	var ws = ew->AddWorksheet("Inconsistencias LRE")
	inconsistenciaAddHeader(ws)
	
	mostrarProgresso(wstr(!"\tInconsist�ncias nas entradas"), 0)
	
	lua_getglobal(lua, "analisarInconsistenciasLRE")
	lua_pushlightuserdata(lua, db)
	lua_pushlightuserdata(lua, ws)
	lua_call(lua, 2, 0)
	
	mostrarProgresso(null, 1)

end sub

''''''''
sub Efd.analisarInconsistenciasLRS(mostrarProgresso as ProgressoCB)
	
	var ws = ew->AddWorksheet("Inconsistencias LRS")
	inconsistenciaAddHeader(ws)
	
	mostrarProgresso(wstr(!"\tInconsist�ncias nas sa�das"), 0)

	lua_getglobal(lua, "analisarInconsistenciasLRS")
	lua_pushlightuserdata(lua, db)
	lua_pushlightuserdata(lua, ws)
	lua_call(lua, 2, 0)
	
	mostrarProgresso(null, 1)
end sub



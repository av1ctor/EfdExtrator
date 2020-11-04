
#include once "EfdExt.bi"
#include once "EfdSpedImport.bi"
#include once "EfdSintegraImport.bi"
#include once "EfdBoCsvLoader.bi"
#include once "EfdBoXlsxLoader.bi"
#include once "EfdAnalisador.bi"
#include once "EfdResumidor.bi"
#include once "EfdPdfExport.bi"
#include once "bfile.bi"
#include once "Dict.bi"
#include once "TableWriter.bi"
#include once "SQLite.bi"
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#include once "trycatch.bi"
#undef imp

private function my_lua_Alloc cdecl _
	( _
		byval ud as any ptr, _
		byval p as any ptr, _
		byval osize as uinteger, _
		byval nsize as uinteger _
	) as any ptr

	if( nsize = 0 ) then
		deallocate( p )
		function = NULL
	else
		function = reallocate( p, nsize )
	end if

end function

''''''''
constructor EfdExt(onProgress as OnProgressCB, onError as OnErrorCB)
	
	'' eventos
	this.onProgress = onProgress
	this.onError = onError
	
	''
	baseTemplatesDir = ExePath() + "\templates\"
	
	municipDict = new TDict(2^10, true, true, true)
	
	''
	configDb = new SQLite
	configDb->open(ExePath + "\db\config.db")
	
end constructor

destructor EfdExt()

	''
	configDb->close()
	delete configDb
	
	''
	delete municipDict
	
end destructor

''''''''
private sub lua_carregarCustoms(d as TDict ptr, L as lua_State ptr) 

	lua_getglobal(L, "getCustomCallbacks")
	lua_call(L, 0, 1)
	if lua_isnil(L, -1) = 0 then
		lua_pushnil(L)
		do while lua_next(L, -2) <> 0
			var key = lua_tostring(L, -2)
			
			var lcb = new CustomLuaCb
			lua_pushnil(L)
			do while lua_next(L, -2) <> 0
				
				var funct = dupstr(lua_tostring(L, -1)) 
				select case *lua_tostring(L, -2)
				case "reader"
					lcb->reader = funct
				case "writer"
					lcb->writer = funct
				case "rel_entradas"
					lcb->rel_entradas = funct
				case "rel_saidas"
					lcb->rel_saidas = funct
				case "rel_outros"
					lcb->rel_outros = funct
				end select
				
				d->add(key, lcb)
				lua_pop(L, 1)
			loop
				
			lua_pop(L, 1)
		loop
		lua_pop(L, lua_gettop(L))
	end if

end sub

''''''''
sub EfdExt.configurarScripting()
	try
		lua = lua_newstate(@my_lua_Alloc, NULL)
		luaL_openlibs(lua)
		
		SQLite.exportAPI(lua)
		TableWriter.exportAPI(lua)
		bfile.exportAPI(lua)
		exportAPI(lua)

		luaL_dofile(lua, ExePath + "\scripts\config.lua")
		luaL_dofile(lua, ExePath + "\scripts\customizacao.lua")	
		
		customLuaCbDict = new TDict(16, true, true, true)
		lua_carregarCustoms(customLuaCbDict, lua)
	catch
		onError("ao carregar script lua. Verifique erros de sintaxe")
	endtry
end sub

''''''''
function EfdExt.iniciar(nomeArquivo as String, opcoes as OpcoesExtracao) as boolean
	
	''
	nomeArquivoSaida = nomeArquivo
	this.opcoes = opcoes
	
	''
	configurarScripting()

	''
	configurarDB()
	
	''
	loaderCtx = new EfdBoLoaderContext()
	
	''
	exp = (new EfdTabelaExport(nomeArquivo, @opcoes)) _
		->withCallbacks(onProgress, onError) _
		->withLua(lua, customLuaCbDict) _
		->withFiltros(@filtrarPorCnpj, @filtrarPorChave)
		
	return exp->criar()
end function

''''''''
sub EfdExt.finalizar()

	''
	if exp <> null then
		exp->finalizar()
		delete exp
	end if
	
	''
	delete loaderCtx
   
	''
	fecharDb()
	if opcoes.dbEmDisco then
		if not opcoes.manterDb then
			kill dbName + ".db"
		end if
	end if
	
	''
	lua_close( lua )
	
end sub

sub EfdExt.descarregarDFe()
	loaderCtx->descarregar()
end sub

function EfdExt.carregarCsv(nomeArquivo as string) as boolean
	var loader = (new EfdBoCsvLoader(loaderCtx, @opcoes)) _
		->withCallbacks(onProgress, onError) _
		->withDBs(db) _
		->withStmts(db_dfeEntradaInsertStmt, db_dfeSaidaInsertStmt, db_itensDfeSaidaInsertStmt)
	
	function = loader->carregar(nomeArquivo)
	
	delete loader
end function

function EfdExt.carregarXlsx(nomeArquivo as string) as boolean
	var loader = (new EfdBoXlsxLoader(loaderCtx, @opcoes)) _
		->withCallbacks(onProgress, onError) _
		->withDBs(db) _
		->withStmts(db_dfeEntradaInsertStmt, db_dfeSaidaInsertStmt, db_itensDfeSaidaInsertStmt)
	
	function = loader->carregar(nomeArquivo)
	
	delete loader
end function

function EfdExt.carregarTxt(nomeArquivo as string) as EfdBaseImport_ ptr
	
	var imp = cast(EfdBaseImport_ ptr, null)
	
	if instr(nomeArquivo, "SpedEFD") >= 0 then
		imp = (new EfdSpedImport(@opcoes)) _
			->withStmts(this.db_LREInsertStmt, db_itensNfLRInsertStmt, db_LRSInsertStmt, db_analInsertStmt, _
				db_ressarcStItensNfLRSInsertStmt, db_itensIdInsertStmt, db_mestreInsertStmt) _
			->withCallbacks(onProgress, onError) _
			->withLua(lua, customLuaCbDict) _
			->withDBs(db)
	
	elseif instr(nomeArquivo, "SP_") >= 0 then
		imp = (new EfdSintegraImport(@opcoes)) _
			->withCallbacks(onProgress, onError) _
			->withLua(lua, customLuaCbDict) _
			->withDBs(db)
	else
		return null
	end if
	
	if imp->carregar(nomeArquivo) then
		return imp
	else
		delete imp
		return null
	end if
	
end function

''''''''
function EfdExt.processar(imp as EfdBaseImport_ ptr, nomeArquivo as string) as Boolean
   
	if opcoes.formatoDeSaida <> FT_NULL then
		exp ->withState(loaderCtx->itemNFeSafiFornecido) _
			->withDicionarios(imp->getParticipanteDict(), imp->getItemIdDict(), loaderCtx->getChaveDFeDict(), _
				imp->getInfoComplDict(), imp->getObsLancamentoDict(), imp->getBemCiapDict()) _
			->gerar(imp->getFirstReg(), imp->getMestreReg(), imp->getNroRegs())
	else
		onProgress(null, 1)
	end if
	
	if opcoes.gerarRelatorios then
		if imp->getTipoArquivo() = TIPO_ARQUIVO_EFD then
			var infAssinatura = cast(EfdSpedImport ptr, imp)->lerInfoAssinatura(nomeArquivo)
		
			var rel = (new EfdPdfExport(baseTemplatesDir, infAssinatura, @opcoes)) _
				->withDBs(configDb) _
				->withCallbacks(onProgress, onError) _
				->withLua(lua, customLuaCbDict) _
				->withFiltros(@filtrarPorCnpj, @filtrarPorChave) _
				->withDicionarios(imp->getParticipanteDict(), imp->getItemIdDict(), loaderCtx->getChaveDFeDict(), imp->getInfoComplDict(), _
					imp->getObsLancamentoDict(), imp->getBemCiapDict(), imp->getContaContabDict(), imp->getCentroCustoDict(), _
					municipDict)
				
			rel->gerar(imp->getFirstReg(), imp->getMestreReg(), imp->getNroRegs())
			
			delete rel
			
			if infAssinatura <> NULL then
				delete infAssinatura
			end if
		end if
	end if
	
	function = true
end function

''''''''
function EfdExt.getDfeMask() as long
	return iif(loaderCtx->nfeDestSafiFornecido, MASK_BO_NFe_DEST_FORNECIDO, 0) or _
		iif(loaderCtx->nfeEmitSafiFornecido, MASK_BO_NFe_EMIT_FORNECIDO, 0) or _
		iif(loaderCtx->itemNFeSafiFornecido, MASK_BO_ITEM_NFE_FORNECIDO, 0) or _
		iif(loaderCtx->cteSafiFornecido, MASK_BO_CTe_FORNECIDO, 0)
end function

''''''''
sub EfdExt.analisar() 
	var anal = (new EfdAnalisador(exp)) _
		->withDBs(db) _
		->withCallbacks(onProgress, onError) _
		->withLua(lua)
	
	anal->executar(getDfeMask())
	delete anal
end sub

''''''''
sub EfdExt.resumir() 
	var res = (new EfdResumidor(@opcoes, exp)) _
		->withDBs(db) _
		->withCallbacks(onProgress, onError) _
		->withLua(lua)
	
	res->executar(getDfeMask())
	delete res
end sub

''''''''
private function lua_criarTabela(lua as lua_State ptr, db as SQLite ptr, tabela as const zstring ptr, onError as OnErrorCB) as SQLiteStmt ptr

	try
		lua_getglobal(lua, "criarTabela_" + *tabela)
		lua_pushlightuserdata(lua, db)
		lua_call(lua, 1, 1)
		var res = db->prepare(lua_tostring(lua, -1))
		if res = null then
			onError("ao executar script lua de criação de tabela: " + "criarTabela_" + *tabela + ": " + *db->getErrorMsg())
		end if
		function = res
		lua_pop(lua, 1)
	catch
		onError("ao executar script lua de criação de tabela: " + "criarTabela_" + *tabela + ". Verifique erros de sintaxe")
	endtry

end function

''''''''
sub EfdExt.configurarDB()

	db = new SQLite
	if not opcoes.dbEmDisco then
		db->open()
	else
		if opcoes.manterDb then
			dbName = "__inter__"
		else
			dbName = "__inter__" & cint(rnd * 1000000)
		end if
		kill dbName + ".db"
		db->open(dbName + ".db")
		db->execNonQuery("PRAGMA JOURNAL_MODE=OFF")
		db->execNonQuery("PRAGMA SYNCHRONOUS=0")
		db->execNonQuery("PRAGMA LOCKING_MODE=EXCLUSIVE")
	end if

	var dbPath = ExePath + "\db\"
	
	try
		'' chamar configurarDB()
		lua_getglobal(lua, "configurarDB")
		lua_pushlightuserdata(lua, db)
		lua_pushstring(lua, dbPath)
		lua_call(lua, 2, 0)

		'' criar tabelas
		db_dfeEntradaInsertStmt = lua_criarTabela(lua, db, "DFe_Entradas", onError)

		db_dfeSaidaInsertStmt = lua_criarTabela(lua, db, "DFe_Saidas", onError)
		
		db_itensDfeSaidaInsertStmt = lua_criarTabela(lua, db, "DFe_Saidas_Itens", onError)
		
		db_LREInsertStmt = lua_criarTabela(lua, db, "EFD_LRE", onError)

		db_itensNfLRInsertStmt = lua_criarTabela(lua, db, "EFD_Itens", onError)

		db_LRSInsertStmt = lua_criarTabela(lua, db, "EFD_LRS", onError)
		
		db_analInsertStmt = lua_criarTabela(lua, db, "EFD_Anal", onError)

		db_ressarcStItensNfLRSInsertStmt = lua_criarTabela(lua, db, "EFD_Ressarc_Itens", onError)
		
		db_itensIdInsertStmt = lua_criarTabela(lua, db, "EFD_ItensId", onError)
		
		db_mestreInsertStmt = lua_criarTabela(lua, db, "EFD_Mestre", onError)
		
		if db_dfeEntradaInsertStmt = null or _
			db_dfeSaidaInsertStmt = null or _
			 db_itensDfeSaidaInsertStmt = null or _
			  db_LREInsertStmt = null or _
			   db_itensNfLRInsertStmt = null or _
			    db_LRSInsertStmt = null or _
				 db_ressarcStItensNfLRSInsertStmt = null or _
					db_itensIdInsertStmt = null or _
						db_analInsertStmt = null then
			
		end if
	catch
		onError("ao executar script lua de criação de DB. Verifique erros de sintaxe")
	endtry

end sub   

''''''''
sub EfdExt.fecharDb()
	if db <> null then
		if db_dfeEntradaInsertStmt <> null then
			delete db_dfeEntradaInsertStmt
		end if
		if db_dfeSaidaInsertStmt <> null then
			delete db_dfeSaidaInsertStmt
		end if
		if db_itensDfeSaidaInsertStmt <> null then
			delete db_itensDfeSaidaInsertStmt
		end if
		if db_LREInsertStmt <> null then
			delete db_LREInsertStmt
		end if
		if db_itensNfLRInsertStmt <> null then
			delete db_itensNfLRInsertStmt
		end if
		if db_LRSInsertStmt <> null then
			delete db_LRSInsertStmt
		end if
		if db_analInsertStmt <> null then
			delete db_analInsertStmt
		end if
		if db_ressarcStItensNfLRSInsertStmt <> null then
			delete db_ressarcStItensNfLRSInsertStmt
		end if
		if db_itensIdInsertStmt <> null then
			delete db_itensIdInsertStmt
		end if
		if db_mestreInsertStmt <> null then
			delete db_mestreInsertStmt
		end if
		
		db->close()
		delete db
		db = null
	end if
end sub

''''''''
private function luacb_efd_plan_get cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	lua_getglobal(L, "efd")
	var g_efd = cast(EfdExt ptr, lua_touserdata(L, -1))
	lua_pop(L, 1)
	
	if args = 1 then
		var planName = lua_tostring(L, 1)

		var plan = g_efd->exp->getPlanilha(planName)
		if plan <> null then
			lua_pushlightuserdata(L, plan)
		else
			lua_pushnil(L)
		end if
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

private function luacb_efd_onProgress cdecl(L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	lua_getglobal(L, "efd")
	var g_efd = cast(EfdExt ptr, lua_touserdata(L, -1))
	lua_pop(L, 1)
	
	if args = 2 then
		var stt = cast(zstring ptr, lua_tostring(L, 1))
		var prog = lua_tonumber(L, 2)
		lua_pushboolean(L, g_efd->onProgress(stt, prog))
	else
		lua_pushboolean(L, false)
	end if
	
	function = 1
end function

private function luacb_efd_onError cdecl(L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	lua_getglobal(L, "efd")
	var g_efd = cast(EfdExt ptr, lua_touserdata(L, -1))
	lua_pop(L, 1)
	
	if args = 1 then
		var msg = cast(zstring ptr, lua_tostring(L, 1))
		g_efd->onError(msg)
	end if
	
	function = 0
end function

''''''''
static function EfdExt.luacb_efd_participante_get cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)

	lua_getglobal(L, "efd")
	var g_efd = cast(EfdExt ptr, lua_touserdata(L, -1))
	lua_pop(L, 1)
	
	if args = 2 then
		var idParticipante = lua_tostring(L, 1)
		var formatar = lua_toboolean(L, 2) <> 0

		var part = cast( TParticipante ptr, /'g_efd->imp->getParticipanteDict()->lookup(idParticipante)'/ null )
		if part <> null then
			lua_newtable(L)
			lua_pushstring(L, "cnpj")
			lua_pushstring(L, iif(formatar, iif(len(part->cpf) > 0, STR2CPF(part->cpf), STR2CNPJ(part->cnpj)), iif(len(part->cpf) > 0, part->cpf, part->cnpj)))
			lua_settable(L, -3)
			lua_pushstring(L, "ie")
			lua_pushstring(L, iif(formatar, STR2IE(part->ie), part->ie))
			lua_settable(L, -3)
			lua_pushstring(L, "uf")
			lua_pushstring(L, MUNICIPIO2SIGLA(part->municip))
			lua_settable(L, -3)
			lua_pushstring(L, "municip")
			lua_pushstring(L, codMunicipio2Nome(part->municip, g_efd->municipDict, g_efd->configDb))
			lua_settable(L, -3)			
			lua_pushstring(L, "nome")
			lua_pushstring(L, iif(formatar, left(part->nome, 64), part->nome))
			lua_settable(L, -3)
		else
			lua_pushnil(L)
		end if
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
sub EfdExt.exportAPI(L as lua_State ptr)
	
	lua_setarGlobal(L, "TI_ESCRIT_FALTA", TI_ESCRIT_FALTA)
	lua_setarGlobal(L, "TI_ESCRIT_FANTASMA", TI_ESCRIT_FANTASMA)
	lua_setarGlobal(L, "TI_ALIQ", TI_ALIQ)
	lua_setarGlobal(L, "TI_DUP", TI_DUP)
	lua_setarGlobal(L, "TI_DIF", TI_DIF)
	lua_setarGlobal(L, "TI_RESSARC_ST", TI_RESSARC_ST)
	lua_setarGlobal(L, "TI_CRED", TI_CRED)
	lua_setarGlobal(L, "TI_SEL", TI_SEL)
	lua_setarGlobal(L, "TI_DEB", TI_DEB)
	
	lua_setarGlobal(L, "TL_ENTRADAS", TL_ENTRADAS)
	lua_setarGlobal(L, "TL_SAIDAS", TL_SAIDAS)

	lua_setarGlobal(L, "TR_CFOP", TR_CFOP)
	lua_setarGlobal(L, "TR_CST", TR_CST)
	lua_setarGlobal(L, "TR_CST_CFOP", TR_CST_CFOP)

	lua_setarGlobal(L, "DFE_NFE_DEST_FORNECIDO", MASK_BO_NFe_DEST_FORNECIDO)
	lua_setarGlobal(L, "DFE_NFE_EMIT_FORNECIDO", MASK_BO_NFe_EMIT_FORNECIDO)
	lua_setarGlobal(L, "DFE_ITEM_NFE_FORNECIDO", MASK_BO_ITEM_NFE_FORNECIDO)
	lua_setarGlobal(L, "DFE_CTE_FORNECIDO", MASK_BO_CTe_FORNECIDO)
	
	lua_setarGlobal(L, "efd", @this)
	
	lua_register(L, "efd_plan_get", @luacb_efd_plan_get)
	'lua_register(L, "efd_participante_get", @luacb_efd_participante_get)
	'lua_register(L, "efd_rel_addItemAnalitico", @luacb_efd_rel_addItemAnalitico)
	lua_register(L, "onProgress", @luacb_efd_onProgress)
	lua_register(L, "onError", @luacb_efd_onError)
	
end sub
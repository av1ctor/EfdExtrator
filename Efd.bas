
#include once "efd.bi"
#include once "bfile.bi"
#include once "Dict.bi"
#include once "ExcelWriter.bi"
#include once "DB.bi"
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#include once "trycatch.bi"

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
constructor Efd(onProgress as OnProgressCB, onError as OnErrorCB)
	
	'' eventos
	this.onProgress = onProgress
	this.onError = onError
	
	''
	chaveDFeDict = new TDict(2^20)
	nfeDestSafiFornecido = false
	nfeEmitSafiFornecido = false
	itemNFeSafiFornecido = false
	cteSafiFornecido = false
	dfeListHead = null
	dfeListTail = null
	arquivos = new TList(10, len(TArquivoInfo))
	
	''
	baseTemplatesDir = ExePath + "\templates\"
	
	municipDict = new TDict(2^10, true, true, true)
	
	''
	dbConfig = new TDb
	dbConfig->open(ExePath + "\db\config.db")
	
end constructor

destructor Efd()

	''
	descarregarDFe()

	''
	dbConfig->close()
	delete dbConfig
	
	''
	delete municipDict
	
	''
	
	delete arquivos
	
end destructor

sub Efd.descarregarDFe
	if chaveDFeDict <> null then
		delete chaveDFeDict
		chaveDFeDict = null
	end if
	
	do while dfeListHead <> null
		var next_ = dfeListHead->next_
		select case dfeListHead->modelo
		case NFE, SAT
			var head = dfeListHead->nfe.itemListHead
			do while head <> null
				var next_ = head->next_
				delete head
				head = next_
			loop
		end select
		delete dfeListHead
		dfeListHead = next_
	loop
end sub

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
sub EFd.configurarScripting()
	try
		lua = lua_newstate(@my_lua_Alloc, NULL)
		luaL_openlibs(lua)
		
		TDb.exportAPI(lua)
		ExcelWriter.exportAPI(lua)
		bfile.exportAPI(lua)
		exportAPI(lua)

		luaL_dofile(lua, ExePath + "\scripts\config.lua")
		luaL_dofile(lua, ExePath + "\scripts\customizacao.lua")	
		
		customLuaCbDict = new TDict(16, true, true, true)
		lua_carregarCustoms(customLuaCbDict, lua)
	catch
		onError("Erro ao carregar script lua. Verifique erros de sintaxe")
	endtry
end sub

''''''''
sub Efd.iniciarExtracao(nomeArquivo as String, opcoes as OpcoesExtracao)
	
	''
	ew = new ExcelWriter
	ew->create(nomeArquivo, opcoes.formatoDeSaida)

	entradas = null
	saidas = null
	nomeArquivoSaida = nomeArquivo
	this.opcoes = opcoes
	
	''
	configurarScripting()

	''
	configurarDB()
	
end sub

''''''''
sub Efd.finalizarExtracao()

	''
	onProgress("Gravando planilha: " + nomeArquivoSaida, 0)
	ew->Flush(onProgress)
	ew->Close
	delete ew
   
	''
	fecharDb()
	if opcoes.dbEmDisco then
		if not opcoes.manterDb then
			kill nomeArquivoSaida + ".db"
		end if
	end if
	
	''
	lua_close( lua )
	
end sub

''''''''
function Efd.processar(nomeArquivo as string) as Boolean
   
	if opcoes.formatoDeSaida <> FT_NULL then
		gerarPlanilhas(nomeArquivo)
	else
		onProgress(null, 1)
	end if
	
	if opcoes.gerarRelatorios then
		if tipoArquivo = TIPO_ARQUIVO_EFD then
			infAssinatura = lerInfoAssinatura(nomeArquivo, assinaturaP7K_DER())
		
			gerarRelatorios(nomeArquivo)
			
			if infAssinatura <> NULL then
				delete infAssinatura
			end if
		end if
	end if
	
	do while regListHead <> null
		var next_ = regListHead->next_
		delete regListHead
		regListHead = next_
	loop

	regListHead = null

	delete sintegraDict
	delete infoComplDict
	delete centroCustoDict
	delete contaContabDict
	delete bemCiapDict
	delete itemIdDict
	delete participanteDict

	function = true
end function

''''''''
private function luacb_efd_plan_get cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	lua_getglobal(L, "efd")
	var g_efd = cast(Efd ptr, lua_touserdata(L, -1))
	lua_pop(L, 1)
	
	if args = 1 then
		var planName = lua_tostring(L, 1)

		var plan = g_efd->getPlanilha(planName)
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

''''''''
static function Efd.luacb_efd_participante_get cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)

	lua_getglobal(L, "efd")
	var g_efd = cast(Efd ptr, lua_touserdata(L, -1))
	lua_pop(L, 1)
	
	if args = 2 then
		var idParticipante = lua_tostring(L, 1)
		var formatar = lua_toboolean(L, 2) <> 0

		var part = cast( TParticipante ptr, g_efd->participanteDict->lookup(idParticipante) )
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
			lua_pushstring(L, g_efd->codMunicipio2Nome(part->municip))
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

private function luacb_efd_onProgress cdecl(L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	lua_getglobal(L, "efd")
	var g_efd = cast(Efd ptr, lua_touserdata(L, -1))
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
	var g_efd = cast(Efd ptr, lua_touserdata(L, -1))
	lua_pop(L, 1)
	
	if args = 1 then
		var msg = cast(zstring ptr, lua_tostring(L, 1))
		g_efd->onError(msg)
	end if
	
	function = 0
end function

''''''''
sub Efd.exportAPI(L as lua_State ptr)
	
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
	lua_register(L, "efd_participante_get", @luacb_efd_participante_get)
	lua_register(L, "efd_rel_addItemAnalitico", @luacb_efd_rel_addItemAnalitico)
	lua_register(L, "onProgress", @luacb_efd_onProgress)
	lua_register(L, "onError", @luacb_efd_onError)
	
end sub
#include once "efd.bi"
#include once "Dict.bi"
#include once "DB.bi"
#include once "trycatch.bi"

''''''''
private function lua_criarTabela(lua as lua_State ptr, db as TDb ptr, tabela as const zstring ptr, onError as OnErrorCB) as TDbStmt ptr

	try
		lua_getglobal(lua, "criarTabela_" + *tabela)
		lua_pushlightuserdata(lua, db)
		lua_call(lua, 1, 1)
		var res = db->prepare(lua_tostring(lua, -1))
		if res = null then
			onError("Erro ao executar script lua de criação de tabela: " + "criarTabela_" + *tabela + ": " + *db->getErrorMsg())
		end if
		function = res
		lua_pop(lua, 1)
	catch
		onError("Erro ao executar script lua de criação de tabela: " + "criarTabela_" + *tabela + ". Verifique erros de sintaxe")
	endtry

end function

''''''''
sub Efd.configurarDB()

	db = new TDb
	if not opcoes.dbEmDisco then
		db->open()
	else
		kill nomeArquivoSaida + ".db"
		db->open(nomeArquivoSaida + ".db")
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
		onError("Erro ao executar script lua de criação de DB. Verifique erros de sintaxe")
	endtry

end sub   

''''''''
sub Efd.fecharDb()
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
function Efd.adicionarDFe(dfe as TDFe ptr, fazerInsert as boolean) as long

	if opcoes.pularResumos andalso opcoes.pularAnalises andalso not opcoes.manterDb then
		return 0
	end if
	
	if chaveDFeDict->lookup(dfe->chave) = null then
		chaveDFeDict->add(dfe->chave, dfe)

		if dfeListHead = null then
			dfeListHead = dfe
		else
			dfeListTail->next_ = dfe
		end if
		
		dfeListTail = dfe
		dfe->next_ = null
	end if

	if fazerInsert then
		'' adicionar ao db
		select case dfe->operacao
		case ENTRADA
			'' (cnpjEmit, ufEmit, serie, numero, modelo, chave, dataEmit, valorOp, ieEmit)
			db_dfeEntradaInsertStmt->reset()
			db_dfeEntradaInsertStmt->bind(1, dfe->cnpjEmit)
			db_dfeEntradaInsertStmt->bind(2, dfe->ufEmit)
			db_dfeEntradaInsertStmt->bind(3, dfe->serie)
			db_dfeEntradaInsertStmt->bind(4, dfe->numero)
			db_dfeEntradaInsertStmt->bind(5, dfe->modelo)
			db_dfeEntradaInsertStmt->bind(6, dfe->chave)
			db_dfeEntradaInsertStmt->bind(7, dfe->dataEmi)
			db_dfeEntradaInsertStmt->bind(8, dfe->valorOperacao)
			if len(dfe->nfe.ieEmit) > 0 then
				db_dfeEntradaInsertStmt->bind(9, dfe->nfe.ieEmit)
			else
				db_dfeEntradaInsertStmt->bindNull(9)
			end if
			
			if not db->execNonQuery(db_dfeEntradaInsertStmt) then
				onError("Erro ao inserir DFe de entrada: " & *db->getErrorMsg())
				return 0
			end if
			
			nroDfe += 1
			return db->lastId()
		
		case SAIDA
			'' (cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp, ieDest)
			db_dfeSaidaInsertStmt->reset()

			db_dfeSaidaInsertStmt->bind(1, dfe->cnpjDest)
			db_dfeSaidaInsertStmt->bind(2, dfe->ufDest)
			db_dfeSaidaInsertStmt->bind(3, dfe->serie)
			db_dfeSaidaInsertStmt->bind(4, dfe->numero)
			db_dfeSaidaInsertStmt->bind(5, dfe->modelo)
			db_dfeSaidaInsertStmt->bind(6, dfe->chave)
			db_dfeSaidaInsertStmt->bind(7, dfe->dataEmi)
			db_dfeSaidaInsertStmt->bind(8, dfe->valorOperacao)
			if len(dfe->nfe.ieDest) > 0 then
				db_dfeSaidaInsertStmt->bind(9, dfe->nfe.ieDest)
			else
				db_dfeSaidaInsertStmt->bindNull(9)
			end if
		
			if not db->execNonQuery(db_dfeSaidaInsertStmt) then
				onError("Erro ao inserir DFe de saída: " & *db->getErrorMsg())
				return 0
			end if

			nroDfe += 1
			return db->lastId()
		end select
		
	end if
	
	return 0

end function

''''''''
function Efd.adicionarItemDFe(chave as const zstring ptr, item as TDFe_NFeItem ptr) as long

	if opcoes.pularResumos andalso opcoes.pularAnalises andalso not opcoes.manterDb then
		return 0
	end if

	'' (serie, numero, modelo, numItem, chave, cfop, valorProd, valorDesc, valorAcess, bc, aliq, icms, bcIcmsST, , aliqST, icmsST, ncm, cst, qtd, unidade, codProduto, descricao) 
	db_itensDfeSaidaInsertStmt->reset()
	db_itensDfeSaidaInsertStmt->bind(1, item->serie)
	db_itensDfeSaidaInsertStmt->bind(2, item->numero)
	db_itensDfeSaidaInsertStmt->bind(3, item->modelo)
	db_itensDfeSaidaInsertStmt->bind(4, item->nroItem)
	db_itensDfeSaidaInsertStmt->bind(5, chave)
	db_itensDfeSaidaInsertStmt->bind(6, item->cfop)
	db_itensDfeSaidaInsertStmt->bind(7, item->valorProduto)
	db_itensDfeSaidaInsertStmt->bind(8, item->desconto)
	db_itensDfeSaidaInsertStmt->bind(9, item->despesasAcess)
	db_itensDfeSaidaInsertStmt->bind(10, item->bcICMS)
	db_itensDfeSaidaInsertStmt->bind(11, item->aliqICMS)
	db_itensDfeSaidaInsertStmt->bind(12, item->icms)
	db_itensDfeSaidaInsertStmt->bind(13, item->bcIcmsST)
	db_itensDfeSaidaInsertStmt->bind(14, item->aliqIcmsST)
	db_itensDfeSaidaInsertStmt->bind(15, item->icmsST)
	db_itensDfeSaidaInsertStmt->bind(16, item->ncm)
	db_itensDfeSaidaInsertStmt->bind(17, item->cst)
	db_itensDfeSaidaInsertStmt->bind(18, item->qtd)
	if opcoes.manterDb then
		db_itensDfeSaidaInsertStmt->bind(19, item->unidade)
		db_itensDfeSaidaInsertStmt->bind(20, item->codProduto)
		db_itensDfeSaidaInsertStmt->bind(21, item->descricao)
	else
		db_itensDfeSaidaInsertStmt->bind(19, null)
		db_itensDfeSaidaInsertStmt->bind(20, null)
		db_itensDfeSaidaInsertStmt->bind(21, null)
	end if

	if not db->execNonQuery(db_itensDfeSaidaInsertStmt) then
		onError("Erro ao inserir Item DFe de entrada: " & *db->getErrorMsg())
		return 0
	end if
	
	return db->lastId()
end function
   

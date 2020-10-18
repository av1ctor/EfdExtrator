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
sub Efd.adicionarMestre(reg as TMestre ptr)

	'' (versao, original, dataIni, dataFim, nome, cnpj, uf, ie)
	db_mestreInsertStmt->reset()
	db_mestreInsertStmt->bind(1, reg->versaoLayout)
	db_mestreInsertStmt->bind(2, cint(reg->original))
	db_mestreInsertStmt->bind(3, reg->dataIni)
	db_mestreInsertStmt->bind(4, reg->dataFim)
	db_mestreInsertStmt->bind(5, reg->nome)
	db_mestreInsertStmt->bind(6, reg->cnpj)
	db_mestreInsertStmt->bind(7, reg->uf)
	db_mestreInsertStmt->bind(8, reg->ie)
	
	if not db->execNonQuery(db_mestreInsertStmt) then
		onError("Erro ao inserir registro na EFD_Mestre: " & *db->getErrorMsg())
	end if

end sub

''''''''
sub Efd.adicionarDocEscriturado(doc as TDocDF ptr)
	
	select case as const doc->situacao
	case REGULAR, EXTEMPORANEO
		var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipante) )
		
		var uf = iif(part->municip >= 1100000 and part->municip <= 5399999, part->municip \ 100000, 99)
		
		'' adicionar ao db
		if doc->operacao = ENTRADA then
			'' (periodo, cnpjEmit, ufEmit, serie, numero, modelo, chave, dataEmit, valorOp, IE)
			db_LREInsertStmt->reset()
			db_LREInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
			if len(part->cpf) > 0 then 
				db_LREInsertStmt->bind(2, part->cpf)
			else
				db_LREInsertStmt->bind(2, part->cnpj)
			end if
			db_LREInsertStmt->bind(3, uf)
			db_LREInsertStmt->bind(4, doc->serie)
			db_LREInsertStmt->bind(5, doc->numero)
			db_LREInsertStmt->bind(6, doc->modelo)
			db_LREInsertStmt->bind(7, doc->chave)
			db_LREInsertStmt->bind(8, doc->dataEmi)
			db_LREInsertStmt->bind(9, doc->valorTotal)
			if len(part->ie) > 0 then
				db_LREInsertStmt->bind(10, trim(part->ie))
			else
				db_LREInsertStmt->bindNull(10)
			end if
			
			if not db->execNonQuery(db_LREInsertStmt) then
				onError("Erro ao inserir registro na EFD_LRE: " & *db->getErrorMsg())
			end if
		else
			'' (periodo, cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp, IE)
			db_LRSInsertStmt->reset()
			db_LRSInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
			if len(part->cpf) > 0 then 
				db_LRSInsertStmt->bind(2, part->cpf)
			else
				db_LRSInsertStmt->bind(2, part->cnpj)
			end if
			db_LRSInsertStmt->bind(3, uf)
			db_LRSInsertStmt->bind(4, doc->serie)
			db_LRSInsertStmt->bind(5, doc->numero)
			db_LRSInsertStmt->bind(6, doc->modelo)
			db_LRSInsertStmt->bind(7, doc->chave)
			db_LRSInsertStmt->bind(8, doc->dataEmi)
			db_LRSInsertStmt->bind(9, doc->valorTotal)
			if len(part->ie) > 0 then
				db_LRSInsertStmt->bind(10, trim(part->ie))
			else
				db_LRSInsertStmt->bindNull(10)
			end if
		
			if not db->execNonQuery(db_LRSInsertStmt) then
				onError("Erro ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
			end if
		end if
	
	case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
		'' !!!TODO!!! inserir em outra tabela para fazermos análises posteriores
	
	case else
		'' !!!TODO!!! como tratar outras situações? os dados vêm completos?
	end select
	
end sub

''''''''
sub Efd.adicionarDocEscriturado(doc as TDocECF ptr)
	
	select case as const doc->situacao
	case REGULAR, EXTEMPORANEO
	
		'' só existe de saída para ECF
		if doc->operacao = SAIDA then
			'' (periodo, cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp)
			db_LRSInsertStmt->reset()
			db_LRSInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
			db_LRSInsertStmt->bind(2, doc->cpfCnpjAdquirente)
			db_LRSInsertStmt->bind(3, 35)
			db_LRSInsertStmt->bind(4, 0)
			db_LRSInsertStmt->bind(5, doc->numero)
			db_LRSInsertStmt->bind(6, doc->modelo)
			db_LRSInsertStmt->bind(7, doc->chave)
			db_LRSInsertStmt->bind(8, doc->dataEmi)
			db_LRSInsertStmt->bind(9, doc->valorTotal)
			db_dfeSaidaInsertStmt->bindNull(10)
		
			if not db->execNonQuery(db_LRSInsertStmt) then
				onError("Erro ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
			end if
		end if
	
	case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
		'' !!!TODO!!! inserir em outra tabela para fazermos análises posteriores
	
	case else
		'' !!!TODO!!! como tratar outras situações? os dados vêm completos?
	end select
	
end sub

''''''''
sub Efd.adicionarDocEscriturado(doc as TDocSAT ptr)
	
	select case as const doc->situacao
	case REGULAR, EXTEMPORANEO
	
		'' só existe de saída para SAT
		if doc->operacao = SAIDA then
			'' (periodo, cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp)
			db_LRSInsertStmt->reset()
			db_LRSInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
			db_LRSInsertStmt->bind(2, 0) '' não é possível usar doc->cpfCnpjAdquirente, porque relatório do BO vem sem essa info
			db_LRSInsertStmt->bind(3, 35)
			db_LRSInsertStmt->bind(4, 0)
			db_LRSInsertStmt->bind(5, doc->numero)
			db_LRSInsertStmt->bind(6, doc->modelo)
			db_LRSInsertStmt->bind(7, doc->chave)
			db_LRSInsertStmt->bind(8, doc->dataEmi)
			db_LRSInsertStmt->bind(9, doc->valorTotal)
			db_dfeSaidaInsertStmt->bindNull(10)
		
			if not db->execNonQuery(db_LRSInsertStmt) then
				onError("Erro ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
			end if
		end if
	
	case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
		'' !!!TODO!!! inserir em outra tabela para fazermos análises posteriores
	
	case else
		'' !!!TODO!!! como tratar outras situações? os dados vêm completos?
	end select
	
end sub

''''''''
sub Efd.adicionarItemNFEscriturado(item as TDocNFItem ptr)
	
	var doc = item->documentoPai
	select case as const doc->situacao
	case REGULAR, EXTEMPORANEO
		var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipante) )
		
		var uf = iif(part->municip >= 1100000 and part->municip <= 5399999, part->municip \ 100000, 99)

		'' (periodo, cnpjEmit, ufEmit, serie, numero, modelo, numItem, cst, cst_origem, cst_tribut, cfop, qtd, valorProd, valorDesc, bc, aliq, icms, bcIcmsST, aliqIcmsST, icmsST, itemId)
		db_itensNfLRInsertStmt->reset()
		db_itensNfLRInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
		db_itensNfLRInsertStmt->bind(2, iif(len(part->cpf) > 0, part->cpf, part->cnpj))
		db_itensNfLRInsertStmt->bind(3, uf)
		db_itensNfLRInsertStmt->bind(4, doc->serie)
		db_itensNfLRInsertStmt->bind(5, doc->numero)
		db_itensNfLRInsertStmt->bind(6, doc->modelo)
		db_itensNfLRInsertStmt->bind(7, item->numItem)
		db_itensNfLRInsertStmt->bind(8, item->cstIcms)
		db_itensNfLRInsertStmt->bind(9, item->cstIcms \ 100)
		db_itensNfLRInsertStmt->bind(10, item->cstIcms mod 100)
		db_itensNfLRInsertStmt->bind(11, item->cfop)
		db_itensNfLRInsertStmt->bind(12, item->qtd)
		db_itensNfLRInsertStmt->bind(13, item->valor)
		db_itensNfLRInsertStmt->bind(14, item->desconto)
		db_itensNfLRInsertStmt->bind(15, item->bcICMS)
		db_itensNfLRInsertStmt->bind(16, item->aliqICMS)
		db_itensNfLRInsertStmt->bind(17, item->icms)
		db_itensNfLRInsertStmt->bind(18, item->bcICMSST)
		db_itensNfLRInsertStmt->bind(19, item->aliqICMSST)
		db_itensNfLRInsertStmt->bind(20, item->icmsST)
		if opcoes.manterDb then
			db_itensNfLRInsertStmt->bind(21, item->itemId)
		else
			db_itensNfLRInsertStmt->bind(21, null)
		end if
		
		if not db->execNonQuery(db_itensNfLRInsertStmt) then
			onError("Erro ao inserir registro na EFD_Itens: " & *db->getErrorMsg())
		end if
	end select
	
end sub

''''''''
sub Efd.adicionarRessarcStEscriturado(doc as TDocNFItemRessarcSt ptr)

	var docPai = doc->documentoPai
	var docAvo = doc->documentoPai->documentoPai
	
	var part = cast( TParticipante ptr, participanteDict->lookup(docAvo->idParticipante) )
	var uf = iif(part->municip >= 1100000 and part->municip <= 5399999, part->municip \ 100000, 99)
	
	var partUlt = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipanteUlt) )
	var ufUlt = iif(partUlt->municip >= 1100000 and partUlt->municip <= 5399999, partUlt->municip \ 100000, 99)
	
	'' (periodo, cnpjEmit, ufEmit, serie, numero, modelo, nroItem, cnpjUlt, ufUlt, serieUlt, numeroUlt, modeloUlt, chaveUlt, dataUlt, valorUlt, bcSTUlt, qtdUlt, nroItemUlt)
	db_ressarcStItensNfLRSInsertStmt->reset()
	db_ressarcStItensNfLRSInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
	db_ressarcStItensNfLRSInsertStmt->bind(2, iif(len(part->cpf) > 0, part->cpf, part->cnpj))
	db_ressarcStItensNfLRSInsertStmt->bind(3, uf)
	db_ressarcStItensNfLRSInsertStmt->bind(4, docAvo->serie)
	db_ressarcStItensNfLRSInsertStmt->bind(5, docAvo->numero)
	db_ressarcStItensNfLRSInsertStmt->bind(6, docAvo->modelo)
	db_ressarcStItensNfLRSInsertStmt->bind(7, docPai->numItem)
	db_ressarcStItensNfLRSInsertStmt->bind(8, partUlt->cnpj)
	db_ressarcStItensNfLRSInsertStmt->bind(9, ufUlt)
	db_ressarcStItensNfLRSInsertStmt->bind(10, doc->serieUlt)
	db_ressarcStItensNfLRSInsertStmt->bind(11, doc->numeroUlt)
	db_ressarcStItensNfLRSInsertStmt->bind(12, doc->modeloUlt)
	if len(doc->chaveNFeUlt) > 0 then
		db_ressarcStItensNfLRSInsertStmt->bind(13, doc->chaveNFeUlt)
	else
		db_ressarcStItensNfLRSInsertStmt->bindNull(13)
	end if
	db_ressarcStItensNfLRSInsertStmt->bind(14, doc->dataUlt)
	db_ressarcStItensNfLRSInsertStmt->bind(15, doc->valorUlt)
	db_ressarcStItensNfLRSInsertStmt->bind(16, doc->valorBcST)
	db_ressarcStItensNfLRSInsertStmt->bind(17, doc->qtdUlt)
	if doc->numItemNFeUlt > 0 then
		db_ressarcStItensNfLRSInsertStmt->bind(18, doc->numItemNFeUlt)
	else
		db_ressarcStItensNfLRSInsertStmt->bindNull(18)
	end if

	if not db->execNonQuery(db_ressarcStItensNfLRSInsertStmt) then
		onError("Erro ao inserir registro na EFD_Ressarc_Itens: " & *db->getErrorMsg())
	end if
	
end sub

''''''''
sub Efd.adicionarItemEscriturado(item as TItemId ptr)

	'' (id, descricao, ncm, cest, aliqInt)
	db_itensIdInsertStmt->reset()
	db_itensIdInsertStmt->bind(1, item->id)
	db_itensIdInsertStmt->bind(2, item->descricao)
	db_itensIdInsertStmt->bind(3, item->ncm)
	db_itensIdInsertStmt->bind(4, item->CEST)
	db_itensIdInsertStmt->bind(5, item->aliqICMSInt)
	
	if not db->execNonQuery(db_itensIdInsertStmt) then
		onError("Erro ao inserir registro na EFD_ItensId: " & *db->getErrorMsg())
	end if

end sub

''''''''
sub Efd.adicionarAnalEscriturado(anal as TDocItemAnal ptr)

	var doc = @anal->documentoPai->nf
	var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipante) )
	
	var uf = iif(part->municip >= 1100000 and part->municip <= 5399999, part->municip \ 100000, 99)

	'' (operacao, periodo, cnpj, uf, serie, numero, modelo, numReg, cst, cst_origem, cst_tribut, cfop, aliq, valorOp, bc, icms, bcIcmsST, icmsST, redBC, ipi)
	db_analInsertStmt->reset()
	db_analInsertStmt->bind(1, doc->operacao)
	db_analInsertStmt->bind(2, valint(regMestre->mestre.dataIni))
	db_analInsertStmt->bind(3, iif(len(part->cpf) > 0, part->cpf, part->cnpj))
	db_analInsertStmt->bind(4, uf)
	db_analInsertStmt->bind(5, doc->serie)
	db_analInsertStmt->bind(6, doc->numero)
	db_analInsertStmt->bind(7, doc->modelo)
	db_analInsertStmt->bind(8, anal->num)
	db_analInsertStmt->bind(9, anal->cst)
	db_analInsertStmt->bind(10, anal->cst \ 100)
	db_analInsertStmt->bind(11, anal->cst mod 100)
	db_analInsertStmt->bind(12, anal->cfop)
	db_analInsertStmt->bind(13, anal->aliq)
	db_analInsertStmt->bind(14, anal->valorOp)
	db_analInsertStmt->bind(15, anal->bc)
	db_analInsertStmt->bind(16, anal->ICMS)
	db_analInsertStmt->bind(17, anal->bcST)
	db_analInsertStmt->bind(18, anal->ICMSST)
	db_analInsertStmt->bind(19, anal->redBC)
	db_analInsertStmt->bind(20, anal->IPI)
	
	if not db->execNonQuery(db_analInsertStmt) then
		onError("Erro ao inserir registro na EDF_Anal: " & *db->getErrorMsg())
	end if

end sub

''''''''
sub Efd.addRegistroAoDB(reg as TRegistro ptr)

	select case as const reg->tipo
	case DOC_NF
		adicionarDocEscriturado(@reg->nf)
	case DOC_NF_ITEM
		adicionarItemNFEscriturado(@reg->itemNF)
	case DOC_NF_ANAL
		adicionarAnalEscriturado(@reg->itemAnal)
	case DOC_CT
		adicionarDocEscriturado(@reg->ct)
	case DOC_ECF
		adicionarDocEscriturado(@reg->ecf)
	case DOC_SAT
		adicionarDocEscriturado(@reg->sat)
	case DOC_NF_ITEM_RESSARC_ST
		adicionarRessarcStEscriturado(@reg->itemRessarcSt)
	case ITEM_ID
		if opcoes.manterDb then
			adicionarItemEscriturado(@reg->itemId)
		end if
	case MESTRE
		adicionarMestre(@reg->mestre)
	end select
	
end sub

''''''''
sub Efd.adicionarDFe(dfe as TDFe ptr, fazerInsert as boolean)
	
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
			end if
		
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
			end if
		end select
		
		nroDfe += 1
	end if

end sub

''''''''
sub Efd.adicionarItemDFe(chave as const zstring ptr, item as TDFe_NFeItem ptr)
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
		end if
end sub
   

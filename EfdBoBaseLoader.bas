#include once "Efd.bi"
#include once "EfdBoBaseLoader.bi"

''''''''
constructor EfdBoBaseLoader(ctx as EfdBoLoaderContext ptr, opcoes as OpcoesExtracao ptr)
	this.ctx = ctx
	this.opcoes = opcoes
end constructor

''''''''
function EfdBoBaseLoader.withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdBoBaseLoader ptr
	this.onProgress = onProgress
	this.onError = onError
	return @this
end function

''''''''
function EfdBoBaseLoader.withDBs(db as TDb ptr) as EfdBoBaseLoader ptr
	this.db = db
	return @this
end function

''''''''
function EfdBoBaseLoader.withStmts( _
	dfeEntradaInsertStmt as TDbStmt ptr, _
	dfeSaidaInsertStmt as TDbStmt ptr, _
	itensDfeSaidaInsertStmt as TDbStmt ptr _
	) as EfdBoBaseLoader ptr
	
	this.db_dfeEntradaInsertStmt = dfeEntradaInsertStmt
	this.db_dfeSaidaInsertStmt = dfeSaidaInsertStmt
	this.db_itensDfeSaidaInsertStmt = itensDfeSaidaInsertStmt
	
	return @this
end function

''''''''
destructor EfdBoBaseLoader()
end destructor

''''''''
function EfdBoBaseLoader.adicionarDFe(dfe as TDFe_NFe ptr, fazerInsert as boolean) as long
	return adicionarDFe(dfe, true, fazerInsert)
end function

''''''''
function EfdBoBaseLoader.adicionarDFe(dfe as TDFe_CTe ptr, fazerInsert as boolean) as long
	return adicionarDFe(dfe, false, fazerInsert)
end function

''''''''
function EfdBoBaseLoader.adicionarDFe(dfe as TDFe ptr, isNfe as boolean, fazerInsert as boolean) as long

	if opcoes->pularResumos andalso opcoes->pularAnalises andalso not opcoes->manterDb then
		return 0
	end if
	
	if ctx->chaveDFeDict->lookup(dfe->chave) = null then
		ctx->chaveDFeDict->add(dfe->chave, dfe)

		if ctx->dfeListHead = null then
			ctx->dfeListHead = dfe
		else
			ctx->dfeListTail->prox = dfe
		end if
		
		ctx->dfeListTail = dfe
		dfe->prox = null
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
			if isNfe andalso len(cast(TDFe_NFe ptr, dfe)->ieEmit) > 0 then
				db_dfeEntradaInsertStmt->bind(9, cast(TDFe_NFe ptr, dfe)->ieEmit)
			else
				db_dfeEntradaInsertStmt->bindNull(9)
			end if
			
			if not db->execNonQuery(db_dfeEntradaInsertStmt) then
				onError("Erro ao inserir DFe de entrada: " & *db->getErrorMsg())
				return 0
			end if
			
			ctx->nroDfe += 1
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
			if isNfe andalso len(cast(TDFe_NFe ptr, dfe)->ieDest) > 0 then
				db_dfeSaidaInsertStmt->bind(9, cast(TDFe_NFe ptr, dfe)->ieDest)
			else
				db_dfeSaidaInsertStmt->bindNull(9)
			end if
		
			if not db->execNonQuery(db_dfeSaidaInsertStmt) then
				onError("Erro ao inserir DFe de saída: " & *db->getErrorMsg())
				return 0
			end if

			ctx->nroDfe += 1
			return db->lastId()
		end select
		
	end if
	
	return 0

end function

''''''''
function EfdBoBaseLoader.adicionarItemDFe(chave as const zstring ptr, item as TDFe_NFeItem ptr) as long

	if opcoes->pularResumos andalso opcoes->pularAnalises andalso not opcoes->manterDb then
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
	if opcoes->manterDb then
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
   

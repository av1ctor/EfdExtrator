#include once "EfdTabelaExport.bi"
#include once "vbcompat.bi"
#include once "trycatch.bi"

''''''''
constructor EfdTabelaExport(nomeArquivo as String, opcoes as OpcoesExtracao ptr)
	this.nomeArquivo = nomeArquivo
	this.opcoes = opcoes
	
	ew = new TableWriter
	ew->create(nomeArquivo, opcoes->formatoDeSaida)

	entradas = null
	saidas = null
end constructor

''''''''
destructor EfdTabelaExport()
	if ew <> null then
		delete ew
	end if
end destructor

''''''''
function EfdTabelaExport.withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdTabelaExport ptr
	this.onProgress = onProgress
	this.onError = onError
	return @this
end function

''''''''
function EfdTabelaExport.withLua(lua as lua_State ptr, customLuaCbDict as TDict ptr) as EfdTabelaExport ptr
	this.lua = lua
	this.customLuaCbDict = customLuaCbDict
	return @this
end function

''''''''
function EfdTabelaExport.withState(itemNFeSafiFornecido as boolean) as EfdTabelaExport ptr
	this.itemNFeSafiFornecido = itemNFeSafiFornecido
	return @this
end function

''''''''
function EfdTabelaExport.withFiltros( _
		filtrarPorCnpj as OnFilterByStrCB, _
		filtrarPorChave as OnFilterByStrCB _
	) as EfdTabelaExport ptr
	this.filtrarPorCnpj = filtrarPorCnpj
	this.filtrarPorChave = filtrarPorChave
	return @this
end function

''''''''
function EfdTabelaExport.withDicionarios( _
		participanteDict as TDict ptr, _
		itemIdDict as TDict ptr, _
		chaveDFeDict as TDict ptr, _
		infoComplDict as TDict ptr, _
		obsLancamentoDict as TDict ptr, _
		bemCiapDict as TDict ptr _
	) as EfdTabelaExport ptr
	this.participanteDict = participanteDict
	this.itemIdDict = itemIdDict
	this.chaveDFeDict = chaveDFeDict
	this.infoComplDict = infoComplDict
	this.obsLancamentoDict = obsLancamentoDict
	this.bemCiapDict = bemCiapDict
	return @this
end function

''''''''
function EfdTabelaExport.getPlanilha(nome as const zstring ptr) as TableTable ptr
		select case lcase(*nome)
		case "entradas"
			return entradas
		case "saidas"
			return saidas
		case "inconsistencias lre"
			return inconsistenciasLRE
		case "inconsistencias lrs"
			return inconsistenciasLRS
		case "resumos lre"
			return resumosLRE
		case "resumos lrs"
			return resumosLRS
		case "ciap"
			return ciap
		case "estoque"
			return estoque
		case "producao"
			return producao
		case "inventario"
			return inventario
		case else
			return null
		end select
end function

''''''''
private sub adicionarColunasComuns(sheet as TableTable ptr, ehEntrada as Boolean)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING, 4)
	sheet->addColumn(CT_STRING, 30)
	sheet->addColumn(CT_STRING, 4)
	sheet->addColumn(CT_STRING, 6)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_STRING, 45)
	sheet->addColumn(CT_STRING, 6)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_NUMBER)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_NUMBER)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_NUMBER)
	sheet->addColumn(CT_STRING, 4)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING, 30)
   
	var row = sheet->addRow(true)
	row->addCell("CNPJ " + iif(ehEntrada, "Emitente", "Destinatario"))
	row->addCell("IE " + iif(ehEntrada, "Emitente", "Destinatario"))
	row->addCell("UF " + iif(ehEntrada, "Emitente", "Destinatario"))
	row->addCell("Razao Social " + iif(ehEntrada, "Emitente", "Destinatario"))
	row->addCell("Modelo")
	row->addCell("Serie")
	row->addCell("Numero")
	row->addCell("Data Emissao")
	row->addCell("Data " + iif(ehEntrada, "Entrada", "Saida"))
	row->addCell("Chave")
	row->addCell("Situacao")
	row->addCell("BC ICMS")
	row->addCell("Aliq ICMS")
	row->addCell("Valor ICMS")
	row->addCell("BC ICMS ST")
	row->addCell("Aliq ICMS ST")
	row->addCell("Valor ICMS ST")
	row->addCell("Valor IPI")
	row->addCell("Valor Item")
	row->addCell("Nro Item")
	row->addCell("Qtd")
	row->addCell("Unidade")
	row->addCell("CFOP")
	row->addCell("CST")
	row->addCell("NCM")
	row->addCell("Codigo Item")
	row->addCell("Descricao Item")
	
	if not ehEntrada then
		sheet->addColumn(CT_MONEY)
		row->addCell("DifAl FCP")
		sheet->addColumn(CT_MONEY)
		row->addCell("DifAl ICMS Orig")
		sheet->addColumn(CT_MONEY)
		row->addCell("DifAl ICMS Dest")
	end if
	
	sheet->addColumn(CT_STRING, 40)
	row->addCell("Info. complementares")

	sheet->addColumn(CT_STRING, 40)
	row->addCell("Obs. lancamento")
end sub

private sub criarColunasApuracaoIcms(sheet as TableTable ptr)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	for i as integer = 1 to MAX_AJUSTES
		sheet->addColumn(CT_STRING, 80)
	next
	
	var row = sheet->addRow(true)
	row->addCell("Inicio")
	row->addCell("Fim")
	row->addCell("Total Debitos")
	row->addCell("Ajustes Debitos")
	row->addCell("Total Ajuste Deb")
	row->addCell("Estornos Credito")
	row->addCell("Total Creditos")
	row->addCell("Ajustes Creditos")
	row->addCell("Total Ajuste Cred")
	row->addCell("Estornos Debito")
	row->addCell("Saldo Cred Anterior")
	row->addCell("Saldo Devedor Apurado")
	row->addCell("Total Deducoes")
	row->addCell("ICMS a Recolher")
	row->addCell("Saldo Credor a Transportar")
	row->addCell("Deb Extra Apuracao")
	for i as integer = 1 to MAX_AJUSTES
		row->addCell("Detalhe Ajuste " & i)
	next
	
end sub

private sub criarColunasApuracaoIcmsST(sheet as TableTable ptr)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_STRING, 4)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	for i as integer = 1 to MAX_AJUSTES
		sheet->addColumn(CT_STRING, 80)
	next

	var row = sheet->addRow(true)
	row->addCell("Inicio")
	row->addCell("Fim")
	row->addCell("UF")
	row->addCell("Movimentacao")
	row->addCell("Saldo Credor Anterior")
	row->addCell("Total Devolucao Merc")
	row->addCell("Total Ressarcimentos")
	row->addCell("Total Ajustes Cred")
	row->addCell("Total Ajustes Cred Docs")
	row->addCell("Total Retencao")
	row->addCell("Total Ajustes Deb")
	row->addCell("Total Ajustes Deb Docs")
	row->addCell("Saldo Devedor ant. Deducoes")
	row->addCell("Total Deducoes")
	row->addCell("ICMS a Recolher")
	row->addCell("Saldo Credor a Transportar")
	row->addCell("Deb Extra Apuracao")
	for i as integer = 1 to MAX_AJUSTES
		row->addCell("Detalhe Ajuste " & i)
	next
end sub

private sub criarColunasInventario(sheet as TableTable ptr)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING, 30)
	sheet->addColumn(CT_STRING, 6)
	sheet->addColumn(CT_NUMBER)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_MONEY)

	var row = sheet->addRow(true)
	row->addCell("Data Inventario")
	row->addCell("Codigo")
	row->addCell("NCM")
	row->addCell("Tipo")
	row->addCell("Tipo (Descricao)")
	row->addCell("Descricao")
	row->addCell("Unidade")
	row->addCell("Qtd")
	row->addCell("Valor Unitario")
	row->addCell("Valor Item")
	row->addCell("Ind. Propriedade")
	row->addCell("CNPJ Proprietario")
	row->addCell("Texto Complementar")
	row->addCell("Codigo Conta Contabil")
	row->addCell("Valor Item IR")
end sub

private sub criarColunasCIAP(sheet as TableTable ptr)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_NUMBER)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_STRING, 6)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_STRING, 4)
	sheet->addColumn(CT_STRING, 6)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_STRING, 30)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING, 4)
	sheet->addColumn(CT_STRING, 30)

	var row = sheet->addRow(true)
	row->addCell("Data Inicial")
	row->addCell("Data Final")
	row->addCell("Soma Total Saidas Tributadas")
	row->addCell("Soma Total Saidas")
	row->addCell("Indice")
	row->addCell("Codigo Bem")
	row->addCell("Descricao Bem")
	row->addCell("Data Movimentacao")
	row->addCell("Tipo Movimentacao")
	row->addCell("Valor ICMS")
	row->addCell("Valor ICMS ST")
	row->addCell("Valor ICMS Frete")
	row->addCell("Valor ICMS Difal")
	row->addCell("Num. Parcela")
	row->addCell("Valor Parcela")
	row->addCell("Modelo")
	row->addCell("Serie")
	row->addCell("Numero")
	row->addCell("Data Emissao")
	row->addCell("Chave NF-e")
	row->addCell("CNPJ")
	row->addCell("IE")
	row->addCell("UF")
	row->addCell("Razao Social")
	end sub

private sub criarColunasEstoque(sheet as TableTable ptr)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING, 30)
	sheet->addColumn(CT_NUMBER)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING, 4)
	sheet->addColumn(CT_STRING, 30)

	var row = sheet->addRow(true)
	row->addCell("Data Inicial")
	row->addCell("Data Final")
	row->addCell("Codigo Item")
	row->addCell("NCM Item")
	row->addCell("Tipo Item")
	row->addCell("Tipo Item (Descricao)")
	row->addCell("Descricao Item")
	row->addCell("Qtd")
	row->addCell("Tipo")
	row->addCell("Prop CNPJ")
	row->addCell("Prop IE")
	row->addCell("Prop UF")
	row->addCell("Prop Razao Social")
end sub

private sub criarColunasProducao(sheet as TableTable ptr)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING, 30)
	sheet->addColumn(CT_NUMBER)
	sheet->addColumn(CT_STRING)

	var row = sheet->addRow(true)
	row->addCell("Data Inicial")
	row->addCell("Data Final")
	row->addCell("Codigo Item")
	row->addCell("NCM Item")
	row->addCell("Tipo Item")
	row->addCell("Tipo Item (Descricao)")
	row->addCell("Descricao Item")
	row->addCell("Qtd")
	row->addCell("Codigo Ordem")
end sub

private sub criarColunasRessarcST(sheet as TableTable ptr)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING, 4)
	sheet->addColumn(CT_STRING, 30)
	sheet->addColumn(CT_STRING, 4)
	sheet->addColumn(CT_STRING, 6)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_DATE)
	sheet->addColumn(CT_NUMBER)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_STRING, 45)
	sheet->addColumn(CT_INTNUMBER)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_NUMBER)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_NUMBER)
	sheet->addColumn(CT_MONEY)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING)
	sheet->addColumn(CT_STRING, 45)
	sheet->addColumn(CT_INTNUMBER)

	var row = sheet->addRow(true)
	row->addCell("CNPJ Emitente Ult NF-e Ent")
	row->addCell("IE Emitente Ult NF-e Ent")
	row->addCell("UF Emitente Ult NF-e Ent")
	row->addCell("Razao Social Emitente Ult NF-e Ent")
	row->addCell("Modelo Ult NF-e Ent")
	row->addCell("Serie Ult NF-e Ent")
	row->addCell("Numero Ult NF-e Ent")
	row->addCell("Data Emissao Ult NF-e Ent")
	row->addCell("Qtd Ult Ent")
	row->addCell("Valor Ult Ent")
	row->addCell("BC ICMS ST")
	row->addCell("Chave Ult NF-e Ent")
	row->addCell("Num Item Ult NF-e Ent")
	row->addCell("BC ICMS")
	row->addCell("Aliq ICMS")
	row->addCell("Lim BC ICMS")
	row->addCell("ICMS")
	row->addCell("Aliq ICMS ST")
	row->addCell("Ressarcimento")
	row->addCell("Responsavel")
	row->addCell("Motivo")
	row->addCell("Tipo Doc Arrecad")
	row->addCell("Num Doc Arrecad")
	row->addCell("Chave NF-e Saida")
	row->addCell("Num Item NF-e Saida")
end sub

''''''''
sub EfdTabelaExport.criarPlanilhas()
	'' planilha de entradas
	entradas = ew->addTable("Entradas")
	adicionarColunasComuns(entradas, true)

	'' planilha de saídas
	saidas = ew->addTable("Saidas")
	adicionarColunasComuns(saidas, false)

	'' apuração do ICMS
	apuracaoIcms = ew->addTable("Apuracao ICMS")
	criarColunasApuracaoIcms(apuracaoIcms)
   
	'' apuração do ICMS ST
	apuracaoIcmsST = ew->addTable("Apuracao ICMS ST")
	criarColunasApuracaoIcmsST(apuracaoIcmsST)
	
	'' Inventário
	inventario = ew->addTable("Inventario")
	criarColunasInventario(inventario)

	'' CIAP
	ciap = ew->addTable("CIAP")
	criarColunasCIAP(ciap)

	'' Estoque
	estoque = ew->addTable("Estoque")
	criarColunasEstoque(estoque)

	'' Producao
	producao = ew->addTable("Producao")
	criarColunasProducao(producao)

	'' Ressarcimento ST
	ressarcST = ew->addTable("Ressarcimento ST")
	criarColunasRessarcST(ressarcST)
	
	'' Inconsistencias LRE
	inconsistenciasLRE = ew->addTable("Inconsistencias LRE")

	'' Inconsistencias LRS
	inconsistenciasLRS = ew->addTable("Inconsistencias LRS")
	
	'' Resumos LRE
	resumosLRE = ew->addTable("Resumos LRE")

	'' Resumos LRS
	resumosLRS = ew->addTable("Resumos LRS")
	
	''
	lua_getglobal(lua, "criarPlanilhas")
	lua_call(lua, 0, 0)

	lua_setarGlobal(lua, "efd_plan_entradas", entradas)
	lua_setarGlobal(lua, "efd_plan_saidas", saidas)
	
end sub

function EfdTabelaExport.getInfoCompl(info as TDocInfoCompl ptr) as string
	var res = ""
	
	do while info <> null
		var compl = cast( TInfoCompl ptr, infoComplDict->lookup(info->idCompl))
		res += iif(len(res) > 0, ",", "")
		res += "{'descricao':'" + compl->descricao + "'"
		if len(info->extra) > 0 then 
			res += ", 'extra':'" + info->extra + "'"
		end if
		res += "}"
		info = info->next_
	loop
	
	function = res
end function

function EfdTabelaExport.getObsLanc(obs as TDocObs ptr) as string
	var res = ""
	
	do while obs <> null
		var lanc = cast( TObsLancamento ptr, obsLancamentoDict->lookup(obs->idLanc))
		res += iif(len(res) > 0, ",", "")
		res += "{'descricao':'" + lanc->descricao + "'"
		if len(obs->extra) > 0 then 
			res += ", 'extra':'" + obs->extra + "'"
		end if
		var ajuste = obs->ajusteListHead
		if ajuste <> null then
			res += ", 'ajustes':["
			var cnt = 0
			do 
				res += iif(cnt > 0, ",", "")
				res += "{'codigo':'" + ajuste->idAjuste + "'"
				if len(ajuste->extra) > 0 then 
					res += ", 'extra':'" + ajuste->extra + "'"
				end if
				if len(ajuste->idItem) > 0 then 
					res += ", 'item':'" + ajuste->idItem + "'"
				end if
				res += ", 'bc':'" + DBL2MONEYBR(ajuste->bcICMS) + "'"
				res += ", 'aliq':'" + DBL2MONEYBR(ajuste->aliqICMS) + "'"
				res += ", 'valor':'" + DBL2MONEYBR(ajuste->icms) + "'"
				res += ", 'outros':'" + DBL2MONEYBR(ajuste->outros) + "'"
				res += "}"
				cnt += 1
				ajuste = ajuste->next_
			loop while ajuste <> null
			res += "]"
		end if
		res += "}"
		obs = obs->next_
	loop
	
	function = res
end function

''''''''
sub EfdTabelaExport.gerar(regListHead as TRegistro ptr, regMestre as TMestre ptr, nroRegs as integer)
	
	if entradas = null then
		criarPlanilhas()
	end if
	
	onProgress(!"\tGerando planilhas", 0)
	
	dim as TRegistro ptr reg = null
	try
		var regCnt = 0
		reg = regListHead
		do while reg <> null
			'para cada registro..
			select case as const reg->tipo
			'item de NF-e?
			case DOC_NF_ITEM
				var item = cast(TDocNFItem ptr, reg)
				var doc = item->documentoPai
				var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipante) )

				var emitirLinha = iif(doc->operacao = SAIDA, not opcoes->pularLrs, not opcoes->pularLre)
				if opcoes->filtrarCnpj andalso emitirLinha then
					if part <> null then
						emitirLinha = filtrarPorCnpj(part->cnpj, opcoes->listaCnpj())
					end if
				end if
				
				if opcoes->filtrarChaves andalso emitirLinha then
					emitirLinha = filtrarPorChave(doc->chave, opcoes->listaChaves())
				end if
				
				if opcoes->somenteRessarcimentoST andalso emitirLinha then
					emitirLinha = item->itemRessarcStListHead <> null
				end if
				
				if emitirLinha then
					'só existe item para entradas (exceto quando há ressarcimento ST)
					dim as TableRow ptr row
					if doc->operacao = ENTRADA then
						row = entradas->AddRow()
					else
						row = saidas->AddRow()
					end if

					if part <> null then
						row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
						row->addCell(part->ie)
						row->addCell(MUNICIPIO2SIGLA(part->municip))
						row->addCell(part->nome)
					else
						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell("")
					end if
					row->addCell(doc->modelo)
					row->addCell(doc->serie)
					row->addCell(doc->numero)
					row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
					row->addCell(YyyyMmDd2Datetime(doc->dataEntSaida))
					row->addCell(doc->chave)
					row->addCell(codSituacao2Str(doc->situacao))
					row->addCell(item->bcICMS)
					row->addCell(item->aliqICMS)
					row->addCell(item->ICMS)
					row->addCell(item->bcICMSST)
					row->addCell(item->aliqICMSST)
					row->addCell(item->ICMSST)
					row->addCell(item->IPI)
					row->addCell(item->valor)
					row->addCell(item->numItem)
					row->addCell(item->qtd)
					row->addCell(item->unidade)
					row->addCell(item->cfop)
					row->addCell(item->cstICMS)
					var itemId = cast( TItemId ptr, itemIdDict->lookup(item->itemId) )
					if itemId <> null then 
						row->addCell(itemId->ncm)
						row->addCell(itemId->id)
						row->addCell(itemId->descricao)
					end if
					row->addCell(getInfoCompl(doc->infoComplListHead))
					row->addCell(getObsLanc(doc->obsListHead))
				end if

			'NF-e?
			case DOC_NF, DOC_NFSCT, DOC_NF_ELETRIC
				var nf = cast(TDocNF ptr, reg)
				if ISREGULAR(nf->situacao) then
					'' NOTA: não existe itemDoc para saídas (exceto quando há ressarcimento ST), só temos informações básicas do DF-e, 
					'' 	     a não ser que sejam carregados os relatórios .csv do SAFI vindos do infoview
					if nf->operacao = SAIDA or (nf->operacao = ENTRADA and nf->nroItens = 0) or reg->tipo <> DOC_NF then
						dim as TDFe_NFeItem ptr item = null
						if itemNFeSafiFornecido and opcoes->acrescentarDados then
							if len(nf->chave) > 0 then
								var dfe = cast(TDFe_NFe ptr, chaveDFeDict->lookup(nf->chave))
								if dfe <> null then
									item = dfe->itemListHead
								end if
							end if
						end if

						var part = cast( TParticipante ptr, participanteDict->lookup(nf->idParticipante) )

						var emitirLinhas = (opcoes->somenteRessarcimentoST = false) andalso _
							iif(nf->operacao = SAIDA, not opcoes->pularLrs, not opcoes->pularLre)
						if opcoes->filtrarCnpj andalso emitirLinhas then
							if part <> null then
								emitirLinhas = filtrarPorCnpj(part->cnpj, opcoes->listaCnpj())
							end if
						end if

						if opcoes->filtrarChaves andalso emitirLinhas then
							emitirLinhas = filtrarPorChave(nf->chave, opcoes->listaChaves())
						end if

						var anal = iif(item = null, nf->itemAnalListHead, null)
						var analCnt = 1
						
						if emitirLinhas then
							do
								dim as TableRow ptr row
								if nf->operacao = SAIDA then
									row = saidas->AddRow()
								else
									row = entradas->AddRow()
								end if
							
								if part <> null then
									row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
									row->addCell(part->ie)
									row->addCell(MUNICIPIO2SIGLA(part->municip))
									row->addCell(part->nome)
								else
									row->addCell("")
									row->addCell("")
									row->addCell("")
									row->addCell("")
								end if
								row->addCell(nf->modelo)
								row->addCell(nf->serie)
								row->addCell(nf->numero)
								row->addCell(YyyyMmDd2Datetime(nf->dataEmi))
								row->addCell(YyyyMmDd2Datetime(nf->dataEntSaida))
								row->addCell(nf->chave)
								row->addCell(codSituacao2Str(nf->situacao))

								if ((itemNFeSafiFornecido andalso opcoes->acrescentarDados) orelse _
								   cbool((nf->operacao = ENTRADA) andalso (reg->tipo = DOC_NF))) andalso _
								   cbool(item <> null) then
									row->addCell(item->bcICMS)
									row->addCell(item->aliqICMS)
									row->addCell(item->ICMS)
									row->addCell(item->bcICMSST)
									row->addCell("")
									row->addCell("")
									row->addCell(item->IPI)
									row->addCell(item->valorProduto)
									row->addCell(item->nroItem)
									row->addCell(item->qtd)
									row->addCell(item->unidade)
									row->addCell(item->cfop)
									row->addCell(item->cst)
									row->addCell(item->ncm)
									row->addCell(item->codProduto)
									row->addCell(item->descricao)

								else
									if anal = null then
										row->addCell(nf->bcICMS)
										row->addCell("")
										row->addCell(nf->ICMS)
										row->addCell(nf->bcICMSST)
										row->addCell(nf->ICMSST)
										row->addCell("")
										row->addCell(nf->IPI)
										row->addCell(nf->valorTotal)
										for cell as integer = 1 to 16-8
											row->addCell("")
										next
									else
										row->addCell(anal->bc)
										row->addCell(anal->aliq)
										row->addCell(anal->ICMS)
										row->addCell(anal->bcST)
										row->addCell("")
										row->addCell(anal->ICMSST)
										row->addCell(anal->IPI)
										row->addCell(anal->valorOp)
										row->addCell(analCnt)
										row->addCell(0)
										row->addCell("")
										row->addCell(anal->cfop)
										row->addCell(anal->cst)
										for cell as integer = 1 to 3
											row->addCell("")
										next
										analCnt += 1
									end if
								end if

								if nf->operacao = SAIDA then
									row->addCell(nf->difal.fcp)
									row->addCell(nf->difal.icmsOrigem)
									row->addCell(nf->difal.icmsDest)
								end if
								
								row->addCell(getInfoCompl(nf->infoComplListHead))
								row->addCell(getObsLanc(nf->obsListHead))
							
								if item = null then
									if anal = null then
										exit do
									end if
									
									anal = anal->next_
								else
									item = item->next_
								end if
								
							loop while (item <> null) or (anal <> null)
						end if
					
					end if
			   
				else
					var emitirLinha = (opcoes->somenteRessarcimentoST = false) andalso _
						iif(nf->operacao = SAIDA, not opcoes->pularLrs, not opcoes->pularLre)
					
					if emitirLinha then
						var row = iif(nf->operacao = SAIDA, saidas, entradas)->AddRow()

						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell(nf->modelo)
						row->addCell(nf->serie)
						row->addCell(nf->numero)
						'' NOTA: cancelados e inutilizados não vêm com a data preenchida, então retiramos a data da chave ou do registro mestre
						var dataEmi = iif( len(nf->chave) = 44, "20" + mid(nf->chave,3,2) + mid(nf->chave,5,2) + "01", regMestre->dataIni )
						row->addCell(YyyyMmDd2Datetime(dataEmi))
						row->addCell("")
						row->addCell(nf->chave)
						row->addCell(codSituacao2Str(nf->situacao))
					end if
				end if
				
			'ressarcimento st?
			case DOC_NF_ITEM_RESSARC_ST
				var item = cast(TDocNFItemRessarcSt ptr, reg)
				var part = cast( TParticipante ptr, participanteDict->lookup(item->idParticipanteUlt) )

				var emitirLinha = not opcoes->pularLre
				if opcoes->filtrarCnpj andalso emitirLinha then
					if part <> null then
						emitirLinha = filtrarPorCnpj(part->cnpj, opcoes->listaCnpj())
					end if
				end if

				if emitirLinha then
					var row = ressarcST->AddRow()

					if part <> null then
						row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
						row->addCell(part->ie)
						row->addCell(MUNICIPIO2SIGLA(part->municip))
						row->addCell(part->nome)
					else
						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell("")
					end if
					row->addCell(item->modeloUlt)
					row->addCell(item->serieUlt)
					row->addCell(item->numeroUlt)
					row->addCell(YyyyMmDd2Datetime(item->dataUlt))
					row->addCell(item->qtdUlt)
					row->addCell(item->valorUlt)
					row->addCell(item->valorBcST)
					row->addCell(item->chaveNFeUlt)
					row->addCell(item->numItemNFeUlt)
					row->addCell(item->bcIcmsUlt)
					row->addCell(item->aliqIcmsUlt)
					row->addCell(item->limiteBcIcmsUlt)
					row->addCell(item->icmsUlt)
					row->addCell(item->aliqIcmsStUlt)
					row->addCell(item->res)
					row->addCell(item->responsavelRet)
					row->addCell(item->motivo)
					row->addCell(item->tipDocArrecadacao)
					row->addCell(item->numDocArrecadacao)
					row->addCell(item->documentoPai->documentoPai->chave)
					row->addCell(item->documentoPai->numItem)
				end if

			'CT-e?
			case DOC_CT
				var ct = cast(TDocCT ptr, reg)
				if ISREGULAR(ct->situacao) then
					var part = cast( TParticipante ptr, participanteDict->lookup(ct->idParticipante) )

					var emitirLinhas = (opcoes->somenteRessarcimentoST = false) and _
						iif(ct->operacao = SAIDA, not opcoes->pularLrs, not opcoes->pularLre)
					
					if opcoes->filtrarCnpj andalso emitirLinhas then
						if part <> null then
							emitirLinhas = filtrarPorCnpj(part->cnpj, opcoes->listaCnpj())
						end if
					end if

					if opcoes->filtrarChaves andalso emitirLinhas then
						emitirLinhas = filtrarPorChave(ct->chave, opcoes->listaChaves())
					end if
						
					if emitirLinhas then
						dim as TDocItemAnal ptr item = null
						if ct->operacao = ENTRADA then
							item = ct->itemAnalListHead
						end if
						
						var itemCnt = 1
						do
							dim as TableRow ptr row 
							if ct->operacao = SAIDA then
								row = saidas->AddRow()
							else
								row = entradas->AddRow()
							end if
							
							if part <> null then
								row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
								row->addCell(part->ie)
								row->addCell(MUNICIPIO2SIGLA(part->municip))
								row->addCell(part->nome)
							else
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
							end if
							row->addCell(ct->modelo)
							row->addCell(ct->serie)
							row->addCell(ct->numero)
							row->addCell(YyyyMmDd2Datetime(ct->dataEmi))
							row->addCell(YyyyMmDd2Datetime(ct->dataEntSaida))
							row->addCell(ct->chave)
							row->addCell(codSituacao2Str(ct->situacao))
							
							if item <> null then
								row->addCell(item->bc)
								row->addCell(item->aliq)
								row->addCell(item->ICMS)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell(item->valorOp)
								row->addCell(itemCnt)
								row->addCell("")
								row->addCell("")
								row->addCell(item->cfop)
								row->addCell(item->cst)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								
								item = item->next_
								itemCnt += 1
							else
								row->addCell(ct->bcICMS)
								row->addCell("")
								row->addCell(ct->ICMS)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell(ct->valorServico)
								row->addCell(1)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								
							end if

							if ct->operacao = SAIDA then
								row->addCell(ct->difal.fcp)
								row->addCell(ct->difal.icmsOrigem)
								row->addCell(ct->difal.icmsDest)
							end if
							
						loop while item <> null
					end if
				
				else
					var emitirLinhas = (opcoes->somenteRessarcimentoST = false) and _
						iif(ct->operacao = SAIDA, not opcoes->pularLrs, not opcoes->pularLre)

					if emitirLinhas then
						var row = iif(ct->operacao = SAIDA, saidas, entradas)->AddRow()

						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell(ct->modelo)
						row->addCell(ct->serie)
						row->addCell(ct->numero)
						'' NOTA: cancelados e inutilizados não vêm com a data preenchida, então retiramos a data da chave ou do registro mestre
						var dataEmi = iif( len(ct->chave) = 44, "20" + mid(ct->chave,3,2) + mid(ct->chave,5,2) + "01", regMestre->dataIni )
						row->addCell(YyyyMmDd2Datetime(dataEmi))
						row->addCell("")
						row->addCell(ct->chave)
						row->addCell(codSituacao2Str(ct->situacao))
					end if
				
				end if
				
			'item de ECF?
			case DOC_ECF_ITEM
				if not opcoes->pularLrs then
					var item = cast(TDocECFItem ptr, reg)
					var doc = item->documentoPai
					if ISREGULAR(doc->situacao) then
						'só existe cupom para saída
						if doc->operacao = SAIDA then
							var emitirLinha = (opcoes->somenteRessarcimentoST = false)
							if opcoes->filtrarCnpj andalso emitirLinha then
								emitirLinha = filtrarPorCnpj(doc->cpfCnpjAdquirente, opcoes->listaCnpj())
							end if

							if opcoes->filtrarChaves andalso emitirLinha then
								emitirLinha = filtrarPorChave(doc->chave, opcoes->listaChaves())
							end if
							
							if emitirLinha then
								var row = saidas->AddRow()

								row->addCell(doc->cpfCnpjAdquirente)
								row->addCell("")
								row->addCell("SP")
								row->addCell(doc->nomeAdquirente)
								row->addCell(iif(doc->modelo = &h2D, "2D", str(doc->modelo)))
								row->addCell("")
								row->addCell(doc->numero)
								row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
								row->addCell(YyyyMmDd2Datetime(doc->dataEntSaida))
								row->addCell(doc->chave)
								row->addCell(codSituacao2Str(doc->situacao))
								row->addCell("")
								row->addCell(item->aliqICMS)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell(item->valor)
								row->addCell(item->numItem)
								row->addCell(item->qtd)
								row->addCell(item->unidade)
								row->addCell(item->cfop)
								row->addCell(item->cstICMS)
								var itemId = cast( TItemId ptr, itemIdDict->lookup(item->itemId) )
								if itemId <> null then 
									row->addCell(itemId->ncm)
									row->addCell(itemId->id)
									row->addCell(itemId->descricao)
								end if
							end if
						end if
					end if
				end if
				
			'SAT?
			case DOC_SAT
				if not opcoes->pularLrs then
					var doc = cast(TDocSAT ptr, reg)
					if ISREGULAR(doc->situacao) then
						'só existe cupom para saída
						if doc->operacao = SAIDA then
							var emitirLinha = (opcoes->somenteRessarcimentoST = false)
							if opcoes->filtrarCnpj andalso emitirLinha then
								emitirLinha = filtrarPorCnpj(doc->cpfCnpjAdquirente, opcoes->listaCnpj())
							end if
							
							if opcoes->filtrarChaves andalso emitirLinha then
								emitirLinha = filtrarPorChave(doc->chave, opcoes->listaChaves())
							end if
							
							if emitirLinha then
								dim as TDFe_NFeItem ptr item = null
								if itemNFeSafiFornecido and opcoes->acrescentarDados then
									var dfe = cast(TDFe_NFe ptr, chaveDFeDict->lookup(doc->chave))
									if dfe <> null then
										item = dfe->itemListHead
									end if
								end if
								
								var anal = iif(item = null, doc->itemAnalListHead, null)
								
								var analCnt = 1
								do
									var row = saidas->AddRow()

									row->addCell(doc->cpfCnpjAdquirente)
									row->addCell("")
									row->addCell("SP")
									row->addCell("")
									row->addCell(str(doc->modelo))
									row->addCell("")
									row->addCell(doc->numero)
									row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
									row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
									row->addCell(doc->chave)
									row->addCell(codSituacao2Str(doc->situacao))
									if item <> null then
										row->addCell(item->bcICMS)
										row->addCell(item->aliqICMS)
										row->addCell(item->ICMS)
										row->addCell("")
										row->addCell("")
										row->addCell("")
										row->addCell("")
										row->addCell(item->valorProduto)
										row->addCell(item->nroItem)
										row->addCell(item->qtd)
										row->addCell(item->unidade)
										row->addCell(item->cfop)
										row->addCell(item->cst)
										row->addCell(item->ncm)
										row->addCell(item->codProduto)
										row->addCell(item->descricao)
										
										item = item->next_
										if item = null then
											exit do
										end if
										
									else
										if anal = null then
											exit do
										end if
											
										row->addCell("")
										row->addCell(anal->aliq)
										row->addCell("")
										row->addCell("")
										row->addCell("")
										row->addCell("")
										row->addCell("")
										row->addCell(anal->valorOp)
										row->addCell(analCnt)
										row->addCell("")
										row->addCell("")
										row->addCell(anal->cfop)
										row->addCell(anal->cst)
										row->addCell("")
										row->addCell("")
										row->addCell("")
										
										analCnt += 1
										anal = anal->next_
										if anal = null then
											exit do
										end if
									end if
								loop
							end if
						end if
					end if
				end if
				
			case APURACAO_ICMS_PERIODO
				if not opcoes->pularLraicms then
					var row = apuracaoIcms->AddRow()

					var apu = cast(TApuracaoIcmsPropPeriodo ptr, reg)
					row->addCell(YyyyMmDd2Datetime(apu->dataIni))
					row->addCell(YyyyMmDd2Datetime(apu->dataFim))
					row->addCell(apu->totalDebitos)
					row->addCell(apu->ajustesDebitos)
					row->addCell(apu->totalAjusteDeb)
					row->addCell(apu->estornosCredito)
					row->addCell(apu->totalCreditos)
					row->addCell(apu->ajustesCreditos)
					row->addCell(apu->totalAjusteCred)
					row->addCell(apu->estornoDebitos)
					row->addCell(apu->saldoCredAnterior)
					row->addCell(apu->saldoDevedorApurado)
					row->addCell(apu->totalDeducoes)
					row->addCell(apu->icmsRecolher)
					row->addCell(apu->saldoCredTransportar)
					row->addCell(apu->debExtraApuracao)
					
					var detalhe = ""
					var ajuste = apu->ajustesListHead
					var cnt = 1
					do while ajuste <> null andalso cnt <= MAX_AJUSTES
						row->addCell("{'codigo':'" & ajuste->codigo & "', 'valor':'" & DBL2MONEYBR(ajuste->valor) & "', 'descricao':'" & ajuste->descricao) & "'}"
						ajuste = ajuste->next_
						cnt += 1
					loop
					
				end if
				
			case APURACAO_ICMS_ST_PERIODO
				if not opcoes->pularLraicms then
					var row = apuracaoIcmsST->AddRow()

					var apu = cast(TApuracaoIcmsSTPeriodo ptr, reg)
					row->addCell(YyyyMmDd2Datetime(apu->dataIni))
					row->addCell(YyyyMmDd2Datetime(apu->dataFim))
					row->addCell(apu->UF)
					row->addCell(iif(apu->mov=0, "N", "S"))
					row->addCell(apu->saldoCredAnterior)
					row->addCell(apu->devolMercadorias)
					row->addCell(apu->totalRessarciment)
					row->addCell(apu->totalOutrosCred)
					row->addCell(apu->ajustesCreditos)
					row->addCell(apu->totalRetencao)
					row->addCell(apu->totalOutrosDeb)
					row->addCell(apu->ajustesDebitos)
					row->addCell(apu->saldoAntesDed)
					row->addCell(apu->totalDeducoes)
					row->addCell(apu->icmsRecolher)
					row->addCell(apu->saldoCredTransportar)
					row->addCell(apu->debExtraApuracao)
				end if

			case INVENTARIO_ITEM
				var row = inventario->AddRow()
				
				var item = cast(TInventarioItem ptr, reg)
				row->addCell(YyyyMmDd2Datetime(item->dataInventario))

				var itemId = cast( TItemId ptr, itemIdDict->lookup(item->itemId) )
				if itemId <> null then 
					row->addCell(itemId->id)
					row->addCell(itemId->ncm)
					row->addCell(itemId->tipoItem)
					row->addCell(tipoItem2Str(itemId->tipoItem))
					row->addCell(itemId->descricao)
				else
					row->addCell(item->itemId)
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if
				
				row->addCell(item->unidade)
				row->addCell(item->qtd)
				row->addCell(item->valorUnitario)
				row->addCell(item->valorItem)
				row->addCell(item->indPropriedade)
				var part = cast( TParticipante ptr, participanteDict->lookup(item->idParticipante) )
				if part <> null then
					row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
				else
					row->addCell("")
				end if
				row->addCell(item->txtComplementar)
				row->addCell(item->codConta)
				row->addCell(item->valorItemIR)

			case CIAP_ITEM
				if not opcoes->pularCiap then
					var item = cast(TCiapItem ptr, reg)
					if item->docCnt = 0 then
						var row = ciap->AddRow()
						
						var pai = item->pai
						row->addCell(YyyyMmDd2Datetime(pai->dataIni))
						row->addCell(YyyyMmDd2Datetime(pai->dataFim))
						row->addCell(pai->valorTributExpSoma)
						row->addCell(pai->valorTotalSaidas)
						row->addCell(pai->indicePercSaidas)
						
						var bemCiap = cast( TBemCiap ptr, bemCiapDict->lookup(item->bemId) )
						if bemCiap <> null then 
							row->addCell(bemCiap->id)
							row->addCell(bemCiap->descricao)
						else
							row->addCell(item->bemId)
							row->addCell("")
						end if
						
						row->addCell(YyyyMmDd2Datetime(item->dataMov))
						row->addCell(item->tipoMov)
						row->addCell(item->valorIcms)
						row->addCell(item->valorIcmsSt)
						row->addCell(item->valorIcmsFrete)
						row->addCell(item->valorIcmsDifal)
						row->addCell(item->parcela)
						row->addCell(item->valorParcela)
					end if
				end if

			case CIAP_ITEM_DOC
				if not opcoes->pularCiap then
				
					var row = ciap->AddRow()
					
					var doc = cast(TCiapItemDoc ptr, reg)
					var pai = doc->pai
					var avo = pai->pai
					row->addCell(YyyyMmDd2Datetime(avo->dataIni))
					row->addCell(YyyyMmDd2Datetime(avo->dataFim))
					row->addCell(avo->valorTributExpSoma)
					row->addCell(avo->valorTotalSaidas)
					row->addCell(avo->indicePercSaidas)
					
					var bemCiap = cast( TBemCiap ptr, bemCiapDict->lookup(pai->bemId) )
					if bemCiap <> null then 
						row->addCell(bemCiap->id)
						row->addCell(bemCiap->descricao)
					else
						row->addCell(pai->bemId)
						row->addCell("")
					end if
					
					row->addCell(YyyyMmDd2Datetime(pai->dataMov))
					row->addCell(pai->tipoMov)
					row->addCell(pai->valorIcms)
					row->addCell(pai->valorIcmsSt)
					row->addCell(pai->valorIcmsFrete)
					row->addCell(pai->valorIcmsDifal)
					row->addCell(pai->parcela)
					row->addCell(pai->valorParcela)
					
					row->addCell(doc->modelo)
					row->addCell(doc->serie)
					row->addCell(doc->numero)
					row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
					row->addCell(doc->chaveNfe)
					
					var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipante) )
					if part <> null then
						row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
						row->addCell(part->ie)
						row->addCell(MUNICIPIO2SIGLA(part->municip))
						row->addCell(part->nome)
					else
						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell("")
					end if
				end if

			case ESTOQUE_ITEM
				var row = estoque->AddRow()
				
				var item = cast(TEstoqueItem ptr, reg)
				var pai = item->pai
				row->addCell(YyyyMmDd2Datetime(pai->dataIni))
				row->addCell(YyyyMmDd2Datetime(pai->dataFim))
				
				var itemId = cast( TItemId ptr, itemIdDict->lookup(item->itemId) )
				if itemId <> null then 
					row->addCell(itemId->id)
					row->addCell(itemId->ncm)
					row->addCell(itemId->tipoItem)
					row->addCell(tipoItem2Str(itemId->tipoItem))
					row->addCell(itemId->descricao)
				else
					row->addCell(item->itemId)
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if
				
				row->addCell(item->qtd)
				row->addCell(item->tipoEst)

				var part = cast( TParticipante ptr, participanteDict->lookup(item->idParticipante) )
				if part <> null then
					row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
					row->addCell(part->ie)
					row->addCell(MUNICIPIO2SIGLA(part->municip))
					row->addCell(part->nome)
				else
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if

			case ESTOQUE_ORDEM_PROD
				var row = producao->AddRow()
				
				var ord = cast(TEstoqueOrdemProd ptr, reg)
				row->addCell(YyyyMmDd2Datetime(ord->dataIni))
				row->addCell(YyyyMmDd2Datetime(ord->dataFim))
				
				var itemId = cast( TItemId ptr, itemIdDict->lookup(ord->itemId) )
				if itemId <> null then 
					row->addCell(itemId->id)
					row->addCell(itemId->ncm)
					row->addCell(itemId->tipoItem)
					row->addCell(tipoItem2Str(itemId->tipoItem))
					row->addCell(itemId->descricao)
				else
					row->addCell(ord->itemId)
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if
				
				row->addCell(ord->qtd)
				row->addCell(ord->idOrdem)

			'item de documento do sintegra?
			case SINTEGRA_DOCUMENTO_ITEM
				if not opcoes->pularLrs then
					var item = cast(TDocumentoItemSintegra ptr, reg)
					var doc = item->doc
					
					dim as TableRow ptr row 
					if doc->operacao = SAIDA then
						row = saidas->AddRow()
					else
						row = entradas->AddRow()
					end if
					
					var itemId = cast( TItemId ptr, itemIdDict->lookup(item->codMercadoria) )
					  
					row->addCell(doc->cnpj)
					row->addCell(doc->ie)
					row->addCell(ufCod2Sigla(doc->uf))
					row->addCell("")
					row->addCell(doc->modelo)
					row->addCell(doc->serie)
					row->addCell(doc->numero)
					row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
					row->addCell("")
					row->addCell("")
					row->addCell(codSituacao2Str(doc->situacao))
					row->addCell(item->bcICMS)
					row->addCell(item->aliqICMS)
					row->addCell(item->bcICMS * item->aliqICMS / 100)
					row->addCell(item->bcICMSST)
					row->addCell("")
					row->addCell("")
					row->addCell(item->valorIPI)
					row->addCell(item->valor)
					row->addCell(item->nroItem)
					row->addCell(item->qtd)
					if itemId <> null then 
						row->addCell(rtrim(itemId->unidInventario))
					else
						row->addCell("")
					end if
					row->addCell(item->cfop)
					row->addCell(item->cst)
					if itemId <> null then 
						row->addCell(itemId->ncm)
						row->addCell(rtrim(itemId->id))
						row->addCell(rtrim(itemId->descricao))
					end if
				end if

			case LUA_CUSTOM
				
				var l = cast(TLuaReg ptr, reg)
				var luaFunc = cast(customLuaCb ptr, customLuaCbDict->lookup(l->tipo))->writer
				
				if luaFunc <> null then
					lua_getglobal(lua, luaFunc)
					lua_rawgeti(lua, LUA_REGISTRYINDEX, l->table)
					lua_call(lua, 1, 0)
				end if
			
			end select

			regCnt += 1
			if not onProgress(null, regCnt / nroRegs) then
				exit do
			end if
			
			reg = reg->prox
		loop
	catch
		onError(!"\r\nErro ao tratar o registro de tipo (" & reg->tipo & !") carregado na linha (" & reg->linha & !")\r\n")
	endtry
	
	onProgress(null, 1)
	
end sub

''''''''
sub EfdTabelaExport.finalizar()
	onProgress("Gravando planilha: " + nomeArquivo, 0)
	ew->Flush(onProgress, onError)
	ew->Close
end sub


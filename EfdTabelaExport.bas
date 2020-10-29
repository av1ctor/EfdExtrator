#include once "EfdTabelaExport.bi"
#include once "vbcompat.bi"
#include once "trycatch.bi"

''''''''
constructor EfdTabelaExport(nomeArquivo as String, opcoes as OpcoesExtracao ptr)
	this.nomeArquivo = nomeArquivo
	this.opcoes = opcoes
	
	ew = new ExcelWriter
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
function EfdTabelaExport.getPlanilha(nome as const zstring ptr) as ExcelWorksheet ptr
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
private sub adicionarColunasComuns(sheet as ExcelWorksheet ptr, ehEntrada as Boolean)

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
	
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING, 45)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 30)
   
	if not ehEntrada then
		row->addCell("DifAl FCP")
		row->addCell("DifAl ICMS Orig")
		row->addCell("DifAl ICMS Dest")
		
		sheet->AddCellType(CT_MONEY)
		sheet->AddCellType(CT_MONEY)
		sheet->AddCellType(CT_MONEY)
	end if
	
	row->addCell("Info. complementares")
	sheet->AddCellType(CT_STRING, 40)

	row->addCell("Obs. lancamento")
	sheet->AddCellType(CT_STRING, 40)
end sub

private sub criarColunasApuracaoIcms(sheet as ExcelWorksheet ptr)
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
	
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	for i as integer = 1 to MAX_AJUSTES
		sheet->AddCellType(CT_STRING, 80)
	next
end sub

private sub criarColunasApuracaoIcmsST(sheet as ExcelWorksheet ptr)
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

	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	for i as integer = 1 to MAX_AJUSTES
		sheet->AddCellType(CT_STRING, 80)
	next
end sub

private sub criarColunasInventario(sheet as ExcelWorksheet ptr)
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

	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_MONEY)
end sub

private sub criarColunasCIAP(sheet as ExcelWorksheet ptr)
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
	
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 30)
end sub

private sub criarColunasEstoque(sheet as ExcelWorksheet ptr)
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
	
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 30)
end sub

private sub criarColunasProducao(sheet as ExcelWorksheet ptr)
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
	
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_STRING)
end sub

private sub criarColunasRessarcST(sheet as ExcelWorksheet ptr)
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
	
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_STRING, 45)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 45)
	sheet->AddCellType(CT_INTNUMBER)
end sub

''''''''
sub EfdTabelaExport.criarPlanilhas()
	'' planilha de entradas
	entradas = ew->AddWorksheet("Entradas")
	adicionarColunasComuns(entradas, true)

	'' planilha de saídas
	saidas = ew->AddWorksheet("Saidas")
	adicionarColunasComuns(saidas, false)

	'' apuração do ICMS
	apuracaoIcms = ew->AddWorksheet("Apuracao ICMS")
	criarColunasApuracaoIcms(apuracaoIcms)
   
	'' apuração do ICMS ST
	apuracaoIcmsST = ew->AddWorksheet("Apuracao ICMS ST")
	criarColunasApuracaoIcmsST(apuracaoIcmsST)
	
	'' Inventário
	inventario = ew->AddWorksheet("Inventario")
	criarColunasInventario(inventario)

	'' CIAP
	ciap = ew->AddWorksheet("CIAP")
	criarColunasCIAP(ciap)

	'' Estoque
	estoque = ew->AddWorksheet("Estoque")
	criarColunasEstoque(estoque)

	'' Producao
	producao = ew->AddWorksheet("Producao")
	criarColunasProducao(producao)

	'' Ressarcimento ST
	ressarcST = ew->AddWorksheet("Ressarcimento ST")
	criarColunasRessarcST(ressarcST)
	
	'' Inconsistencias LRE
	inconsistenciasLRE = ew->AddWorksheet("Inconsistencias LRE")

	'' Inconsistencias LRS
	inconsistenciasLRS = ew->AddWorksheet("Inconsistencias LRS")
	
	'' Resumos LRE
	resumosLRE = ew->AddWorksheet("Resumos LRE")

	'' Resumos LRS
	resumosLRS = ew->AddWorksheet("Resumos LRS")
	
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
sub EfdTabelaExport.gerar(regListHead as TRegistro ptr, regMestre as TRegistro ptr, nroRegs as integer)
	
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
				var doc = reg->itemNF.documentoPai
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
					emitirLinha = reg->itemNF.itemRessarcStListHead <> null
				end if
				
				if emitirLinha then
					'só existe item para entradas (exceto quando há ressarcimento ST)
					dim as ExcelRow ptr row
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
					row->addCell(reg->itemNF.bcICMS)
					row->addCell(reg->itemNF.aliqICMS)
					row->addCell(reg->itemNF.ICMS)
					row->addCell(reg->itemNF.bcICMSST)
					row->addCell(reg->itemNF.aliqICMSST)
					row->addCell(reg->itemNF.ICMSST)
					row->addCell(reg->itemNF.IPI)
					row->addCell(reg->itemNF.valor)
					row->addCell(reg->itemNF.numItem)
					row->addCell(reg->itemNF.qtd)
					row->addCell(reg->itemNF.unidade)
					row->addCell(reg->itemNF.cfop)
					row->addCell(reg->itemNF.cstICMS)
					var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->itemNF.itemId) )
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
				if ISREGULAR(reg->nf.situacao) then
					'' NOTA: não existe itemDoc para saídas (exceto quando há ressarcimento ST), só temos informações básicas do DF-e, 
					'' 	     a não ser que sejam carregados os relatórios .csv do SAFI vindos do infoview
					if reg->nf.operacao = SAIDA or (reg->nf.operacao = ENTRADA and reg->nf.nroItens = 0) or reg->tipo <> DOC_NF then
						dim as TDFe_NFeItem ptr item = null
						if itemNFeSafiFornecido and opcoes->acrescentarDados then
							if len(reg->nf.chave) > 0 then
								var dfe = cast( TDFe ptr, chaveDFeDict->lookup(reg->nf.chave) )
								if dfe <> null then
									item = dfe->nfe.itemListHead
								end if
							end if
						end if

						var part = cast( TParticipante ptr, participanteDict->lookup(reg->nf.idParticipante) )

						var emitirLinhas = (opcoes->somenteRessarcimentoST = false) and _
							iif(reg->nf.operacao = SAIDA, not opcoes->pularLrs, not opcoes->pularLre)
						if opcoes->filtrarCnpj andalso emitirLinhas then
							if part <> null then
								emitirLinhas = filtrarPorCnpj(part->cnpj, opcoes->listaCnpj())
							end if
						end if

						if opcoes->filtrarChaves andalso emitirLinhas then
							emitirLinhas = filtrarPorChave(reg->nf.chave, opcoes->listaChaves())
						end if

						var anal = iif(item = null, reg->nf.itemAnalListHead, null)
						var analCnt = 1
						
						if emitirLinhas then
							do
								dim as ExcelRow ptr row
								if reg->nf.operacao = SAIDA then
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
								row->addCell(reg->nf.modelo)
								row->addCell(reg->nf.serie)
								row->addCell(reg->nf.numero)
								row->addCell(YyyyMmDd2Datetime(reg->nf.dataEmi))
								row->addCell(YyyyMmDd2Datetime(reg->nf.dataEntSaida))
								row->addCell(reg->nf.chave)
								row->addCell(codSituacao2Str(reg->nf.situacao))

								if ((itemNFeSafiFornecido and opcoes->acrescentarDados) or _
								   cbool((reg->nf.operacao = ENTRADA) and (reg->tipo = DOC_NF))) and _
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
										row->addCell(reg->nf.bcICMS)
										row->addCell("")
										row->addCell(reg->nf.ICMS)
										row->addCell(reg->nf.bcICMSST)
										row->addCell(reg->nf.ICMSST)
										row->addCell("")
										row->addCell(reg->nf.IPI)
										row->addCell(reg->nf.valorTotal)
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

								if reg->nf.operacao = SAIDA then
									row->addCell(reg->nf.difal.fcp)
									row->addCell(reg->nf.difal.icmsOrigem)
									row->addCell(reg->nf.difal.icmsDest)
								end if
								
								row->addCell(getInfoCompl(reg->nf.infoComplListHead))
								row->addCell(getObsLanc(reg->nf.obsListHead))
							
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
						iif(reg->nf.operacao = SAIDA, not opcoes->pularLrs, not opcoes->pularLre)
					
					if emitirLinha then
						var row = iif(reg->nf.operacao = SAIDA, saidas, entradas)->AddRow()

						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell(reg->nf.modelo)
						row->addCell(reg->nf.serie)
						row->addCell(reg->nf.numero)
						'' NOTA: cancelados e inutilizados não vêm com a data preenchida, então retiramos a data da chave ou do registro mestre
						var dataEmi = iif( len(reg->nf.chave) = 44, "20" + mid(reg->nf.chave,3,2) + mid(reg->nf.chave,5,2) + "01", regMestre->mestre.dataIni )
						row->addCell(YyyyMmDd2Datetime(dataEmi))
						row->addCell("")
						row->addCell(reg->nf.chave)
						row->addCell(codSituacao2Str(reg->nf.situacao))
					end if
				end if
				
			'ressarcimento st?
			case DOC_NF_ITEM_RESSARC_ST
				var doc = @reg->itemRessarcSt
				var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipanteUlt) )

				var emitirLinha = iif(reg->ct.operacao = SAIDA, not opcoes->pularLrs, not opcoes->pularLre)
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
					row->addCell(doc->modeloUlt)
					row->addCell(doc->serieUlt)
					row->addCell(doc->numeroUlt)
					row->addCell(YyyyMmDd2Datetime(doc->dataUlt))
					row->addCell(doc->qtdUlt)
					row->addCell(doc->valorUlt)
					row->addCell(doc->valorBcST)
					row->addCell(doc->chaveNFeUlt)
					row->addCell(doc->numItemNFeUlt)
					row->addCell(doc->bcIcmsUlt)
					row->addCell(doc->aliqIcmsUlt)
					row->addCell(doc->limiteBcIcmsUlt)
					row->addCell(doc->icmsUlt)
					row->addCell(doc->aliqIcmsStUlt)
					row->addCell(doc->res)
					row->addCell(doc->responsavelRet)
					row->addCell(doc->motivo)
					row->addCell(doc->tipDocArrecadacao)
					row->addCell(doc->numDocArrecadacao)
					row->addCell(doc->documentoPai->documentoPai->chave)
					row->addCell(doc->documentoPai->numItem)
				end if

			'CT-e?
			case DOC_CT
				if ISREGULAR(reg->ct.situacao) then
					var part = cast( TParticipante ptr, participanteDict->lookup(reg->ct.idParticipante) )

					var emitirLinhas = (opcoes->somenteRessarcimentoST = false) and _
						iif(reg->ct.operacao = SAIDA, not opcoes->pularLrs, not opcoes->pularLre)
					
					if opcoes->filtrarCnpj andalso emitirLinhas then
						if part <> null then
							emitirLinhas = filtrarPorCnpj(part->cnpj, opcoes->listaCnpj())
						end if
					end if

					if opcoes->filtrarChaves andalso emitirLinhas then
						emitirLinhas = filtrarPorChave(reg->ct.chave, opcoes->listaChaves())
					end if
						
					if emitirLinhas then
						dim as TDocItemAnal ptr item = null
						if reg->ct.operacao = ENTRADA then
							item = reg->ct.itemAnalListHead
						end if
						
						var itemCnt = 1
						do
							dim as ExcelRow ptr row 
							if reg->ct.operacao = SAIDA then
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
							row->addCell(reg->ct.modelo)
							row->addCell(reg->ct.serie)
							row->addCell(reg->ct.numero)
							row->addCell(YyyyMmDd2Datetime(reg->ct.dataEmi))
							row->addCell(YyyyMmDd2Datetime(reg->ct.dataEntSaida))
							row->addCell(reg->ct.chave)
							row->addCell(codSituacao2Str(reg->ct.situacao))
							
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
								row->addCell(reg->ct.bcICMS)
								row->addCell("")
								row->addCell(reg->ct.ICMS)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell(reg->ct.valorServico)
								row->addCell(1)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								
							end if

							if reg->ct.operacao = SAIDA then
								row->addCell(reg->ct.difal.fcp)
								row->addCell(reg->ct.difal.icmsOrigem)
								row->addCell(reg->ct.difal.icmsDest)
							end if
							
						loop while item <> null
					end if
				
				else
					var emitirLinhas = (opcoes->somenteRessarcimentoST = false) and _
						iif(reg->ct.operacao = SAIDA, not opcoes->pularLrs, not opcoes->pularLre)

					if emitirLinhas then
						var row = iif(reg->ct.operacao = SAIDA, saidas, entradas)->AddRow()

						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell(reg->ct.modelo)
						row->addCell(reg->ct.serie)
						row->addCell(reg->ct.numero)
						'' NOTA: cancelados e inutilizados não vêm com a data preenchida, então retiramos a data da chave ou do registro mestre
						var dataEmi = iif( len(reg->ct.chave) = 44, "20" + mid(reg->ct.chave,3,2) + mid(reg->ct.chave,5,2) + "01", regMestre->mestre.dataIni )
						row->addCell(YyyyMmDd2Datetime(dataEmi))
						row->addCell("")
						row->addCell(reg->ct.chave)
						row->addCell(codSituacao2Str(reg->ct.situacao))
					end if
				
				end if
				
			'item de ECF?
			case DOC_ECF_ITEM
				if not opcoes->pularLrs then
					var doc = reg->itemECF.documentoPai
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
								row->addCell(reg->itemECF.aliqICMS)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell(reg->itemECF.valor)
								row->addCell(reg->itemECF.numItem)
								row->addCell(reg->itemECF.qtd)
								row->addCell(reg->itemECF.unidade)
								row->addCell(reg->itemECF.cfop)
								row->addCell(reg->itemECF.cstICMS)
								var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->itemECF.itemId) )
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
					var doc = @reg->sat
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
									var dfe = cast( TDFe ptr, chaveDFeDict->lookup(doc->chave) )
									if dfe <> null then
										item = dfe->nfe.itemListHead
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

					row->addCell(YyyyMmDd2Datetime(reg->apuIcms.dataIni))
					row->addCell(YyyyMmDd2Datetime(reg->apuIcms.dataFim))
					row->addCell(reg->apuIcms.totalDebitos)
					row->addCell(reg->apuIcms.ajustesDebitos)
					row->addCell(reg->apuIcms.totalAjusteDeb)
					row->addCell(reg->apuIcms.estornosCredito)
					row->addCell(reg->apuIcms.totalCreditos)
					row->addCell(reg->apuIcms.ajustesCreditos)
					row->addCell(reg->apuIcms.totalAjusteCred)
					row->addCell(reg->apuIcms.estornoDebitos)
					row->addCell(reg->apuIcms.saldoCredAnterior)
					row->addCell(reg->apuIcms.saldoDevedorApurado)
					row->addCell(reg->apuIcms.totalDeducoes)
					row->addCell(reg->apuIcms.icmsRecolher)
					row->addCell(reg->apuIcms.saldoCredTransportar)
					row->addCell(reg->apuIcms.debExtraApuracao)
					
					var detalhe = ""
					var ajuste = reg->apuIcms.ajustesListHead
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

					row->addCell(YyyyMmDd2Datetime(reg->apuIcmsST.dataIni))
					row->addCell(YyyyMmDd2Datetime(reg->apuIcmsST.dataFim))
					row->addCell(reg->apuIcmsST.UF)
					row->addCell(iif(reg->apuIcmsST.mov=0, "N", "S"))
					row->addCell(reg->apuIcmsST.saldoCredAnterior)
					row->addCell(reg->apuIcmsST.devolMercadorias)
					row->addCell(reg->apuIcmsST.totalRessarciment)
					row->addCell(reg->apuIcmsST.totalOutrosCred)
					row->addCell(reg->apuIcmsST.ajustesCreditos)
					row->addCell(reg->apuIcmsST.totalRetencao)
					row->addCell(reg->apuIcmsST.totalOutrosDeb)
					row->addCell(reg->apuIcmsST.ajustesDebitos)
					row->addCell(reg->apuIcmsST.saldoAntesDed)
					row->addCell(reg->apuIcmsST.totalDeducoes)
					row->addCell(reg->apuIcmsST.icmsRecolher)
					row->addCell(reg->apuIcmsST.saldoCredTransportar)
					row->addCell(reg->apuIcmsST.debExtraApuracao)
				end if

			case INVENTARIO_ITEM
				var row = inventario->AddRow()
				
				row->addCell(YyyyMmDd2Datetime(reg->invItem.dataInventario))

				var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->invItem.itemId) )
				if itemId <> null then 
					row->addCell(itemId->id)
					row->addCell(itemId->ncm)
					row->addCell(itemId->tipoItem)
					row->addCell(tipoItem2Str(itemId->tipoItem))
					row->addCell(itemId->descricao)
				else
					row->addCell(reg->invItem.itemId)
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if
				
				row->addCell(reg->invItem.unidade)
				row->addCell(reg->invItem.qtd)
				row->addCell(reg->invItem.valorUnitario)
				row->addCell(reg->invItem.valorItem)
				row->addCell(reg->invItem.indPropriedade)
				var part = cast( TParticipante ptr, participanteDict->lookup(reg->invItem.idParticipante) )
				if part <> null then
					row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
				else
					row->addCell("")
				end if
				row->addCell(reg->invItem.txtComplementar)
				row->addCell(reg->invItem.codConta)
				row->addCell(reg->invItem.valorItemIR)

			case CIAP_ITEM
				if not opcoes->pularCiap then
					if reg->ciapItem.docCnt = 0 then
						var row = ciap->AddRow()
						
						var pai = reg->ciapItem.pai
						row->addCell(YyyyMmDd2Datetime(pai->dataIni))
						row->addCell(YyyyMmDd2Datetime(pai->dataFim))
						row->addCell(pai->valorTributExpSoma)
						row->addCell(pai->valorTotalSaidas)
						row->addCell(pai->indicePercSaidas)
						
						var bemCiap = cast( TBemCiap ptr, bemCiapDict->lookup(reg->ciapItem.bemId) )
						if bemCiap <> null then 
							row->addCell(bemCiap->id)
							row->addCell(bemCiap->descricao)
						else
							row->addCell(reg->ciapItem.bemId)
							row->addCell("")
						end if
						
						row->addCell(YyyyMmDd2Datetime(reg->ciapItem.dataMov))
						row->addCell(reg->ciapItem.tipoMov)
						row->addCell(reg->ciapItem.valorIcms)
						row->addCell(reg->ciapItem.valorIcmsSt)
						row->addCell(reg->ciapItem.valorIcmsFrete)
						row->addCell(reg->ciapItem.valorIcmsDifal)
						row->addCell(reg->ciapItem.parcela)
						row->addCell(reg->ciapItem.valorParcela)
					end if
				end if

			case CIAP_ITEM_DOC
				if not opcoes->pularCiap then
				
					var row = ciap->AddRow()
					
					var pai = reg->ciapItemDoc.pai
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
					
					row->addCell(reg->ciapItemDoc.modelo)
					row->addCell(reg->ciapItemDoc.serie)
					row->addCell(reg->ciapItemDoc.numero)
					row->addCell(YyyyMmDd2Datetime(reg->ciapItemDoc.dataEmi))
					row->addCell(reg->ciapItemDoc.chaveNfe)
					
					var part = cast( TParticipante ptr, participanteDict->lookup(reg->ciapItemDoc.idParticipante) )
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
				
				var pai = reg->estItem.pai
				row->addCell(YyyyMmDd2Datetime(pai->dataIni))
				row->addCell(YyyyMmDd2Datetime(pai->dataFim))
				
				var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->estItem.itemId) )
				if itemId <> null then 
					row->addCell(itemId->id)
					row->addCell(itemId->ncm)
					row->addCell(itemId->tipoItem)
					row->addCell(tipoItem2Str(itemId->tipoItem))
					row->addCell(itemId->descricao)
				else
					row->addCell(reg->estItem.itemId)
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if
				
				row->addCell(reg->estItem.qtd)
				row->addCell(reg->estItem.tipoEst)

				var part = cast( TParticipante ptr, participanteDict->lookup(reg->estItem.idParticipante) )
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
				
				row->addCell(YyyyMmDd2Datetime(reg->estOrdem.dataIni))
				row->addCell(YyyyMmDd2Datetime(reg->estOrdem.dataFim))
				
				var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->estOrdem.itemId) )
				if itemId <> null then 
					row->addCell(itemId->id)
					row->addCell(itemId->ncm)
					row->addCell(itemId->tipoItem)
					row->addCell(tipoItem2Str(itemId->tipoItem))
					row->addCell(itemId->descricao)
				else
					row->addCell(reg->estOrdem.itemId)
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if
				
				row->addCell(reg->estOrdem.qtd)
				row->addCell(reg->estOrdem.idOrdem)

			'item de documento do sintegra?
			case SINTEGRA_DOCUMENTO_ITEM
				if not opcoes->pularLrs then
					var doc = reg->docItemSint.doc
					
					dim as ExcelRow ptr row 
					if doc->operacao = SAIDA then
						row = saidas->AddRow()
					else
						row = entradas->AddRow()
					end if
					
					var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->docItemSint.codMercadoria) )
					  
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
					row->addCell(reg->docItemSint.bcICMS)
					row->addCell(reg->docItemSint.aliqICMS)
					row->addCell(reg->docItemSint.bcICMS * reg->docItemSint.aliqICMS / 100)
					row->addCell(reg->docItemSint.bcICMSST)
					row->addCell("")
					row->addCell(reg->docSint.ICMSST)
					row->addCell(reg->docItemSint.valorIPI)
					row->addCell(reg->docItemSint.valor)
					row->addCell(reg->docItemSint.nroItem)
					row->addCell(reg->docItemSint.qtd)
					if itemId <> null then 
						row->addCell(rtrim(itemId->unidInventario))
					else
						row->addCell("")
					end if
					row->addCell(reg->docItemSint.cfop)
					row->addCell(reg->docItemSint.cst)
					if itemId <> null then 
						row->addCell(itemId->ncm)
						row->addCell(rtrim(itemId->id))
						row->addCell(rtrim(itemId->descricao))
					end if
				end if

			case LUA_CUSTOM
				
				var luaFunc = cast(customLuaCb ptr, customLuaCbDict->lookup(reg->lua.tipo))->writer
				
				if luaFunc <> null then
					lua_getglobal(lua, luaFunc)
					lua_rawgeti(lua, LUA_REGISTRYINDEX, reg->lua.table)
					lua_call(lua, 1, 0)
				end if
			
			end select

			regCnt += 1
			if not onProgress(null, regCnt / nroRegs) then
				exit do
			end if
			
			reg = reg->next_
		loop
	catch
		onError(!"\r\nErro ao tratar o registro de tipo (" & reg->tipo & !") carregado na linha (" & reg->linha & !")\r\n")
	endtry
	
	onProgress(null, 1)
	
end sub

''''''''
sub EfdTabelaExport.finalizar()
	onProgress("Gravando planilha: " + nomeArquivo, 0)
	ew->Flush(onProgress)
	ew->Close
end sub


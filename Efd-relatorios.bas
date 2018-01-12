#include once "efd.bi"
#include once "Dict.bi"
#include once "vbcompat.bi"
#include once "DocxFactoryDyn.bi"
#include once "DB.bi"

''''''''
sub Efd.gerarRelatorios(nomeArquivo as string, mostrarProgresso as ProgressoCB)
	
	mostrarProgresso(wstr(!"\tGerando relatórios"), 0)
	
	ultimoRelatorio = -1
	
	'' NOTA: por limitação do DocxFactory, que só consegue trabalhar com um template por vez, 
	''		 precisamos processar entradas primeiro, depois saídas e por último os registros 
	''		 que são sequenciais (LRE e LRS vêm intercalados na EFD)
	
	'' LRE
	iniciarRelatorio(REL_LRE, "entradas", "LRE")
	
	var reg = regListHead
	do while reg <> null
		'para cada registro..
		select case reg->tipo
		'NF-e?
		case DOC_NF
			if reg->nf.operacao = ENTRADA then
				var part = cast( TParticipante ptr, participanteDict[reg->nf.idParticipante] )
				adicionarDocRelatorioEntradas(@reg->nf, part)
			end if
		
		'CT-e?
		case DOC_CT
			if reg->ct.operacao = ENTRADA then
				var part = cast( TParticipante ptr, participanteDict[reg->ct.idParticipante] )
				adicionarDocRelatorioEntradas(@reg->ct, part)
			end if
		end select
		
		reg = reg->next_
	loop
	
	finalizarRelatorio()
	
	'' LRS
	iniciarRelatorio(REL_LRS, "saidas", "LRS")
	
	reg = regListHead
	do while reg <> null
		'para cada registro..
		select case reg->tipo
		'NF-e?
		case DOC_NF
			if reg->nf.operacao = SAIDA then
				var part = cast( TParticipante ptr, participanteDict[reg->nf.idParticipante] )
				adicionarDocRelatorioSaidas(@reg->nf, part)
			end if

		'CT-e?
		case DOC_CT
			if reg->ct.operacao = SAIDA then
				var part = cast( TParticipante ptr, participanteDict[reg->ct.idParticipante] )
				adicionarDocRelatorioSaidas(@reg->ct, part)
			end if
			
		'ECF Redução Z?
		case ECF_REDUCAO_Z
			adicionarDocRelatorioSaidas(@reg->ecfRedZ)
		end select

		reg = reg->next_
	loop
	
	finalizarRelatorio()
	
	'' outros livros..
	reg = regListHead
	do while reg <> null
		'para cada registro..
		select case reg->tipo
		case APURACAO_ICMS_PERIODO
			gerarRelatorioApuracaoICMS(nomeArquivo, reg)

		case APURACAO_ICMS_ST_PERIODO
			gerarRelatorioApuracaoICMSST(nomeArquivo, reg)
			
		end select

		reg = reg->next_
	loop
	
	mostrarProgresso(null, 1)

end sub

''''''''
sub Efd.gerarRelatorioApuracaoICMS(nomeArquivo as string, reg as TRegistro ptr)

	iniciarRelatorio(REL_RAICMS, "apuracao_icms", "RAICMS")
	
	dfwd->setClipboardValueByStrW("grid", "nome", regListHead->mestre.nome)
	dfwd->setClipboardValueByStr("grid", "cnpj", STR2CNPJ(regListHead->mestre.cnpj))
	dfwd->setClipboardValueByStr("grid", "ie", STR2IE(regListHead->mestre.ie))
	dfwd->setClipboardValueByStr("grid", "escrit", YyyyMmDd2DatetimeBR(regListHead->mestre.dataIni) + " a " + YyyyMmDd2DatetimeBR(regListHead->mestre.dataFim))
	dfwd->setClipboardValueByStr("grid", "apur", YyyyMmDd2DatetimeBR(reg->apuIcms.dataIni) + " a " + YyyyMmDd2DatetimeBR(reg->apuIcms.dataFim))
	
	dfwd->setClipboardValueByStr("grid", "saidas", DBL2MONEYBR(reg->apuIcms.totalDebitos))
	dfwd->setClipboardValueByStr("grid", "ajuste_deb", DBL2MONEYBR(reg->apuIcms.ajustesDebitos))
	dfwd->setClipboardValueByStr("grid", "ajuste_deb_imp", DBL2MONEYBR(reg->apuIcms.totalAjusteDeb))
	dfwd->setClipboardValueByStr("grid", "estorno_cred", DBL2MONEYBR(reg->apuIcms.estornosCredito))
	dfwd->setClipboardValueByStr("grid", "credito", DBL2MONEYBR(reg->apuIcms.totalCreditos))
	dfwd->setClipboardValueByStr("grid", "ajuste_cred", DBL2MONEYBR(reg->apuIcms.ajustesCreditos))
	dfwd->setClipboardValueByStr("grid", "ajuste_cred_imp", DBL2MONEYBR(reg->apuIcms.totalAjusteCred))
	dfwd->setClipboardValueByStr("grid", "estorno_deb", DBL2MONEYBR(reg->apuIcms.estornoDebitos))
	dfwd->setClipboardValueByStr("grid", "cred_anterior", DBL2MONEYBR(reg->apuIcms.saldoCredAnterior))
	dfwd->setClipboardValueByStr("grid", "saldo_dev", DBL2MONEYBR(reg->apuIcms.saldoDevedorApurado))
	dfwd->setClipboardValueByStr("grid", "deducoes", DBL2MONEYBR(reg->apuIcms.totalDeducoes))
	dfwd->setClipboardValueByStr("grid", "a_recolher", DBL2MONEYBR(reg->apuIcms.icmsRecolher))
	dfwd->setClipboardValueByStr("grid", "a_transportar", DBL2MONEYBR(reg->apuIcms.saldoCredTransportar))
	dfwd->setClipboardValueByStr("grid", "extra_apu", DBL2MONEYBR(reg->apuIcms.debExtraApuracao))
	
	dfwd->paste("grid")

	finalizarRelatorio()
	
end sub

''''''''
sub Efd.gerarRelatorioApuracaoICMSST(nomeArquivo as string, reg as TRegistro ptr)

	iniciarRelatorio(REL_RAICMSST, "apuracao_icms_st", "RAICMSST_" + reg->apuIcmsST.UF)

	dfwd->setClipboardValueByStrW("grid", "nome", regListHead->mestre.nome)
	dfwd->setClipboardValueByStr("grid", "cnpj", STR2CNPJ(regListHead->mestre.cnpj))
	dfwd->setClipboardValueByStr("grid", "ie", STR2IE(regListHead->mestre.ie))
	dfwd->setClipboardValueByStr("grid", "escrit", YyyyMmDd2DatetimeBR(regListHead->mestre.dataIni) + " a " + YyyyMmDd2DatetimeBR(regListHead->mestre.dataFim))
	dfwd->setClipboardValueByStrW("grid", "apur", YyyyMmDd2DatetimeBR(reg->apuIcmsST.dataIni) + " a " + YyyyMmDd2DatetimeBR(reg->apuIcmsST.dataFim) + " - INSCRIÃ‡ÃƒO ESTADUAL:")
	dfwd->setClipboardValueByStr("grid", "UF", reg->apuIcmsST.UF)
	dfwd->setClipboardValueByStr("grid", "MOV", iif(reg->apuIcmsST.mov, "1 - COM", "0 - SEM"))
	
	dfwd->setClipboardValueByStr("grid", "saldo_cred", DBL2MONEYBR(reg->apuIcmsST.saldoCredAnterior))
	dfwd->setClipboardValueByStr("grid", "devolucoes", DBL2MONEYBR(reg->apuIcmsST.devolMercadorias))
	dfwd->setClipboardValueByStr("grid", "ressarcimentos", DBL2MONEYBR(reg->apuIcmsST.totalRessarciment))
	dfwd->setClipboardValueByStr("grid", "outros_cred", DBL2MONEYBR(reg->apuIcmsST.totalOutrosCred))
	dfwd->setClipboardValueByStr("grid", "ajuste_cred", DBL2MONEYBR(reg->apuIcmsST.ajusteCred))
	dfwd->setClipboardValueByStr("grid", "icms_st", DBL2MONEYBR(reg->apuIcmsST.totalRetencao))
	dfwd->setClipboardValueByStr("grid", "outros_deb", DBL2MONEYBR(reg->apuIcmsST.totalOutrosDeb))
	dfwd->setClipboardValueByStr("grid", "ajuste_deb", DBL2MONEYBR(reg->apuIcmsST.ajusteDeb))
	dfwd->setClipboardValueByStr("grid", "saldo_dev", DBL2MONEYBR(reg->apuIcmsST.saldoAntesDed))
	dfwd->setClipboardValueByStr("grid", "deducoes", DBL2MONEYBR(reg->apuIcmsST.totalDeducoes))
	dfwd->setClipboardValueByStr("grid", "a_recolher", DBL2MONEYBR(reg->apuIcmsST.icmsRecolher))
	dfwd->setClipboardValueByStr("grid", "a_transportar", DBL2MONEYBR(reg->apuIcmsST.saldoCredTransportar))
	dfwd->setClipboardValueByStr("grid", "extra_apu", DBL2MONEYBR(reg->apuIcmsST.debExtraApuracao))

	dfwd->paste("grid")

	finalizarRelatorio()
	
end sub

''''''''
function EFd.codMunicipio2Nome(cod as integer) as string
	
	var nome = cast(zstring ptr, municipDict[cod])
	if nome <> null then
		return *nome
	end if
	
	var nomedb = dbConfig->execScalar("select Nome || ' - ' || uf nome from Municipio where Codigo = " & cod)
	if nomedb = null then
		return ""
	end if
	
	municipDict.add(cod, nomedb)
	
	function = *nomedb
end function

''''''''
sub Efd.iniciarRelatorio(relatorio as TipoRelatorio, nomeRelatorio as string, sufixo as string)

	if ultimoRelatorio = relatorio then
		return
	end if
		
	finalizarRelatorio()
	
	ultimoRelatorioSufixo = sufixo
	ultimoRelatorio = relatorio
	nroRegistrosRel = 0

	dfwd->load(baseTemplatesDir + nomeRelatorio + ".dfw")

	dfwd->setClipboardValueByStrW("_header", "nome", regListHead->mestre.nome)
	dfwd->setClipboardValueByStr("_header", "cnpj", STR2CNPJ(regListHead->mestre.cnpj))
	dfwd->setClipboardValueByStr("_header", "ie", STR2IE(regListHead->mestre.ie))
	dfwd->setClipboardValueByStr("_header", "uf", MUNICIPIO2SIGLA(regListHead->mestre.municip))
	
	select case relatorio
	case REL_LRE, REL_LRS
		dfwd->setClipboardValueByStr("_header", "municipio", codMunicipio2Nome(regListHead->mestre.municip))
		dfwd->setClipboardValueByStr("_header", "apu", YyyyMmDd2DatetimeBR(regListHead->mestre.dataIni) + " a " + YyyyMmDd2DatetimeBR(regListHead->mestre.dataFim))
	
		relSomaLRList.init(10, len(RelSomatorioLR))
		relSomaLRDict.init(10)
		
		dfwd->paste("tabela")
	end select
	
end sub

private function cmpFunc(key as any ptr, node as any ptr) as boolean
	function = *cast(zstring ptr, key) < cast(RelSomatorioLR ptr, node)->chave
end function

''''''''
private sub Efd.relatorioSomarLR(sit as TipoSituacao, anal as TDocItemAnal ptr)
	
	dim as string chave = iif(ultimoRelatorio = REL_LRS, str(sit), "0")
	
	chave &= format(anal->cst,"000") & anal->cfop & format(anal->aliq, "00")
	
	var soma = cast(RelSomatorioLR ptr, relSomaLRDict[chave])
	if soma = null then
		soma = relSomaLRList.addOrdAsc(strptr(chave), @cmpFunc)
		soma->chave = chave
		soma->situacao = sit
		soma->cst = anal->cst
		soma->cfop = anal->cfop
		soma->aliq = anal->aliq
		relSomaLRDict.add(soma->chave, soma)
	end if
	
	soma->valorOp += anal->valorOp
	soma->bc += anal->bc
	soma->icms += anal->icms
	soma->bcST += anal->bcST
	soma->icmsST += anal->icmsST
	soma->ipi += anal->ipi
end sub

''''''''
sub Efd.adicionarDocRelatorioItemAnal(sit as TipoSituacao, anal as TDocItemAnal ptr)
	
	do while anal <> null
		relatorioSomarLR(sit, anal)

		select case sit
		case REGULAR, EXTEMPORANEO
			dfwd->setClipboardValueByStr("linha_anal", "cst", format(anal->cst,"000"))
			dfwd->setClipboardValueByStr("linha_anal", "cfop", anal->cfop)
			dfwd->setClipboardValueByStr("linha_anal", "aliq", DBL2MONEYBR(anal->aliq))
			dfwd->setClipboardValueByStr("linha_anal", "bc", DBL2MONEYBR(anal->bc))
			dfwd->setClipboardValueByStr("linha_anal", "icms", DBL2MONEYBR(anal->ICMS))
			dfwd->setClipboardValueByStr("linha_anal", "bcst", DBL2MONEYBR(anal->bcST))
			dfwd->setClipboardValueByStr("linha_anal", "icmsst", DBL2MONEYBR(anal->ICMSST))
			dfwd->setClipboardValueByStr("linha_anal", "ipi", DBL2MONEYBR(anal->IPI))
			dfwd->setClipboardValueByStr("linha_anal", "valop", DBL2MONEYBR(anal->valorOp))
			if ultimoRelatorio = REL_LRE then
				dfwd->setClipboardValueByStr("linha_anal", "redbc", DBL2MONEYBR(anal->redBC))
			end if
			
			dfwd->paste("linha_anal")
		end select

		anal = anal->next_
	loop
end sub

''''''''
sub Efd.adicionarDocRelatorioSaidas(doc as TDocDF ptr, part as TParticipante ptr)

	if len(doc->dataEmi) > 0 then
		dfwd->setClipboardValueByStr("linha", "demi", YyyyMmDd2DatetimeBR(doc->dataEmi))
	end if
	if len(doc->dataEntSaida) > 0 then
		dfwd->setClipboardValueByStr("linha", "dsaida", YyyyMmDd2DatetimeBR(doc->dataEntSaida))
	end if
	dfwd->setClipboardValueByStr("linha", "nrini", doc->numero)
	dfwd->setClipboardValueByStr("linha", "md", doc->modelo)
	dfwd->setClipboardValueByStr("linha", "sr", doc->serie)
	dfwd->setClipboardValueByStr("linha", "sit", format(cdbl(doc->situacao), "00"))
	
	select case doc->situacao
	case REGULAR, EXTEMPORANEO
		dfwd->setClipboardValueByStr("linha", "cnpjdest", STR2CNPJ(part->cnpj))
		dfwd->setClipboardValueByStr("linha", "iedest", STR2IE(part->ie))
		dfwd->setClipboardValueByStr("linha", "uf", MUNICIPIO2SIGLA(part->municip))
		dfwd->setClipboardValueByStr("linha", "mundest", part->municip)
		dfwd->setClipboardValueByStrW("linha", "razaodest", left(part->nome, 64))
	end select
	
	dfwd->paste("linha")
	
	adicionarDocRelatorioItemAnal(doc->situacao, doc->itemAnalListHead)
	
	nroRegistrosRel += 1
	
end sub

''''''''
sub Efd.adicionarDocRelatorioEntradas(doc as TDocDF ptr, part as TParticipante ptr)

	dfwd->setClipboardValueByStr("linha", "demi", YyyyMmDd2DatetimeBR(doc->dataEmi))
	dfwd->setClipboardValueByStr("linha", "dent", YyyyMmDd2DatetimeBR(doc->dataEntSaida))
	dfwd->setClipboardValueByStr("linha", "nro", doc->numero)
	dfwd->setClipboardValueByStr("linha", "mod", doc->modelo)
	dfwd->setClipboardValueByStr("linha", "ser", doc->serie)
	dfwd->setClipboardValueByStr("linha", "sit", format(cdbl(doc->situacao), "00"))
	dfwd->setClipboardValueByStr("linha", "cnpj", STR2CNPJ(part->cnpj))
	dfwd->setClipboardValueByStr("linha", "ie", STR2IE(part->ie))
	dfwd->setClipboardValueByStr("linha", "uf", MUNICIPIO2SIGLA(part->municip))
	dfwd->setClipboardValueByStr("linha", "municip", codMunicipio2Nome(part->municip))
	dfwd->setClipboardValueByStrW("linha", "razao", left(part->nome, 64))
	
	dfwd->paste("linha")
	
	adicionarDocRelatorioItemAnal(doc->situacao, doc->itemAnalListHead)
	
	nroRegistrosRel += 1
	
end sub

''''''''
sub Efd.adicionarDocRelatorioSaidas(doc as TECFReducaoZ ptr)

	var equip = doc->equipECF

	dfwd->setClipboardValueByStr("linha", "demi", YyyyMmDd2DatetimeBR(doc->dataMov))
	dfwd->setClipboardValueByStr("linha", "nrini", doc->numIni)
	dfwd->setClipboardValueByStr("linha", "nrfim", doc->numFim)
	dfwd->setClipboardValueByStr("linha", "ncaixa", equip->numCaixa)
	dfwd->setClipboardValueByStr("linha", "ecf", equip->numSerie)
	dfwd->setClipboardValueByStr("linha", "md", iif(equip->modelo = &h2D, "2D", str(equip->modelo)))
	dfwd->setClipboardValueByStr("linha", "sr", "")
	dfwd->setClipboardValueByStr("linha", "sit", "00")
	
	dfwd->paste("linha")
	
	adicionarDocRelatorioItemAnal(REGULAR, doc->itemAnalListHead)
	
	nroRegistrosRel += 1
	
end sub

''''''''
sub Efd.finalizarRelatorio()

	if ultimoRelatorio = -1 then
		return
	end if
	
	dim as string bookmark
	select case ultimoRelatorio
	case REL_LRE, REL_LRS
		bookmark = "_header"
	case else
		bookmark = "ass"
	end select
	
	dfwd->setClipboardValueByStr(bookmark, "nome_ass", infAssinatura->assinante)
	dfwd->setClipboardValueByStr(bookmark, "cpf_ass", STR2CPF(infAssinatura->cpf))
	dfwd->setClipboardValueByStr(bookmark, "hash_ass", infAssinatura->hashDoArquivo)

	select case ultimoRelatorio
	case REL_LRE, REL_LRS
		
		if nroRegistrosRel = 0 then
			dfwd->paste("vazio")
		else
			dfwd->paste("resumo")
		
			dim as RelSomatorioLR totSoma
			
			var soma = cast(RelSomatorioLR ptr, relSomaLRList.head)
			do while soma <> null
				if ultimoRelatorio = REL_LRS then
					dfwd->setClipboardValueByStr("resumo_linha", "sit", format(cdbl(soma->situacao), "00"))
				end if
				
				dfwd->setClipboardValueByStr("resumo_linha", "cst", format(soma->cst,"000"))
				dfwd->setClipboardValueByStr("resumo_linha", "cfop", soma->cfop)
				dfwd->setClipboardValueByStr("resumo_linha", "aliq", DBL2MONEYBR(soma->aliq))
				dfwd->setClipboardValueByStr("resumo_linha", "valop", DBL2MONEYBR(soma->valorOp))
				dfwd->setClipboardValueByStr("resumo_linha", "bc", DBL2MONEYBR(soma->bc))
				dfwd->setClipboardValueByStr("resumo_linha", "icms", DBL2MONEYBR(soma->icms))
				dfwd->setClipboardValueByStr("resumo_linha", "bcst", DBL2MONEYBR(soma->bcST))
				dfwd->setClipboardValueByStr("resumo_linha", "icmsst", DBL2MONEYBR(soma->ICMSST))
				dfwd->setClipboardValueByStr("resumo_linha", "ipi", DBL2MONEYBR(soma->ipi))
				
				dfwd->paste("resumo_linha")
				
				totSoma.valorOp += soma->valorOp
				totSoma.bc += soma->bc
				totSoma.icms += soma->icms
				totSoma.bcST += soma->bcST
				totSoma.ICMSST += soma->ICMSST
				totSoma.ipi += soma->ipi
				
				soma = relSomaLRList.next_(soma)
			loop
			
			dfwd->paste("resumo_sep")
			
			dfwd->setClipboardValueByStr("resumo_total", "totvalop", DBL2MONEYBR(totSoma.valorOp))
			dfwd->setClipboardValueByStr("resumo_total", "totbc", DBL2MONEYBR(totSoma.bc))
			dfwd->setClipboardValueByStr("resumo_total", "toticms", DBL2MONEYBR(totSoma.icms))
			dfwd->setClipboardValueByStr("resumo_total", "totbcst", DBL2MONEYBR(totSoma.bcST))
			dfwd->setClipboardValueByStr("resumo_total", "toticmsst", DBL2MONEYBR(totSoma.ICMSST))
			dfwd->setClipboardValueByStr("resumo_total", "totipi", DBL2MONEYBR(totSoma.ipi))
			
			dfwd->paste("resumo_total")
		end if
		
		relSomaLRDict.end_()
		relSomaLRList.end_()
	case else
		dfwd->paste("ass")
	end select
	
	dfwd->save(DdMmYyyy2Yyyy_Mm(regListHead->mestre.dataIni) + "_" + ultimoRelatorioSufixo + ".docx")
	
	dfwd->close()
	
	ultimoRelatorio = -1
	nroRegistrosRel = 0

end sub
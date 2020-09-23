#include once "efd.bi"
#include once "Dict.bi"
#include once "vbcompat.bi"
#include once "DB.bi"
#include once "trycatch.bi"

const PAGE_LEFT = 30
const PAGE_RIGHT = 813
const PAGE_TOP = 514
const PAGE_BOTTOM = 441.9
const ROW_SPACE_BEFORE = 3
const STROKE_WIDTH = 0.5
const ROW_HEIGHT = STROKE_WIDTH + 9.5 + STROKE_WIDTH + 0.5 	'' espaço anterior, linha superior, conteúdo, linha inferior, espaço posterior
const ANAL_HEIGHT = STROKE_WIDTH + 9.5 						'' linha superior, conteúdo, linha inferior
const LRE_RESUMO_TITLE_HEIGHT = 9
const LRE_RESUMO_HEADER_HEIGHT = 10
const LRE_RESUMO_ROW_HEIGHT = 10.0
const LRS_RESUMO_TITLE_HEIGHT = 9.0
const LRS_RESUMO_HEADER_HEIGHT = 9.0
const LRS_RESUMO_ROW_HEIGHT = 12.0

#macro list_add_ANAL(__doc, __sit)
	var anal = __doc.itemAnalListHead
	do while anal <> null
		var height = ANAL_HEIGHT
		if relYPos + height > PAGE_BOTTOM then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList.add())
		lin->tipo = REL_LIN_DF_ITEM_ANAL
		lin->anal.item = anal
		lin->anal.sit = __sit
		relYPos += height
		relNroLinhas += 1
		anal = anal->next_
	loop
#endmacro

#macro list_add_DF_ENTRADA(__doc, __part)
	scope
		var height = iif(relNroLinhas > 0, ROW_SPACE_BEFORE, 0) + ROW_HEIGHT
		if relYPos + height > PAGE_BOTTOM then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList.add())
		lin->tipo = REL_LIN_DF_ENTRADA
		lin->highlight = false
		lin->df.doc = @__doc
		lin->df.part = __part
		relYPos += iif(relNroLinhas > 0, ROW_SPACE_BEFORE, 0) + ROW_HEIGHT
		relNroLinhas += 1
		nroRegistrosRel += 1
		list_add_ANAL(__doc, __doc.situacao)
	end scope
#endmacro

#macro list_add_DF_SAIDA(__doc, __part)
	scope
		var height = iif(relNroLinhas > 0, ROW_SPACE_BEFORE, 0) + ROW_HEIGHT
		if relYPos + height > PAGE_BOTTOM then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList.add())
		lin->tipo = REL_LIN_DF_SAIDA
		lin->highlight = false
		lin->df.doc = @__doc
		lin->df.part = __part
		relYPos += iif(relNroLinhas > 0, ROW_SPACE_BEFORE, 0) + ROW_HEIGHT
		relNroLinhas += 1
		nroRegistrosRel += 1
		list_add_ANAL(__doc, __doc.situacao)
	end scope
#endmacro

#macro list_add_REDZ(__doc)
	scope
		var height = iif(relNroLinhas > 0, ROW_SPACE_BEFORE, 0) + ROW_HEIGHT
		if relYPos + height > PAGE_BOTTOM then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList.add())
		lin->tipo = REL_LIN_DF_REDZ
		lin->highlight = false
		lin->redz.doc = @__doc
		relYPos += iif(relNroLinhas > 0, ROW_SPACE_BEFORE, 0) + ROW_HEIGHT
		relNroLinhas += 1
		nroRegistrosRel += 1
		list_add_ANAL(__doc, REGULAR)
	end scope
#endmacro

#macro list_add_SAT(__doc)
	scope
		var height = iif(relNroLinhas > 0, ROW_SPACE_BEFORE, 0) + ROW_HEIGHT
		if relYPos + height > PAGE_BOTTOM then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList.add())
		lin->tipo = REL_LIN_DF_SAT
		lin->highlight = false
		lin->sat.doc = @__doc
		relYPos += iif(relNroLinhas > 0, ROW_SPACE_BEFORE, 0) + ROW_HEIGHT
		relNroLinhas += 1
		nroRegistrosRel += 1
		list_add_ANAL(__doc, REGULAR)
	end scope
#endmacro

''''''''
sub Efd.gerarRelatorios(nomeArquivo as string, mostrarProgresso as ProgressoCB)
	
	if opcoes.somenteRessarcimentoST then
		print wstr(!"\tNão será possivel gerar relatórios porque só foram extraídos os registros com ressarcimento ST")
	end if
	
	mostrarProgresso(wstr(!"\tGerando relatórios"), 0)
	
	ultimoRelatorio = -1

	relLinhasList.init(cint(PAGE_BOTTOM / ROW_HEIGHT + 0.5), len(RelLinha))
	relPaginasList.init(1000, len(RelPagina))
	
	mostrarProgresso(null, .1)
	
	if not opcoes.pularLreAoGerarRelatorios then
		'' LRE
		iniciarRelatorio(REL_LRE, "entradas", "LRE")
		
		var reg = regListHead
		try
			do while reg <> null
				'para cada registro..
				select case as const reg->tipo
				'NF-e?
				case DOC_NF, DOC_NFSCT, DOC_NF_ELETRIC
					if reg->nf.operacao = ENTRADA then
						var part = cast( TParticipante ptr, participanteDict[reg->nf.idParticipante] )
						list_add_DF_ENTRADA(reg->nf, part)
					end if
				
				'CT-e?
				case DOC_CT
					if reg->ct.operacao = ENTRADA then
						var part = cast( TParticipante ptr, participanteDict[reg->ct.idParticipante] )
						list_add_DF_ENTRADA(reg->ct, part)
					end if

				case LUA_CUSTOM
					var luaFunc = cast(customLuaCb ptr, customLuaCbDict[reg->lua.tipo])->rel_entradas
					
					if luaFunc <> null then
						'lua_getglobal(lua, luaFunc)
						'lua_pushlightuserdata(lua, dfwd)
						'lua_rawgeti(lua, LUA_REGISTRYINDEX, reg->lua.table)
						'lua_call(lua, 2, 0)
					end if
				end select
				
				reg = reg->next_
			loop
		catch
			print !"\r\nErro ao tratar o registro de tipo (" & reg->tipo & !") carregado na linha (" & reg->linha & !")\r\n"
		endtry
		
		finalizarRelatorio()
	end if
	
	mostrarProgresso(null, .5)
		
	if not opcoes.pularLrsAoGerarRelatorios then
		'' LRS
		iniciarRelatorio(REL_LRS, "saidas", "LRS")
		
		var reg = regListHead
		try
			do while reg <> null
				'para cada registro..
				select case as const reg->tipo
				'NF-e?
				case DOC_NF, DOC_NFSCT, DOC_NF_ELETRIC
					if reg->nf.operacao = SAIDA then
						var part = cast( TParticipante ptr, participanteDict[reg->nf.idParticipante] )
						list_add_DF_SAIDA(reg->nf, part)
					end if

				'CT-e?
				case DOC_CT
					if reg->ct.operacao = SAIDA then
						var part = cast( TParticipante ptr, participanteDict[reg->ct.idParticipante] )
						list_add_DF_SAIDA(reg->ct, part)
					end if
					
				'ECF Redução Z?
				case ECF_REDUCAO_Z
					list_add_REDZ(reg->ecfRedZ)
				
				'SAT?
				case DOC_SAT
					list_add_SAT(reg->sat)
				
				case LUA_CUSTOM
					var luaFunc = cast(customLuaCb ptr, customLuaCbDict[reg->lua.tipo])->rel_saidas
					
					if luaFunc <> null then
						'lua_getglobal(lua, luaFunc)
						'lua_pushlightuserdata(lua, dfwd)
						'lua_rawgeti(lua, LUA_REGISTRYINDEX, reg->lua.table)
						'lua_call(lua, 2, 0)
					end if
				end select

				reg = reg->next_
			loop
		catch
			print !"\r\nErro ao tratar o registro de tipo (" & reg->tipo & !") carregado na linha (" & reg->linha & !")\r\n"
		endtry
		
		finalizarRelatorio()
	end if
	
	'' outros livros..
	mostrarProgresso(null, .9)
	
	var reg = regListHead
	try
		do while reg <> null
			'para cada registro..
			select case as const reg->tipo
			case APURACAO_ICMS_PERIODO
				if not opcoes.pularLRaicmsAoGerarRelatorios then
					gerarRelatorioApuracaoICMS(nomeArquivo, reg)
				end if

			case APURACAO_ICMS_ST_PERIODO
				if not opcoes.pularLRaicmsAoGerarRelatorios then
					gerarRelatorioApuracaoICMSST(nomeArquivo, reg)
				end if
				
			case LUA_CUSTOM
				var luaFunc = cast(customLuaCb ptr, customLuaCbDict[reg->lua.tipo])->rel_outros
				
				if luaFunc <> null then
					'lua_getglobal(lua, luaFunc)
					'lua_pushlightuserdata(lua, dfwd)
					'lua_rawgeti(lua, LUA_REGISTRYINDEX, reg->lua.table)
					'lua_call(lua, 2, 0)
				end if
			end select

			reg = reg->next_
		loop
	catch
		print !"\r\nErro ao tratar o registro de tipo (" & reg->tipo & !") carregado na linha (" & reg->linha & !")\r\n"
	endtry
	
	relPaginasList.end_()
	relLinhasList.end_()
	
	mostrarProgresso(null, 1)

end sub

''''''''
private function efd.criarPaginaRelatorio(emitir as boolean) as RelPagina ptr
	var pagina = cast(RelPagina ptr, relPaginasList.add())
	pagina->emitir = emitir

	if emitir then
		relPage = relTemplate->clonePage(0)
		pagina->page = relPage
	end if
	
	relNroLinhas = 0
	relYPos = 0
	relNroPaginas += 1
	
	return pagina
end function

''''''''
private sub efd.gerarPaginaRelatorio(lastPage as boolean)

	var gerarPagina = true
	
	if not lastPage then
		if opcoes.filtrarCnpj then
			gerarPagina = false
			var n = cast(RelLinha ptr, relLinhasList.head)
			do while n <> null
				dim as TParticipante ptr part = null
				select case as const n->tipo
				case REL_LIN_DF_ENTRADA
					part = n->df.part
				case REL_LIN_DF_SAIDA
					part = n->df.part
				end select
				
				if part <> null then
					if filtrarPorCnpj(part->cnpj) then
						gerarPagina = true
						if not opcoes.highlight then
							exit do
						end if
						n->highlight = true
					end if
				end if
				
				n = relLinhasList.next_(n)
			loop
		end if
		
		if gerarPagina andalso opcoes.filtrarChaves then
			gerarPagina = false
			var n = cast(RelLinha ptr, relLinhasList.head)
			do while n <> null
				dim as zstring ptr chave = null
				select case as const n->tipo
				case REL_LIN_DF_ENTRADA, _
					 REL_LIN_DF_SAIDA
					chave = @n->df.doc->chave
				case REL_LIN_DF_SAT
					chave = @n->sat.doc->chave
				end select
				
				if chave <> null then
					if filtrarPorChave(chave) then
						gerarPagina = true
						if not opcoes.highlight then
							exit do
						end if
						n->highlight = true
					end if
				end if
				
				n = relLinhasList.next_(n)
			loop
		end if
	end if

	criarPaginaRelatorio(gerarPagina)

	'' emitir header e footer
	var n = cast(RelLinha ptr, relLinhasList.head)
	do while n <> null
		
		if gerarPagina then
			select case as const n->tipo
			case REL_LIN_DF_ENTRADA
				adicionarDocRelatorioEntradas(n->df.doc, n->df.part, n->highlight)
			case REL_LIN_DF_SAIDA
				adicionarDocRelatorioSaidas(n->df.doc, n->df.part, n->highlight)
			case REL_LIN_DF_REDZ
				adicionarDocRelatorioSaidas(n->redz.doc, n->highlight)
			case REL_LIN_DF_SAT
				adicionarDocRelatorioSaidas(n->sat.doc, n->highlight)
			case REL_LIN_DF_ITEM_ANAL
				adicionarDocRelatorioItemAnal(n->anal.sit, n->anal.item)
			end select
		else
			select case as const n->tipo
			case REL_LIN_DF_ENTRADA, REL_LIN_DF_SAIDA, REL_LIN_DF_REDZ, REL_LIN_DF_SAT
				relYPos += ROW_SPACE_BEFORE + ROW_HEIGHT
			case REL_LIN_DF_ITEM_ANAL
				relYPos += ANAL_HEIGHT
				relatorioSomarLR(n->anal.sit, n->anal.item)
			end select
		end if
		
		var p = n
		n = relLinhasList.next_(n)
		relLinhasList.del(p)
	loop
	
	if not lastPage then
		relNroLinhas = 0
		relYPos = 0
	end if

end sub

''''''''
sub Efd.gerarRelatorioApuracaoICMS(nomeArquivo as string, reg as TRegistro ptr)

	iniciarRelatorio(REL_RAICMS, "apuracao_icms", "RAICMS")
	
	criarPaginaRelatorio(true)
	
	setNodeText(relPage, "NOME", regMestre->mestre.nome)
	setNodeText(relPage, "CNPJ", STR2CNPJ(regMestre->mestre.cnpj))
	setNodeText(relPage, "IE", regMestre->mestre.ie)
	setNodeText(relPage, "ESCRIT", YyyyMmDd2DatetimeBR(regMestre->mestre.dataIni) + " a " + YyyyMmDd2DatetimeBR(regMestre->mestre.dataFim))
	setNodeText(relPage, "APU", YyyyMmDd2DatetimeBR(reg->apuIcms.dataIni) + " a " + YyyyMmDd2DatetimeBR(reg->apuIcms.dataFim))
	
	setNodeText(relPage, "SAIDAS", DBL2MONEYBR(reg->apuIcms.totalDebitos))
	setNodeText(relPage, "AJUSTE_DEB", DBL2MONEYBR(reg->apuIcms.ajustesDebitos))
	setNodeText(relPage, "AJUSTE_DEB_IMP", DBL2MONEYBR(reg->apuIcms.totalAjusteDeb))
	setNodeText(relPage, "ESTORNO_CRED", DBL2MONEYBR(reg->apuIcms.estornosCredito))
	setNodeText(relPage, "CREDITO", DBL2MONEYBR(reg->apuIcms.totalCreditos))
	setNodeText(relPage, "AJUSTE_CRED", DBL2MONEYBR(reg->apuIcms.ajustesCreditos))
	setNodeText(relPage, "AJUSTE_CRED_IMP", DBL2MONEYBR(reg->apuIcms.totalAjusteCred))
	setNodeText(relPage, "ESTORNO_DEB", DBL2MONEYBR(reg->apuIcms.estornoDebitos))
	setNodeText(relPage, "CRED_ANTERIOR", DBL2MONEYBR(reg->apuIcms.saldoCredAnterior))
	setNodeText(relPage, "SALDO_DEV", DBL2MONEYBR(reg->apuIcms.saldoDevedorApurado))
	setNodeText(relPage, "DEDUCOES", DBL2MONEYBR(reg->apuIcms.totalDeducoes))
	setNodeText(relPage, "A_RECOLHER", DBL2MONEYBR(reg->apuIcms.icmsRecolher))
	setNodeText(relPage, "A_TRANSPORTAR", DBL2MONEYBR(reg->apuIcms.saldoCredTransportar))
	setNodeText(relPage, "EXTRA_APU", DBL2MONEYBR(reg->apuIcms.debExtraApuracao))

	finalizarRelatorio()
	
end sub

''''''''
sub Efd.gerarRelatorioApuracaoICMSST(nomeArquivo as string, reg as TRegistro ptr)

	iniciarRelatorio(REL_RAICMSST, "apuracao_icms_st", "RAICMSST_" + reg->apuIcmsST.UF)

	criarPaginaRelatorio(true)
	
	setNodeText(relPage, "NOME", regMestre->mestre.nome)
	setNodeText(relPage, "CNPJ", STR2CNPJ(regMestre->mestre.cnpj))
	setNodeText(relPage, "IE", regMestre->mestre.ie)
	setNodeText(relPage, "ESCRIT", YyyyMmDd2DatetimeBR(regMestre->mestre.dataIni) + " a " + YyyyMmDd2DatetimeBR(regMestre->mestre.dataFim))
	setNodeText(relPage, "APU", YyyyMmDd2DatetimeBR(reg->apuIcmsST.dataIni) + " a " + YyyyMmDd2DatetimeBR(reg->apuIcmsST.dataFim))
	setNodeText(relPage, "UF", reg->apuIcmsST.UF)
	setNodeText(relPage, "MOV", iif(reg->apuIcmsST.mov, "1 - COM", "0 - SEM"))
	
	setNodeText(relPage, "SALDO_CRED", DBL2MONEYBR(reg->apuIcmsST.saldoCredAnterior))
	setNodeText(relPage, "DEVOLUCOES", DBL2MONEYBR(reg->apuIcmsST.devolMercadorias))
	setNodeText(relPage, "RESSARCIMENTOS", DBL2MONEYBR(reg->apuIcmsST.totalRessarciment))
	setNodeText(relPage, "OUTROS_CRED", DBL2MONEYBR(reg->apuIcmsST.totalOutrosCred))
	setNodeText(relPage, "AJUSTE_CRED", DBL2MONEYBR(reg->apuIcmsST.ajusteCred))
	setNodeText(relPage, "ICMS_ST", DBL2MONEYBR(reg->apuIcmsST.totalRetencao))
	setNodeText(relPage, "OUTROS_DEB", DBL2MONEYBR(reg->apuIcmsST.totalOutrosDeb))
	setNodeText(relPage, "AJUSTE_DEB", DBL2MONEYBR(reg->apuIcmsST.ajusteDeb))
	setNodeText(relPage, "SALDO_DEV", DBL2MONEYBR(reg->apuIcmsST.saldoAntesDed))
	setNodeText(relPage, "DEDUCOES", DBL2MONEYBR(reg->apuIcmsST.totalDeducoes))
	setNodeText(relPage, "A_RECOLHER", DBL2MONEYBR(reg->apuIcmsST.icmsRecolher))
	setNodeText(relPage, "A_TRANSPORTAR", DBL2MONEYBR(reg->apuIcmsST.saldoCredTransportar))
	setNodeText(relPage, "EXTRA_APU", DBL2MONEYBR(reg->apuIcmsST.debExtraApuracao))

	finalizarRelatorio()
	
end sub

''''''''
sub Efd.iniciarRelatorio(relatorio as TipoRelatorio, nomeRelatorio as string, sufixo as string)

	if ultimoRelatorio = relatorio then
		return
	end if
		
	finalizarRelatorio()
	
	ultimoRelatorioSufixo = sufixo
	ultimoRelatorio = relatorio
	nroRegistrosRel = 0
	
	relYPos = 0
	relNroLinhas = 0
	relNroPaginas = 0
	relPage = null
	
	select case relatorio
	case REL_LRE, REL_LRS
		relSomaLRList.init(10, len(RelSomatorioLR))
		relSomaLRDict.init(10)
	end select
	
	relTemplate = new PdfTemplate(baseTemplatesDir + nomeRelatorio + ".xml")
	relTemplate->load()
	
	var page = relTemplate->getPage(0)
	
	'' alterar header e footer
	var header = page->getNode("header")
	header->setAttrib("hidden", false)
	
	setNodeText(page, "NOME", regMestre->mestre.nome)
	setNodeText(page, "CNPJ", STR2CNPJ(regMestre->mestre.cnpj))
	setNodeText(page, "IE", regMestre->mestre.ie)
	
	select case relatorio
	case REL_LRE, REL_LRS
		setNodeText(page, "UF", MUNICIPIO2SIGLA(regMestre->mestre.municip))
		setNodeText(page, "MUNICIPIO", codMunicipio2Nome(regMestre->mestre.municip))
		setNodeText(page, "APU", YyyyMmDd2DatetimeBR(regMestre->mestre.dataIni) + " a " + YyyyMmDd2DatetimeBR(regMestre->mestre.dataFim))
	end select

	var footer = page->getNode("footer")
	footer->setAttrib("hidden", false)
	
	if infAssinatura <> null then
		setNodeText(page, "NOME_ASS", infAssinatura->assinante)
		setNodeText(page, "CPF_ASS", STR2CPF(infAssinatura->cpf))
		setNodeText(page, "HASH_ASS", infAssinatura->hashDoArquivo)
		if relatorio = REL_LRE then
			setNodeText(page, "NOME2_ASS", infAssinatura->assinante)
		end if
	end if

end sub

private function cmpFunc(key as any ptr, node as any ptr) as boolean
	function = *cast(zstring ptr, key) < cast(RelSomatorioLR ptr, node)->chave
end function

''''''''
function Efd.gerarLinhaDFe(highlight as boolean) as PdfTemplateNode ptr
	if relNroLinhas > 0 then
		relYPos += ROW_SPACE_BEFORE
	end if
	
	if highlight then
		var hl = new PdfTemplateHighlightNode(PAGE_LEFT, (PAGE_TOP-relYpos-ROW_HEIGHT), PAGE_RIGHT, (PAGE_TOP-relYPos), relPage)
	end if
	
	var row = relPage->getNode("row")
	var clone = row->clone(relPage, relPage)
	clone->setAttrib("hidden", false)
	clone->translateY(-relYPos)
	
	relYPos += ROW_HEIGHT
	relNroLinhas += 1
	
	return clone
end function

''''''''
function Efd.gerarLinhaAnal() as PdfTemplateNode ptr
	var anal = relPage->getNode("anal")
	var clone = anal->clone(relPage, relPage)
	clone->setAttrib("hidden", false)
	clone->translateY(-relYPos)
	
	relYPos += ANAL_HEIGHT
	relNroLinhas += 1

	return clone
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
sub Efd.setChildText(row as PdfTemplateNode ptr, id as string, value as string)
	var node = row->getChild(id)
	node->setAttrib("text", value)
end sub

''''''''
sub Efd.setChildText(row as PdfTemplateNode ptr, id as string, value as wstring ptr)
	var node = row->getChild(id)
	node->setAttrib("text", value)
end sub

''''''''
sub Efd.setNodeText(page as PdfTemplatePageNode ptr, id as string, value as string)
	var node = page->getNode(id)
	node->setAttrib("text", value)
end sub

''''''''
sub Efd.setNodeText(page as PdfTemplatePageNode ptr, id as string, value as wstring ptr)
	var node = page->getNode(id)
	node->setAttrib("text", value)
end sub

''''''''
sub Efd.adicionarDocRelatorioItemAnal(sit as TipoSituacao, anal as TDocItemAnal ptr)
	
	relatorioSomarLR(sit, anal)

	select case sit
	case REGULAR, EXTEMPORANEO
		var row = gerarLinhaAnal()
		setChildText(row, "CST", format(anal->cst,"000"))
		setChildText(row, "CFOP", str(anal->cfop))
		setChildText(row, "ALIQ", DBL2MONEYBR(anal->aliq))
		setChildText(row, "BCICMS", DBL2MONEYBR(anal->bc))
		setChildText(row, "ICMS", DBL2MONEYBR(anal->ICMS))
		setChildText(row, "BCICMSST", DBL2MONEYBR(anal->bcST))
		setChildText(row, "ICMSST", DBL2MONEYBR(anal->ICMSST))
		setChildText(row, "IPI", DBL2MONEYBR(anal->IPI))
		setChildText(row, "VALOP", DBL2MONEYBR(anal->valorOp))
		if ultimoRelatorio = REL_LRE then
			setChildText(row, "REDBC", DBL2MONEYBR(anal->redBC))
		end if
	end select

end sub

''''''''
static function Efd.luacb_efd_rel_addItemAnalitico cdecl(L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	lua_getglobal(L, "efd")
	var g_efd = cast(Efd ptr, lua_touserdata(L, -1))
	lua_pop(L, 1)
	
	if args = 2 then
		var sit = lua_tointeger(L, 1)
		
		dim as TDocItemAnal anal
		
		lua_pushnil(L)
		do while lua_next(L, -2) <> 0
			var value = lua_tonumber(L, -1)
			
			select case lcase(*lua_tostring(L, -2))
			case "cst"
				anal.cst = cint(value)
			case "cfop"
				anal.cfop = cint(value)
			case "aliq"
				anal.aliq = value
			case "valorop"
				anal.valorOp = value
			case "bc"
				anal.bc = value
			case "icms"
				anal.ICMS = value
			case "bcst"
				anal.bcST = value
			case "icmsst"
				anal.ICMSST = value
			case "redbc"
				anal.redBC = value
			case "ipi"
				anal.IPI = value
			end select
			
			lua_pop(L, 1)
		loop

		anal.next_ = null
		g_efd->adicionarDocRelatorioItemAnal(sit, @anal)
	end if
	
	function = 0
	
end function

''''''''
sub Efd.adicionarDocRelatorioSaidas(doc as TDocDF ptr, part as TParticipante ptr, highlight as boolean)
	var row = gerarLinhaDFe(highlight)
	
	if len(doc->dataEmi) > 0 then
		setChildText(row, "DEMI", YyyyMmDd2DatetimeBR(doc->dataEmi))
	end if
	if len(doc->dataEntSaida) > 0 then
		setChildText(row, "DSAIDA", YyyyMmDd2DatetimeBR(doc->dataEntSaida))
	end if
	setChildText(row, "NRINI", str(doc->numero))
	setChildText(row, "MD", str(doc->modelo))
	setChildText(row, "SR", doc->serie)
	setChildText(row, "SUB", doc->subserie)
	setChildText(row, "SIT", format(cdbl(doc->situacao), "00"))
	
	select case doc->situacao
	case REGULAR, EXTEMPORANEO
		if part <> null then
			setChildText(row, "CNPJDEST", iif(len(part->cpf) > 0, STR2CPF(part->cpf), STR2CNPJ(part->cnpj)))
			setChildText(row, "IEDEST", part->ie)
			setChildText(row, "UFDEST", MUNICIPIO2SIGLA(part->municip))
			setChildText(row, "MUNDEST", str(part->municip))
			setChildText(row, "RAZAODEST", left(part->nome, 36))
		end if
	end select
end sub

''''''''
sub Efd.adicionarDocRelatorioEntradas(doc as TDocDF ptr, part as TParticipante ptr, highlight as boolean)
	var row = gerarLinhaDFe(highlight)
	
	setChildText(row, "DEMI", YyyyMmDd2DatetimeBR(doc->dataEmi))
	setChildText(row, "DENT", YyyyMmDd2DatetimeBR(doc->dataEntSaida))
	setChildText(row, "NRO", str(doc->numero))
	setChildText(row, "MOD", str(doc->modelo))
	setChildText(row, "SER", doc->serie)
	setChildText(row, "SUBSER", doc->subserie)
	setChildText(row, "SIT", format(cdbl(doc->situacao), "00"))
	if part <> null then
		setChildText(row, "CNPJEMI", iif(len(part->cpf) > 0, STR2CPF(part->cpf), STR2CNPJ(part->cnpj)))
		setChildText(row, "IEEMI", part->ie)
		setChildText(row, "UFEMI", MUNICIPIO2SIGLA(part->municip))
		setChildText(row, "MUNEMI", codMunicipio2Nome(part->municip))
		setChildText(row, "RAZAOEMI", left(part->nome, 34))
	end if
end sub

''''''''
sub Efd.adicionarDocRelatorioSaidas(doc as TECFReducaoZ ptr, highlight as boolean)
	var equip = doc->equipECF

	var row = gerarLinhaDFe(highlight)
	
	setChildText(row, "DEMI", YyyyMmDd2DatetimeBR(doc->dataMov))
	setChildText(row, "NRINI", str(doc->numIni))
	setChildText(row, "NRINI", str(doc->numFim))
	setChildText(row, "NCAIXA", str(equip->numCaixa))
	setChildText(row, "ECF", equip->numSerie)
	setChildText(row, "MD", iif(equip->modelo = &h2D, "2D", str(equip->modelo)))
	setChildText(row, "SIT", "00")
end sub

''''''''
sub Efd.adicionarDocRelatorioSaidas(doc as TDocSAT ptr, highlight as boolean)
	var row = gerarLinhaDFe(highlight)
	
	setChildText(row, "DEMI", YyyyMmDd2DatetimeBR(doc->dataEmi))
	setChildText(row, "NRINI", str(doc->numero))
	setChildText(row, "ECF", doc->serieEquip)
	setChildText(row, "MD", str(doc->modelo))
	setChildText(row, "SIT", format(cdbl(doc->situacao), "00"))
end sub

''''''''
sub efd.gerarResumoRelatorioHeader()
	relYPos += ROW_SPACE_BEFORE
	
	var title = relPage->getNode("resumo-title")
	title->setAttrib("hidden", false)
	title->translateY(-relYPos)
	relYPos += iif(ultimoRelatorio = REL_LRS, LRS_RESUMO_TITLE_HEIGHT, LRE_RESUMO_TITLE_HEIGHT)

	var header = relPage->getNode("resumo-header")
	header->setAttrib("hidden", false)
	header->translateY(-relYPos)
	relYPos += iif(ultimoRelatorio = REL_LRS, LRS_RESUMO_HEADER_HEIGHT, LRE_RESUMO_HEADER_HEIGHT)
end sub

sub efd.gerarResumoRelatorio()
	var titleHeight = iif(ultimoRelatorio = REL_LRS, LRS_RESUMO_TITLE_HEIGHT, LRE_RESUMO_TITLE_HEIGHT)
	var headerHeight = iif(ultimoRelatorio = REL_LRS, LRS_RESUMO_HEADER_HEIGHT, LRE_RESUMO_HEADER_HEIGHT)
	var rowHeight = iif(ultimoRelatorio = REL_LRS, LRS_RESUMO_ROW_HEIGHT, LRE_RESUMO_ROW_HEIGHT)
	
	'' header
	if relPage = null orElse relYPos + ROW_SPACE_BEFORE + titleHeight + headerHeight + rowHeight > PAGE_BOTTOM then
		criarPaginaRelatorio(true)
	end if
	
	gerarResumoRelatorioHeader()

	'' tabela de totais
	dim as RelSomatorioLR totSoma
	
	var soma = cast(RelSomatorioLR ptr, relSomaLRList.head)
	do while soma <> null
		if relYPos + rowHeight > PAGE_BOTTOM then
			criarPaginaRelatorio(true)
			gerarResumoRelatorioHeader()
		end if
	
		var org = relPage->getNode("resumo-row")
		var row = org->clone(relPage, relPage)
		row->setAttrib("hidden", false)
		row->translateY(-relYPos)
		relYPos += rowHeight
	
		if ultimoRelatorio = REL_LRS then
			setChildText(row, "SIT", format(cdbl(soma->situacao), "00"))
		end if
		
		setChildText(row, "CST", format(soma->cst,"000"))
		setChildText(row, "CFOP", str(soma->cfop))
		setChildText(row, "ALIQ", DBL2MONEYBR(soma->aliq))
		setChildText(row, "OPER", DBL2MONEYBR(soma->valorOp))
		setChildText(row, "BCICMS", DBL2MONEYBR(soma->bc))
		setChildText(row, "ICMS", DBL2MONEYBR(soma->icms))
		setChildText(row, "BCICMSST", DBL2MONEYBR(soma->bcST))
		setChildText(row, "ICMSST", DBL2MONEYBR(soma->ICMSST))
		setChildText(row, "IPI", DBL2MONEYBR(soma->ipi))
		
		totSoma.valorOp += soma->valorOp
		totSoma.bc += soma->bc
		totSoma.icms += soma->icms
		totSoma.bcST += soma->bcST
		totSoma.ICMSST += soma->ICMSST
		totSoma.ipi += soma->ipi
		
		soma = relSomaLRList.next_(soma)
	loop
	
	'' totais
	if relYPos + ROW_SPACE_BEFORE + headerHeight > PAGE_BOTTOM then
		criarPaginaRelatorio(true)
		gerarResumoRelatorioHeader()
	end if
	
	relYPos += ROW_SPACE_BEFORE

	var total = relPage->getNode("resumo-total")
	total->setAttrib("hidden", false)
	total->translateY(-relYPos)
	relYPos += headerHeight
	
	setChildText(total, "OPERTOT", DBL2MONEYBR(totSoma.valorOp))
	setChildText(total, "BCICMSTOT", DBL2MONEYBR(totSoma.bc))
	setChildText(total, "ICMSTOT", DBL2MONEYBR(totSoma.icms))
	setChildText(total, "BCICMSSTTOT", DBL2MONEYBR(totSoma.bcST))
	setChildText(total, "ICMSSTTOT", DBL2MONEYBR(totSoma.ICMSST))
	setChildText(total, "IPITOT", DBL2MONEYBR(totSoma.ipi))
end sub

''''''''
sub Efd.finalizarRelatorio()

	if ultimoRelatorio = -1 then
		return
	end if
	
	var outf = new PdfDoc()
	
	select case ultimoRelatorio
	case REL_LRE, REL_LRS
		if nroRegistrosRel = 0 then
			criarPaginaRelatorio(true)
			var empty = relPage->getNode("empty")
			empty->setAttrib("hidden", false)
		else
			if relNroLinhas > 0 then
				gerarPaginaRelatorio(true)
			else
				criarPaginaRelatorio(true)
			end if
			gerarResumoRelatorio()
		end if

		relSomaLRDict.end_()
		relSomaLRList.end_()
	end select
	
	'' atribuir número de cada página
	var cnt = 1
	var pagina = cast(RelPagina ptr, relPaginasList.head)
	do while pagina <> null
		if pagina->emitir then
			var page = pagina->page
			var pg = page->getNode("PAGINA")
			if pg <> null then
				pg->setAttrib("text", wstr(cnt & "de " & relNroPaginas))
			end if
			page->emit(outf, cnt-1)
			delete page
		end if
		
		cnt += 1
		
		var last = pagina
		pagina = relPaginasList.next_(pagina)
		relPaginasList.del(last)
	loop
	
	'' salvar PDF
	outf->saveTo(DdMmYyyy2Yyyy_Mm(regMestre->mestre.dataIni) + "_" + ultimoRelatorioSufixo + ".pdf")
	delete outf
	
	delete relTemplate

	ultimoRelatorio = -1
	nroRegistrosRel = 0
	relYPos = 0
	relNroLinhas = 0
	relNroPaginas = 0
	relPage = null

end sub
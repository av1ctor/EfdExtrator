#include once "efd.bi"
#include once "Dict.bi"
#include once "vbcompat.bi"
#include once "DB.bi"
#include once "trycatch.bi"

const PAGE_VLIMIT = 441.9
const ROW_HEIGHT = 3 + 0.5 + (19*0.5) + 0.5 + 0.5 	'' espaço anterior, linha superior, conteúdo, linha inferior, espaço posterior
const ANAL_HEIGHT = 0.5 + (19*0.5) 					'' linha superior, conteúdo, linha inferior

#macro list_add_ANAL(__doc, __sit)
	var anal = __doc.itemAnalListHead
	do while anal <> null
		var height = ANAL_HEIGHT
		if relYPos + height > PAGE_VLIMIT then
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
		var height = ROW_HEIGHT + iif(relNroLinhas = 0, -3, 0)
		if relYPos + height > PAGE_VLIMIT then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList.add())
		lin->tipo = REL_LIN_DF_ENTRADA
		lin->df.doc = @__doc
		lin->df.part = __part
		relYPos += ROW_HEIGHT + iif(relNroLinhas = 0, -3, 0)
		relNroLinhas += 1
		list_add_ANAL(__doc, __doc.situacao)
	end scope
#endmacro

#macro list_add_DF_SAIDA(__doc, __part)
	scope
		var height = ROW_HEIGHT + iif(relNroLinhas = 0, -3, 0)
		if relYPos + height > PAGE_VLIMIT then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList.add())
		lin->tipo = REL_LIN_DF_SAIDA
		lin->df.doc = @__doc
		lin->df.part = __part
		relYPos += ROW_HEIGHT + iif(relNroLinhas = 0, -3, 0)
		relNroLinhas += 1
		list_add_ANAL(__doc, __doc.situacao)
	end scope
#endmacro

#macro list_add_REDZ(__doc)
	scope
		var height = ROW_HEIGHT + iif(relNroLinhas = 0, -3, 0)
		if relYPos + height > PAGE_VLIMIT then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList.add())
		lin->tipo = REL_LIN_DF_REDZ
		lin->redz.doc = @__doc
		relYPos += ROW_HEIGHT + iif(relNroLinhas = 0, -3, 0)
		relNroLinhas += 1
		list_add_ANAL(__doc, REGULAR)
	end scope
#endmacro

#macro list_add_SAT(__doc)
	scope
		var height = ROW_HEIGHT + iif(relNroLinhas = 0, -3, 0)
		if relYPos + height > PAGE_VLIMIT then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList.add())
		lin->tipo = REL_LIN_DF_SAT
		lin->sat.doc = @__doc
		relYPos += ROW_HEIGHT + iif(relNroLinhas = 0, -3, 0)
		relNroLinhas += 1
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

	relLinhasList.init(cint(PAGE_VLIMIT / ROW_HEIGHT + 0.5), len(RelLinha))
	relPaginasList.init(1000, len(RelPagina))
	
	'' NOTA: por limitação do DocxFactory, que só consegue trabalhar com um template por vez, 
	''		 precisamos processar entradas primeiro, depois saídas e por último os registros 
	''		 que são sequenciais (LRE e LRS vêm intercalados na EFD)
	
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
	var reg = regListHead
	try
		do while reg <> null
			'para cada registro..
			select case as const reg->tipo
			case APURACAO_ICMS_PERIODO
				gerarRelatorioApuracaoICMS(nomeArquivo, reg)

			case APURACAO_ICMS_ST_PERIODO
				gerarRelatorioApuracaoICMSST(nomeArquivo, reg)
				
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
private sub efd.gerarPaginaRelatorio(lastPage as boolean)

	var gerarPagina = true
	
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
					exit do
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
					exit do
				end if
			end if
			
			n = relLinhasList.next_(n)
		loop
	end if

	var pagina = cast(RelPagina ptr, relPaginasList.add())
	pagina->emitir = gerarPagina

	'' emitir header e footer
	if gerarPagina then
		relPage = relTemplate->clonePage(0)
		pagina->page = relPage
	end if
	
	var n = cast(RelLinha ptr, relLinhasList.head)
	relNroLinhas = 0
	relYPos = 0
	do while n <> null
		
		if gerarPagina then
			select case as const n->tipo
			case REL_LIN_DF_ENTRADA
				adicionarDocRelatorioEntradas(n->df.doc, n->df.part)
			case REL_LIN_DF_SAIDA
				adicionarDocRelatorioSaidas(n->df.doc, n->df.part)
			case REL_LIN_DF_REDZ
				adicionarDocRelatorioSaidas(n->redz.doc)
			case REL_LIN_DF_SAT
				adicionarDocRelatorioSaidas(n->sat.doc)
			case REL_LIN_DF_ITEM_ANAL
				adicionarDocRelatorioItemAnal(n->anal.sit, n->anal.item)
			end select
		else
			nroRegistrosRel += 1
			
			if n->tipo = REL_LIN_DF_ITEM_ANAL then
				relatorioSomarLR(n->anal.sit, n->anal.item)
			end if
		end if
		
		var p = n
		n = relLinhasList.next_(n)
		relLinhasList.del(p)
	loop
	
	relNroPaginas += 1
	relNroLinhas = 0
	relYPos = 0

end sub

''''''''
sub Efd.gerarRelatorioApuracaoICMS(nomeArquivo as string, reg as TRegistro ptr)

	iniciarRelatorio(REL_RAICMS, "apuracao_icms", "RAICMS")
	
	/'dfwd->setClipboardValueByStrW("grid", "nome", regMestre->mestre.nome)
	dfwd->setClipboardValueByStr("grid", "cnpj", STR2CNPJ(regMestre->mestre.cnpj))
	dfwd->setClipboardValueByStr("grid", "ie", STR2IE(regMestre->mestre.ie))
	dfwd->setClipboardValueByStr("grid", "escrit", YyyyMmDd2DatetimeBR(regMestre->mestre.dataIni) + " a " + YyyyMmDd2DatetimeBR(regMestre->mestre.dataFim))
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
	
	dfwd->paste("grid")'/

	finalizarRelatorio()
	
end sub

''''''''
sub Efd.gerarRelatorioApuracaoICMSST(nomeArquivo as string, reg as TRegistro ptr)

	iniciarRelatorio(REL_RAICMSST, "apuracao_icms_st", "RAICMSST_" + reg->apuIcmsST.UF)

	/'dfwd->setClipboardValueByStrW("grid", "nome", regMestre->mestre.nome)
	dfwd->setClipboardValueByStr("grid", "cnpj", STR2CNPJ(regMestre->mestre.cnpj))
	dfwd->setClipboardValueByStr("grid", "ie", STR2IE(regMestre->mestre.ie))
	dfwd->setClipboardValueByStr("grid", "escrit", YyyyMmDd2DatetimeBR(regMestre->mestre.dataIni) + " a " + YyyyMmDd2DatetimeBR(regMestre->mestre.dataFim))
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

	dfwd->paste("grid")'/

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
	
	var nome = page->getNode("NOME")
	nome->setAttrib("text", regMestre->mestre.nome)
	var cnpj = page->getNode("CNPJ")
	cnpj->setAttrib("text", STR2CNPJ(regMestre->mestre.cnpj))
	var ie = page->getNode("IE")
	ie->setAttrib("text", STR2IE(regMestre->mestre.ie))
	var uf = page->getNode("UF")
	uf->setAttrib("text", MUNICIPIO2SIGLA(regMestre->mestre.municip))

	select case relatorio
	case REL_LRE, REL_LRS
		var munic = page->getNode("MUNICIPIO")
		munic->setAttrib("text", codMunicipio2Nome(regMestre->mestre.municip))
		var apu = page->getNode("APU")
		apu->setAttrib("text", YyyyMmDd2DatetimeBR(regMestre->mestre.dataIni) + " a " + YyyyMmDd2DatetimeBR(regMestre->mestre.dataFim))
	end select

	var footer = page->getNode("footer")
	footer->setAttrib("hidden", false)
	
	if infAssinatura <> null then
		var nomeAss = page->getNode("NOME_ASS")
		nomeAss->setAttrib("text", infAssinatura->assinante)
		var cpfAss = page->getNode("CPF_ASS")
		cpfAss->setAttrib("text", STR2CPF(infAssinatura->cpf))
		var hashAss = page->getNode("HASH_ASS")
		hashAss->setAttrib("text", infAssinatura->hashDoArquivo)
	end if

end sub

private function cmpFunc(key as any ptr, node as any ptr) as boolean
	function = *cast(zstring ptr, key) < cast(RelSomatorioLR ptr, node)->chave
end function

''''''''
function Efd.gerarLinhaDFe() as PdfTemplateNode ptr
	if relNroLinhas > 0 then
		relYPos -= 3
	end if
	
	var row = relPage->getNode("row")
	var clone = row->clone(relPage, relPage)
	clone->setAttrib("hidden", false)
	clone->translateY(relYPos)
	
	relYPos -= 0.5 + (19*0.5) + 0.5 + 0.5
	relNroLinhas += 1
	
	return clone
end function

''''''''
function Efd.gerarLinhaAnal() as PdfTemplateNode ptr
	var anal = relPage->getNode("anal")
	var clone = anal->clone(relPage, relPage)
	clone->setAttrib("hidden", false)
	clone->translateY(relYPos)
	relYPos -= 0.5 + (19*0.5)
	
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
sub Efd.adicionarDocRelatorioItemAnal(sit as TipoSituacao, anal as TDocItemAnal ptr)
	
	relatorioSomarLR(sit, anal)

	select case sit
	case REGULAR, EXTEMPORANEO
	
		var row = gerarLinhaAnal()

		var cst = row->getChild("CST")
		cst->setAttrib("text", format(anal->cst,"000"))
		var cfop = row->getChild("")
		cfop->setAttrib("text", str(anal->cfop))
		var aliq = row->getChild("ALIQ")
		aliq->setAttrib("text", DBL2MONEYBR(anal->aliq))
		var bc = row->getChild("BC")
		bc->setAttrib("text", DBL2MONEYBR(anal->bc))
		var icms = row->getChild("ICMS")
		icms->setAttrib("text", DBL2MONEYBR(anal->ICMS))
		var bcst = row->getChild("BCST")
		bcst->setAttrib("text", DBL2MONEYBR(anal->bcST))
		var icmsSt = row->getChild("ICMSST")
		icmsSt->setAttrib("text", DBL2MONEYBR(anal->ICMSST))
		var ipi = row->getChild("IPI")
		ipi->setAttrib("text", DBL2MONEYBR(anal->IPI))
		var valop = row->getChild("VALOP")
		valop->setAttrib("text", DBL2MONEYBR(anal->valorOp))
		if ultimoRelatorio = REL_LRE then
			'var redbc = row->getChild("REDBC")
			'redbc->setAttrib("text", DBL2MONEYBR(anal->redBC))
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
sub Efd.adicionarDocRelatorioSaidas(doc as TDocDF ptr, part as TParticipante ptr)

	var row = gerarLinhaDFe()
	
	if len(doc->dataEmi) > 0 then
		var demi = row->getChild("DEMI")
		demi->setAttrib("text", YyyyMmDd2DatetimeBR(doc->dataEmi))
	end if
	if len(doc->dataEntSaida) > 0 then
		var dsaida = row->getChild("DSAIDA")
		dsaida->setAttrib("text", YyyyMmDd2DatetimeBR(doc->dataEntSaida))
	end if
	var nrini = row->getChild("NRINI")
	nrini->setAttrib("text", str(doc->numero))
	var md = row->getChild("MD")
	md->setAttrib("text", str(doc->modelo))
	var sr = row->getChild("SR")
	sr->setAttrib("text", doc->serie)
	var subsr = row->getChild("SUB")
	subsr->setAttrib("text", doc->subserie)
	var sit = row->getChild("SIT")
	sit->setAttrib("text", format(cdbl(doc->situacao), "00"))
	
	select case doc->situacao
	case REGULAR, EXTEMPORANEO
		if part <> null then
			var cnpjdest = row->getChild("CNPJDEST")
			cnpjdest->setAttrib("text", iif(len(part->cpf) > 0, STR2CPF(part->cpf), STR2CNPJ(part->cnpj)))
			var iedest = row->getChild("IEDEST")
			iedest->setAttrib("text", STR2IE(part->ie))
			var uf = row->getChild("UFDEST")
			uf->setAttrib("text", MUNICIPIO2SIGLA(part->municip))
			var mundest = row->getChild("MUNDEST")
			mundest->setAttrib("text", str(part->municip))
			var razaodest = row->getChild("RAZAODEST")
			razaodest->setAttrib("text", left(part->nome, 64))
		end if
	end select
	
	nroRegistrosRel += 1
	
end sub

''''''''
sub Efd.adicionarDocRelatorioEntradas(doc as TDocDF ptr, part as TParticipante ptr)

	/'
	dfwd->setClipboardValueByStr("linha", "demi", YyyyMmDd2DatetimeBR(doc->dataEmi))
	dfwd->setClipboardValueByStr("linha", "dent", YyyyMmDd2DatetimeBR(doc->dataEntSaida))
	dfwd->setClipboardValueByStr("linha", "nro", doc->numero)
	dfwd->setClipboardValueByStr("linha", "mod", doc->modelo)
	dfwd->setClipboardValueByStr("linha", "ser", doc->serie)
	dfwd->setClipboardValueByStr("linha", "subser", doc->subserie)
	dfwd->setClipboardValueByStr("linha", "sit", format(cdbl(doc->situacao), "00"))
	if part <> null then
		dfwd->setClipboardValueByStr("linha", "cnpj", iif(len(part->cpf) > 0, STR2CPF(part->cpf), STR2CNPJ(part->cnpj)))
		dfwd->setClipboardValueByStr("linha", "ie", STR2IE(part->ie))
		dfwd->setClipboardValueByStr("linha", "uf", MUNICIPIO2SIGLA(part->municip))
		dfwd->setClipboardValueByStr("linha", "municip", codMunicipio2Nome(part->municip))
		dfwd->setClipboardValueByStrW("linha", "razao", left(part->nome, 64))
	else
		dfwd->setClipboardValueByStr("linha", "cnpj", "")
		dfwd->setClipboardValueByStr("linha", "ie", "")
		dfwd->setClipboardValueByStr("linha", "uf", "")
		dfwd->setClipboardValueByStr("linha", "municip", "")
		dfwd->setClipboardValueByStr("linha", "razao", "")
	end if
	
	dfwd->paste("linha")'/
	
	nroRegistrosRel += 1
	
end sub

''''''''
sub Efd.adicionarDocRelatorioSaidas(doc as TECFReducaoZ ptr)

	var equip = doc->equipECF

	var row = gerarLinhaDFe()
	
	var demi = row->getChild("DEMI")
	demi->setAttrib("text", YyyyMmDd2DatetimeBR(doc->dataMov))
	var nrini = row->getChild("NRINI")
	nrini->setAttrib("text", str(doc->numIni))
	var nrfim = row->getChild("NRINI")
	nrfim->setAttrib("text", str(doc->numFim))
	var ncaixa = row->getChild("NCAIXA")
	ncaixa->setAttrib("text", str(equip->numCaixa))
	var ecf = row->getChild("ECF")
	ecf->setAttrib("text", equip->numSerie)
	var md = row->getChild("MD")
	md->setAttrib("text", iif(equip->modelo = &h2D, "2D", str(equip->modelo)))
	var sit = row->getChild("SIT")
	sit->setAttrib("text", "00")
	
	nroRegistrosRel += 1
	
end sub

''''''''
sub Efd.adicionarDocRelatorioSaidas(doc as TDocSAT ptr)

	var row = gerarLinhaDFe()
	
	var demi = row->getChild("DEMI")
	demi->setAttrib("text", YyyyMmDd2DatetimeBR(doc->dataEmi))
	var nrini = row->getChild("NRINI")
	nrini->setAttrib("text", str(doc->numero))
	var ecf = row->getChild("ECF")
	ecf->setAttrib("text", doc->serieEquip)
	var md = row->getChild("MD")
	md->setAttrib("text", str(doc->modelo))
	var sit = row->getChild("SIT")
	sit->setAttrib("text", format(cdbl(doc->situacao), "00"))
	
	nroRegistrosRel += 1
	
end sub

''''''''
sub Efd.finalizarRelatorio()

	if ultimoRelatorio = -1 then
		return
	end if
	
	gerarPaginaRelatorio(true)
	
	var outf = new PdfDoc()
	
	select case ultimoRelatorio
	case REL_LRE, REL_LRS
		
		if nroRegistrosRel = 0 then
			var empty = relPage->getNode("empty")
			empty->setAttrib("hidden", false)
		
		else
			if relPage = null then
				relPage = relTemplate->clonePage(0)
			end if

			var resumo = relPage->getNode("resumo")
			resumo->setAttrib("hidden", false)
		
			dim as RelSomatorioLR totSoma
			
			var soma = cast(RelSomatorioLR ptr, relSomaLRList.head)
			do while soma <> null
				/'
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
				'/
				
				totSoma.valorOp += soma->valorOp
				totSoma.bc += soma->bc
				totSoma.icms += soma->icms
				totSoma.bcST += soma->bcST
				totSoma.ICMSST += soma->ICMSST
				totSoma.ipi += soma->ipi
				
				soma = relSomaLRList.next_(soma)
			loop
			
			/'dfwd->paste("resumo_sep")
			
			dfwd->setClipboardValueByStr("resumo_total", "totvalop", DBL2MONEYBR(totSoma.valorOp))
			dfwd->setClipboardValueByStr("resumo_total", "totbc", DBL2MONEYBR(totSoma.bc))
			dfwd->setClipboardValueByStr("resumo_total", "toticms", DBL2MONEYBR(totSoma.icms))
			dfwd->setClipboardValueByStr("resumo_total", "totbcst", DBL2MONEYBR(totSoma.bcST))
			dfwd->setClipboardValueByStr("resumo_total", "toticmsst", DBL2MONEYBR(totSoma.ICMSST))
			dfwd->setClipboardValueByStr("resumo_total", "totipi", DBL2MONEYBR(totSoma.ipi))
			
			dfwd->paste("resumo_total")'/
		end if
		
		var cnt = 1
		var pagina = cast(RelPagina ptr, relPaginasList.head)
		do while pagina <> null
			if pagina->emitir /'andalso cnt < 50'/ then
				var page = pagina->page
				var pg = page->getNode("PAGINA")
				pg->setAttrib("text", wstr(cnt & "de " & relNroPaginas))
				page->emit(outf, cnt-1)
				delete page
			end if
			
			cnt += 1
			
			var last = pagina
			pagina = relPaginasList.next_(pagina)
			relPaginasList.del(last)
		loop
		
		relSomaLRDict.end_()
		relSomaLRList.end_()
	end select
	
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
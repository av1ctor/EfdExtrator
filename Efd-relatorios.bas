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
const ROW_HEIGHT_LG = ROW_HEIGHT + 5.5						'' linha larga (quando len(razãoSocial) > MAX_NAME_LEN)
const ANAL_HEIGHT = STROKE_WIDTH + 9.5 						'' linha superior, conteúdo, linha inferior
const LRE_MAX_NAME_LEN = 31.25
const LRS_MAX_NAME_LEN = 34.50
const LRE_RESUMO_TITLE_HEIGHT = 9
const LRE_RESUMO_HEADER_HEIGHT = 10
const LRE_RESUMO_ROW_HEIGHT = 10.0
const LRS_RESUMO_TITLE_HEIGHT = 9.0
const LRS_RESUMO_HEADER_HEIGHT = 9.0
const LRS_RESUMO_ROW_HEIGHT = 12.0

	dim shared charWidth(0 to 255) as single = {_
		0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,_
		0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,_
		0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,_
		0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,_
		0.50,0.25,0.25,1.00,1.00,1.00,1.00,0.25,_ 		'   ,!   ,"   ,#   ,$   ,%   ,&   ,'   
		0.25,0.25,0.25,0.25,0.25,0.25,0.25,0.50,_		'(  ,)   ,*   ,+   ,,   ,-   ,.   ,/   
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'0  ,1   ,2   ,3   ,4   ,5   ,6   ,7
		1.00,1.00,0.25,0.25,1.00,1.00,1.00,0.50,_		'8  ,9   ,:   ,;   ,<   ,=   ,>   ,?
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'@  ,A   ,B   ,C   ,D   ,E   ,F   ,G
		1.00,0.25,0.75,1.00,1.00,1.00,1.00,1.00,_		'H  ,I   ,J   ,K   ,L   ,M   ,N   ,O
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'P  ,Q   ,R   ,S   ,T   ,U   ,V   ,W
		1.00,1.00,1.00,0.25,0.25,0.25,0.50,0.50,_		'X  ,Y   ,Z   ,[   ,\   ,]   ,^   ,_
		0.25,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'`  ,a   ,b   ,c   ,d   ,e   ,f   ,g
		1.00,0.25,0.25,1.00,0.25,1.25,1.00,1.00,_		'h  ,i   ,j   ,k   ,l   ,m   ,n   ,o
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_		'p  ,q   ,r   ,s   ,t   ,u   ,v   ,w
		1.00,1.00,1.00,0.25,0.25,0.25,0.50,0.00,_		'x  ,y   ,z   ,{   ,|   ,}   ,~   , 
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00,_
		1.00,1.00,1.00,1.00,1.00,1.00,1.00,1.00 _
	}

private function calcLen(src as const zstring ptr) as double
	var lgt = 0.0
	for i as integer = 0 to len(*src) - 1
		lgt += charWidth(cast(ubyte ptr, src)[i])
	next
	return lgt
end function

private function substr(src as const zstring ptr, byref start as single, maxWidth as single) as string
	var res = ""
	var i = 0
	var width_ = 0.0
	do 
		var c = cast(ubyte ptr, src)[i]
		if c = 0 then
			exit do
		end if
		i += 1
		
		var cw = charWidth(c)
		if width_+cw > start+maxWidth then
			start = width_
			exit do
		end if
		
		if width_ >= start then
			res += chr(c)
		end if
		
		width_ += cw
	loop
	return res
end function

#macro list_add_ANAL(__doc, __sit)
	var anal = __doc.itemAnalListHead
	do while anal <> null
		if relYPos + ANAL_HEIGHT > PAGE_BOTTOM then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList->add())
		lin->tipo = REL_LIN_DF_ITEM_ANAL
		lin->anal.item = anal
		lin->anal.sit = __sit
		relYPos += ANAL_HEIGHT
		relNroLinhas += 1
		anal = anal->next_
	loop
#endmacro

#define calcHeight(lg) iif(relNroLinhas > 0, ROW_SPACE_BEFORE, 0) + iif(lg, ROW_HEIGHT_LG, ROW_HEIGHT)

#macro list_add_DF_ENTRADA(__doc, __part)
	scope
		var len_ = iif(part <> null, calcLen(part->nome), 0)
		var lg = len_ > cint(LRE_MAX_NAME_LEN + 0.5)
		if relYPos + calcHeight(lg) > PAGE_BOTTOM then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList->add())
		lin->tipo = REL_LIN_DF_ENTRADA
		lin->highlight = false
		lin->large = lg
		lin->df.doc = @__doc
		lin->df.part = __part
		relYPos += calcHeight(lg)
		relNroLinhas += 1
		nroRegistrosRel += 1
		list_add_ANAL(__doc, __doc.situacao)
	end scope
#endmacro

#macro list_add_DF_SAIDA(__doc, __part)
	scope
		var len_ = iif(part <> null, calcLen(part->nome), 0)
		var lg = len_ > cint(LRS_MAX_NAME_LEN + 0.5)
		if relYPos + calcHeight(lg) > PAGE_BOTTOM then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList->add())
		lin->tipo = REL_LIN_DF_SAIDA
		lin->highlight = false
		lin->large = lg
		lin->df.doc = @__doc
		lin->df.part = __part
		relYPos += calcHeight(lg)
		relNroLinhas += 1
		nroRegistrosRel += 1
		list_add_ANAL(__doc, __doc.situacao)
	end scope
#endmacro

#macro list_add_REDZ(__doc)
	scope
		if relYPos + calcHeight(false) > PAGE_BOTTOM then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList->add())
		lin->tipo = REL_LIN_DF_REDZ
		lin->highlight = false
		lin->large = false
		lin->redz.doc = @__doc
		relYPos += calcHeight(false)
		relNroLinhas += 1
		nroRegistrosRel += 1
		list_add_ANAL(__doc, REGULAR)
	end scope
#endmacro

#macro list_add_SAT(__doc)
	scope
		if relYPos + calcHeight(false) > PAGE_BOTTOM then
			gerarPaginaRelatorio()
		end if
		var lin = cast(RelLinha ptr, relLinhasList->add())
		lin->tipo = REL_LIN_DF_SAT
		lin->highlight = false
		lin->large = false
		lin->sat.doc = @__doc
		relYPos += calcHeight(false)
		relNroLinhas += 1
		nroRegistrosRel += 1
		list_add_ANAL(__doc, REGULAR)
	end scope
#endmacro

''''''''
sub Efd.gerarRelatorios(nomeArquivo as string)
	
	if opcoes.somenteRessarcimentoST then
		onError(!"\tNão será possivel gerar relatórios porque só foram extraídos os registros com ressarcimento ST")
	end if
	
	onProgress(wstr(!"\tGerando relatórios"), 0)
	
	ultimoRelatorio = -1

	relLinhasList = new TList(cint(PAGE_BOTTOM / ROW_HEIGHT + 0.5), len(RelLinha))
	relPaginasList = new Tlist(1000, len(RelPagina))
	
	onProgress(null, .1)
	
	if not opcoes.pularLre then
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
						var part = cast( TParticipante ptr, participanteDict->lookup(reg->nf.idParticipante) )
						list_add_DF_ENTRADA(reg->nf, part)
					end if
				
				'CT-e?
				case DOC_CT
					if reg->ct.operacao = ENTRADA then
						var part = cast( TParticipante ptr, participanteDict->lookup(reg->ct.idParticipante) )
						list_add_DF_ENTRADA(reg->ct, part)
					end if

				case LUA_CUSTOM
					var luaFunc = cast(customLuaCb ptr, customLuaCbDict->lookup(reg->lua.tipo))->rel_entradas
					
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
			onError(!"\r\nErro ao tratar o registro de tipo (" & reg->tipo & !") carregado na linha (" & reg->linha & !")\r\n")
		endtry
		
		finalizarRelatorio()
	end if
	
	onProgress(null, .5)
		
	if not opcoes.pularLrs then
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
						var part = cast( TParticipante ptr, participanteDict->lookup(reg->nf.idParticipante) )
						list_add_DF_SAIDA(reg->nf, part)
					end if

				'CT-e?
				case DOC_CT
					if reg->ct.operacao = SAIDA then
						var part = cast( TParticipante ptr, participanteDict->lookup(reg->ct.idParticipante) )
						list_add_DF_SAIDA(reg->ct, part)
					end if
					
				'ECF Redução Z?
				case ECF_REDUCAO_Z
					list_add_REDZ(reg->ecfRedZ)
				
				'SAT?
				case DOC_SAT
					list_add_SAT(reg->sat)
				
				case LUA_CUSTOM
					var luaFunc = cast(customLuaCb ptr, customLuaCbDict->lookup(reg->lua.tipo))->rel_saidas
					
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
			onError(!"\r\nErro ao tratar o registro de tipo (" & reg->tipo & !") carregado na linha (" & reg->linha & !")\r\n")
		endtry
		
		finalizarRelatorio()
	end if
	
	'' outros livros..
	onProgress(null, .9)
	
	var reg = regListHead
	try
		do while reg <> null
			'para cada registro..
			select case as const reg->tipo
			case APURACAO_ICMS_PERIODO
				if not opcoes.pularLraicms then
					gerarRelatorioApuracaoICMS(nomeArquivo, reg)
				end if

			case APURACAO_ICMS_ST_PERIODO
				if not opcoes.pularLraicms then
					gerarRelatorioApuracaoICMSST(nomeArquivo, reg)
				end if
				
			case LUA_CUSTOM
				var luaFunc = cast(customLuaCb ptr, customLuaCbDict->lookup(reg->lua.tipo))->rel_outros
				
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
		onError(!"\r\nErro ao tratar o registro de tipo (" & reg->tipo & !") carregado na linha (" & reg->linha & !")\r\n")
	endtry
	
	delete relPaginasList
	delete relLinhasList
	
	onProgress(null, 1)

end sub

''''''''
private function efd.criarPaginaRelatorio(emitir as boolean) as RelPagina ptr
	var pagina = cast(RelPagina ptr, relPaginasList->add())
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
private function efd.gerarPaginaRelatorio(lastPage as boolean) as boolean

	var gerarPagina = true
	
	if opcoes.filtrarCnpj then
		gerarPagina = false
		var n = cast(RelLinha ptr, relLinhasList->head)
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
			
			n = relLinhasList->next_(n)
		loop
	end if
	
	if gerarPagina andalso opcoes.filtrarChaves then
		gerarPagina = false
		var n = cast(RelLinha ptr, relLinhasList->head)
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
			
			n = relLinhasList->next_(n)
		loop
	end if

	criarPaginaRelatorio(gerarPagina)

	'' emitir header e footer
	var n = cast(RelLinha ptr, relLinhasList->head)
	do while n <> null
		
		if gerarPagina then
			select case as const n->tipo
			case REL_LIN_DF_ENTRADA
				adicionarDocRelatorioEntradas(n->df.doc, n->df.part, n->highlight, n->large)
			case REL_LIN_DF_SAIDA
				adicionarDocRelatorioSaidas(n->df.doc, n->df.part, n->highlight, n->large)
			case REL_LIN_DF_REDZ
				adicionarDocRelatorioSaidas(n->redz.doc, n->highlight)
			case REL_LIN_DF_SAT
				adicionarDocRelatorioSaidas(n->sat.doc, n->highlight)
			case REL_LIN_DF_ITEM_ANAL
				adicionarDocRelatorioItemAnal(n->anal.sit, n->anal.item)
			end select
		else
			select case as const n->tipo
			case REL_LIN_DF_ENTRADA, REL_LIN_DF_SAIDA
				relYPos += ROW_SPACE_BEFORE + iif(n->large, ROW_HEIGHT_LG, ROW_HEIGHT)
			case REL_LIN_DF_REDZ, REL_LIN_DF_SAT
				relYPos += ROW_SPACE_BEFORE + ROW_HEIGHT
			case REL_LIN_DF_ITEM_ANAL
				relYPos += ANAL_HEIGHT
				relatorioSomarLR(n->anal.sit, n->anal.item)
			end select
		end if
		
		var p = n
		n = relLinhasList->next_(n)
		relLinhasList->del(p)
	loop
	
	if not lastPage then
		relNroLinhas = 0
		relYPos = 0
	end if
	
	return gerarPagina

end function

''''''''
sub Efd.gerarRelatorioApuracaoICMS(nomeArquivo as string, reg as TRegistro ptr)

	iniciarRelatorio(REL_RAICMS, "apuracao_icms", "RAICMS")
	
	criarPaginaRelatorio(true)
	
	setNodeText(relPage, "NOME", regMestre->mestre.nome, true)
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
	
	setNodeText(relPage, "NOME", regMestre->mestre.nome, true)
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
	setNodeText(relPage, "AJUSTE_CRED", DBL2MONEYBR(reg->apuIcmsST.ajustesCreditos))
	setNodeText(relPage, "ICMS_ST", DBL2MONEYBR(reg->apuIcmsST.totalRetencao))
	setNodeText(relPage, "OUTROS_DEB", DBL2MONEYBR(reg->apuIcmsST.totalOutrosDeb))
	setNodeText(relPage, "AJUSTE_DEB", DBL2MONEYBR(reg->apuIcmsST.ajustesDebitos))
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
		relSomaLRList = new Tlist(10, len(RelSomatorioLR))
		relSomaLRDict = new TDict(10)
	end select
	
	relTemplate = new PdfTemplate(baseTemplatesDir + nomeRelatorio + ".xml")
	relTemplate->load()
	
	var page = relTemplate->getPage(0)
	
	'' alterar header e footer
	var header = page->getNode("header")
	header->setAttrib("hidden", false)
	
	setNodeText(page, "NOME", regMestre->mestre.nome, true)
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
		setNodeText(page, "NOME_ASS", infAssinatura->assinante, true)
		setNodeText(page, "CPF_ASS", STR2CPF(infAssinatura->cpf))
		setNodeText(page, "HASH_ASS", infAssinatura->hashDoArquivo)
		if relatorio = REL_LRE then
			setNodeText(page, "NOME2_ASS", infAssinatura->assinante, true)
		end if
	end if

end sub

private function cmpFunc(key as any ptr, node as any ptr) as boolean
	function = *cast(zstring ptr, key) < cast(RelSomatorioLR ptr, node)->chave
end function

''''''''
function Efd.gerarLinhaDFe(lg as boolean, highlight as boolean) as PdfTemplateNode ptr
	if relNroLinhas > 0 then
		relYPos += ROW_SPACE_BEFORE
	end if
	
	var height = iif(lg, ROW_HEIGHT_LG, ROW_HEIGHT)
	
	if highlight then
		var hl = new PdfTemplateHighlightNode(PAGE_LEFT, (PAGE_TOP-relYpos-height), PAGE_RIGHT, (PAGE_TOP-relYPos), relPage)
	end if
	
	var row = relPage->getNode(iif(lg, "row-lg", "row"))
	var clone = row->clone(relPage, relPage)
	clone->setAttrib("hidden", false)
	clone->translateY(-relYPos)
	
	relYPos += height
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
	
	var soma = cast(RelSomatorioLR ptr, relSomaLRDict->lookup(chave))
	if soma = null then
		soma = relSomaLRList->addOrdAsc(strptr(chave), @cmpFunc)
		soma->chave = chave
		soma->situacao = sit
		soma->cst = anal->cst
		soma->cfop = anal->cfop
		soma->aliq = anal->aliq
		relSomaLRDict->add(soma->chave, soma)
	end if
	
	soma->valorOp += anal->valorOp
	soma->bc += anal->bc
	soma->icms += anal->icms
	soma->bcST += anal->bcST
	soma->icmsST += anal->icmsST
	soma->ipi += anal->ipi
end sub

''''''''
sub Efd.setChildText(row as PdfTemplateNode ptr, id as string, value as wstring ptr)
	var node = row->getChild(id)
	node->setAttrib("text", value)
end sub

''''''''
sub Efd.setChildText(row as PdfTemplateNode ptr, id as string, value as string, convert as boolean)
	var node = row->getChild(id)
	if not convert then
		node->setAttrib("text", value)
	else
		var utf16le = latinToUtf16le(value)
		if utf16le <> null then
			node->setAttrib("text", utf16le)
			deallocate utf16le
		end if
	end if
end sub

''''''''
sub Efd.setNodeText(page as PdfTemplatePageNode ptr, id as string, value as wstring ptr)
	var node = page->getNode(id)
	node->setAttrib("text", value)
end sub

''''''''
sub Efd.setNodeText(page as PdfTemplatePageNode ptr, id as string, value as string, convert as boolean)
	var node = page->getNode(id)
	if not convert then
		node->setAttrib("text", value)
	else
		var utf16le = latinToUtf16le(value)
		if utf16le <> null then
			node->setAttrib("text", utf16le)
			deallocate utf16le
		end if
	end if
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
sub Efd.adicionarDocRelatorioSaidas(doc as TDocDF ptr, part as TParticipante ptr, highlight as boolean, lg as boolean)
	var row = gerarLinhaDFe(lg, highlight)
	
	if len(doc->dataEmi) > 0 then
		setChildText(row, iif(lg, "DEMI-LG", "DEMI"), YyyyMmDd2DatetimeBR(doc->dataEmi))
	end if
	if len(doc->dataEntSaida) > 0 then
		setChildText(row, iif(lg, "DSAIDA-LG", "DSAIDA"), YyyyMmDd2DatetimeBR(doc->dataEntSaida))
	end if
	setChildText(row, iif(lg, "NRINI-LG", "NRINI"), str(doc->numero))
	setChildText(row, iif(lg, "MD-LG", "MD"), str(doc->modelo))
	setChildText(row, iif(lg, "SR-LG", "SR"), doc->serie)
	setChildText(row, iif(lg, "SUB-LG", "SUB"), doc->subserie)
	setChildText(row, iif(lg, "SIT-LG", "SIT"), format(cdbl(doc->situacao), "00"))
	
	select case doc->situacao
	case REGULAR, EXTEMPORANEO
		if part <> null then
			setChildText(row, iif(lg, "CNPJDEST-LG", "CNPJDEST"), iif(len(part->cpf) > 0, STR2CPF(part->cpf), STR2CNPJ(part->cnpj)))
			setChildText(row, iif(lg, "IEDEST-LG", "IEDEST"), part->ie)
			setChildText(row, iif(lg, "UFDEST-LG", "UFDEST"), MUNICIPIO2SIGLA(part->municip))
			setChildText(row, iif(lg, "MUNDEST-LG", "MUNDEST"), str(part->municip))
			var start = 0.0!
			setChildText(row, iif(lg, "RAZAODEST-LG", "RAZAODEST"), substr(part->nome, start, LRS_MAX_NAME_LEN), true)
			if lg then
				setChildText(row, "RAZAODEST2-LG", substr(part->nome, start, LRS_MAX_NAME_LEN), true)
			end if
		end if
	end select
end sub

''''''''
sub Efd.adicionarDocRelatorioEntradas(doc as TDocDF ptr, part as TParticipante ptr, highlight as boolean, lg as boolean)
	var row = gerarLinhaDFe(lg, highlight)
	
	setChildText(row, iif(lg, "DEMI-LG", "DEMI"), YyyyMmDd2DatetimeBR(doc->dataEmi))
	setChildText(row, iif(lg, "DENT-LG", "DENT"), YyyyMmDd2DatetimeBR(doc->dataEntSaida))
	setChildText(row, iif(lg, "NRO-LG", "NRO"), str(doc->numero))
	setChildText(row, iif(lg, "MOD-LG", "MOD"), str(doc->modelo))
	setChildText(row, iif(lg, "SER-LG", "SER"), doc->serie)
	setChildText(row, iif(lg, "SUBSER-LG", "SUBSER"), doc->subserie)
	setChildText(row, iif(lg, "SIT-LG", "SIT"), format(cdbl(doc->situacao), "00"))
	if part <> null then
		setChildText(row, iif(lg, "CNPJEMI-LG", "CNPJEMI"), iif(len(part->cpf) > 0, STR2CPF(part->cpf), STR2CNPJ(part->cnpj)))
		setChildText(row, iif(lg, "IEEMI-LG", "IEEMI"), part->ie)
		setChildText(row, iif(lg, "UFEMI-LG", "UFEMI"), MUNICIPIO2SIGLA(part->municip))
		setChildText(row, iif(lg, "MUNEMI-LG", "MUNEMI"), codMunicipio2Nome(part->municip))
		var start = 0.0!
		setChildText(row, iif(lg, "RAZAOEMI-LG", "RAZAOEMI"), substr(part->nome, start, LRE_MAX_NAME_LEN), true)
		if lg then
			setChildText(row, "RAZAOEMI2-LG", substr(part->nome, start, LRS_MAX_NAME_LEN), true)
		end if
	end if
end sub

''''''''
sub Efd.adicionarDocRelatorioSaidas(doc as TECFReducaoZ ptr, highlight as boolean)
	var equip = doc->equipECF

	var row = gerarLinhaDFe(false, highlight)
	
	setChildText(row, "DEMI", YyyyMmDd2DatetimeBR(doc->dataMov))
	setChildText(row, "NRINI", str(doc->numIni))
	setChildText(row, "NRFIM", str(doc->numFim))
	setChildText(row, "NCAIXA", str(equip->numCaixa))
	setChildText(row, "ECF", equip->numSerie)
	setChildText(row, "MD", iif(equip->modelo = &h2D, "2D", str(equip->modelo)))
	setChildText(row, "SIT", "00")
end sub

''''''''
sub Efd.adicionarDocRelatorioSaidas(doc as TDocSAT ptr, highlight as boolean)
	var row = gerarLinhaDFe(false, highlight)
	
	setChildText(row, "DEMI", YyyyMmDd2DatetimeBR(doc->dataEmi))
	setChildText(row, "NRINI", str(doc->numero))
	setChildText(row, "ECF", doc->serieEquip)
	setChildText(row, "MD", str(doc->modelo))
	setChildText(row, "SIT", format(cdbl(doc->situacao), "00"))
end sub

''''''''
sub efd.gerarResumoRelatorioHeader(emitir as boolean)
	relYPos += ROW_SPACE_BEFORE
	
	if emitir then
		var title = relPage->getNode("resumo-title")
		title->setAttrib("hidden", false)
		title->translateY(-relYPos)
	end if
	relYPos += iif(ultimoRelatorio = REL_LRS, LRS_RESUMO_TITLE_HEIGHT, LRE_RESUMO_TITLE_HEIGHT)

	if emitir then
		var header = relPage->getNode("resumo-header")
		header->setAttrib("hidden", false)
		header->translateY(-relYPos)
	end if
	relYPos += iif(ultimoRelatorio = REL_LRS, LRS_RESUMO_HEADER_HEIGHT, LRE_RESUMO_HEADER_HEIGHT)
end sub

sub efd.gerarResumoRelatorio(emitir as boolean)
	var titleHeight = iif(ultimoRelatorio = REL_LRS, LRS_RESUMO_TITLE_HEIGHT, LRE_RESUMO_TITLE_HEIGHT)
	var headerHeight = iif(ultimoRelatorio = REL_LRS, LRS_RESUMO_HEADER_HEIGHT, LRE_RESUMO_HEADER_HEIGHT)
	var rowHeight = iif(ultimoRelatorio = REL_LRS, LRS_RESUMO_ROW_HEIGHT, LRE_RESUMO_ROW_HEIGHT)
	
	'' header
	if relPage = null orElse relYPos + ROW_SPACE_BEFORE + titleHeight + headerHeight + rowHeight > PAGE_BOTTOM then
		criarPaginaRelatorio(emitir)
	end if
	
	gerarResumoRelatorioHeader(emitir)

	'' tabela de totais
	dim as RelSomatorioLR totSoma
	
	var soma = cast(RelSomatorioLR ptr, relSomaLRList->head)
	do while soma <> null
		if relYPos + rowHeight > PAGE_BOTTOM then
			criarPaginaRelatorio(emitir)
			gerarResumoRelatorioHeader(emitir)
		end if
	
		if emitir then
			var org = relPage->getNode("resumo-row")
			var row = org->clone(relPage, relPage)
		
			row->setAttrib("hidden", false)
			row->translateY(-relYPos)
			
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
		end if
		relYPos += rowHeight
		
		totSoma.valorOp += soma->valorOp
		totSoma.bc += soma->bc
		totSoma.icms += soma->icms
		totSoma.bcST += soma->bcST
		totSoma.ICMSST += soma->ICMSST
		totSoma.ipi += soma->ipi
		
		soma = relSomaLRList->next_(soma)
	loop
	
	'' totais
	if relYPos + ROW_SPACE_BEFORE + headerHeight > PAGE_BOTTOM then
		criarPaginaRelatorio(emitir)
		gerarResumoRelatorioHeader(emitir)
	end if
	
	relYPos += ROW_SPACE_BEFORE

	if emitir then
		var total = relPage->getNode("resumo-total")
		total->setAttrib("hidden", false)
		total->translateY(-relYPos)

		setChildText(total, "OPERTOT", DBL2MONEYBR(totSoma.valorOp))
		setChildText(total, "BCICMSTOT", DBL2MONEYBR(totSoma.bc))
		setChildText(total, "ICMSTOT", DBL2MONEYBR(totSoma.icms))
		setChildText(total, "BCICMSSTTOT", DBL2MONEYBR(totSoma.bcST))
		setChildText(total, "ICMSSTTOT", DBL2MONEYBR(totSoma.ICMSST))
		setChildText(total, "IPITOT", DBL2MONEYBR(totSoma.ipi))
	end if
	relYPos += headerHeight
	
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
			var gerarResumo = true
			if relNroLinhas > 0 then
				var paginaGerada = gerarPaginaRelatorio(true)
				if not paginaGerada then
					gerarResumo = false
				end if
			else
				if opcoes.filtrarCnpj orelse opcoes.filtrarChaves then
					gerarResumo = false
				end if
				criarPaginaRelatorio(gerarResumo)
			end if
			gerarResumoRelatorio(gerarResumo)
		end if

		delete relSomaLRDict
		delete relSomaLRList
	end select
	
	'' atribuir número de cada página
	var cnt = 1
	var pagina = cast(RelPagina ptr, relPaginasList->head)
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
		pagina = relPaginasList->next_(pagina)
		relPaginasList->del(last)
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
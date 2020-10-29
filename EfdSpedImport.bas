#include once "EfdSpedImport.bi"
#include once "ssl_helper.bi"
#include once "trycatch.bi"

const ASSINATURA_P7K_HEADER = "SBRCAAEPDR"

''''''''
constructor EfdSpedImport(opcoes as OpcoesExtracao ptr)
	base(opcoes)
end constructor

''''''''
destructor EfdSpedImport()
end destructor

''''''''
function EfdSpedImport.withStmts( _
	lreInsertStmt as TDbStmt ptr, _
	itensNfLRInsertStmt as TDbStmt ptr, _
	lrsInsertStmt as TDbStmt ptr, _
	analInsertStmt as TDbStmt ptr, _
	ressarcStItensNfLRSInsertStmt as TDbStmt ptr, _
	itensIdInsertStmt as TDbStmt ptr, _
	mestreInsertStmt as TDbStmt ptr _
	) as EfdSpedImport ptr
	
	this.db_LREInsertStmt = lreInsertStmt
	this.db_itensNfLRInsertStmt = itensNfLRInsertStmt
	this.db_LRSInsertStmt = lrsInsertStmt
	this.db_analInsertStmt = analInsertStmt
	this.db_ressarcStItensNfLRSInsertStmt = ressarcStItensNfLRSInsertStmt
	this.db_itensIdInsertStmt = itensIdInsertStmt
	this.db_mestreInsertStmt = mestreInsertStmt
	
	return @this
end function

''''''''
private function yyyyMmDd2Days(d as const zstring ptr) as uinteger

	if d = null then
		return (1900 * 31*12) + 01
	end if
	
	var days = (cuint(d[0] - asc("0")) * 1000 + _
				cuint(d[1] - asc("0")) * 0100 + _
				cuint(d[2] - asc("0")) * 0010 + _
				cuint(d[3] - asc("0")) * 0001) * (31*12)
	
	days = days + _
			   ((cuint(d[4] - asc("0")) * 10 + _
				 cuint(d[5] - asc("0")) * 01) - 1) * 31

	days = days + _
			   (cuint(d[6] - asc("0")) * 10 + _
				cuint(d[7] - asc("0")) * 01) 
				
	function = days - (1900 * (31*12))

end function

''''''''
private function mergeLists(pSrc1 as TRegistro ptr, pSrc2 as TRegistro ptr) as TRegistro ptr
	dim as TRegistro ptr pDst = NULL
	dim as TRegistro ptr ptr ppDst = @pDst
    if pSrc1 = NULL then
        return pSrc2
	end if
    if pSrc2 = NULL then
        return pSrc1
	end if
    
	dim as zstring ptr dReg
	dim as uinteger nro
	dim as boolean isReg

	do while true
		select case as const pSrc1->tipo
		case DOC_NF, DOC_NFSCT, DOC_NF_ELETRIC
			isReg = ISREGULAR(pSrc1->nf.situacao)
			dReg = @pSrc1->nf.dataEntSaida
			nro = pSrc1->nf.numero
		case DOC_CT
			isReg = ISREGULAR(pSrc1->ct.situacao)
			dReg = @pSrc1->ct.dataEntSaida
			nro = pSrc1->ct.numero
		case DOC_NF_ITEM
			isReg = ISREGULAR(pSrc1->itemNF.documentoPai->situacao)
			dReg = @pSrc1->itemNF.documentoPai->dataEntSaida
			nro = pSrc1->itemNF.documentoPai->numero
		case ECF_REDUCAO_Z
			isReg = true
			dReg = @pSrc1->ecfRedZ.dataMov
			nro = pSrc1->ecfRedZ.numIni
		case DOC_SAT
			isReg = true
			dReg = @pSrc1->sat.dataEntSaida
			nro = pSrc1->sat.numero
		case else
			isReg = false
			dReg = null
			nro = 0
		end select
		
		var date1 = iif(isReg, yyyyMmDd2Days(dReg) shl 32, 0) + nro

		select case as const pSrc2->tipo
		case DOC_NF, DOC_NFSCT, DOC_NF_ELETRIC
			isReg = ISREGULAR(pSrc2->nf.situacao)
			dReg = @pSrc2->nf.dataEntSaida
			nro = pSrc2->nf.numero
		case DOC_CT
			isReg = ISREGULAR(pSrc2->ct.situacao)
			dReg = @pSrc2->ct.dataEntSaida
			nro = pSrc2->ct.numero
		case DOC_NF_ITEM
			isReg = ISREGULAR(pSrc2->itemNF.documentoPai->situacao)
			dReg = @pSrc2->itemNF.documentoPai->dataEntSaida
			nro = pSrc2->itemNF.documentoPai->numero
		case ECF_REDUCAO_Z
			isReg = true
			dReg = @pSrc2->ecfRedZ.dataMov
			nro = pSrc2->ecfRedZ.numIni
		case DOC_SAT
			isReg = true
			dReg = @pSrc2->sat.dataEntSaida
			nro = pSrc2->sat.numero
		case else
			isReg = false
			dReg = null
			nro = 0
		end select

		var date2 = iif(isReg, yyyyMmDd2Days(dReg) shl 32, 0) + nro

		if date2 < date1 then
			*ppDst = pSrc2
			ppDst = @pSrc2->next_
			pSrc2 = *ppDst
			if pSrc2 = NULL then
				*ppDst = pSrc1
				exit do
			end if
		else
			*ppDst = pSrc1
			ppDst = @pSrc1->next_
			pSrc1 = *ppDst
			if pSrc1 = NULL then
				*ppDst = pSrc2
				exit do
			end if
		end if
    loop
	
    function = pDst
end function

''''''''
private function ordenarRegistrosPorData(head as TRegistro ptr) as TRegistro ptr

	const NUMLISTS = 1000
	dim as TRegistro ptr aList(0 to NUMLISTS-1)
    
	if head = NULL then
        return NULL
	end if
    
	var n = head
	do while n <> NULL
        var nn = n->next_
        n->next_ = NULL
		var i = 0
        do while (i < NUMLISTS) and (aList(i) <> NULL)
            n = mergeLists(aList(i), n)
            aList(i) = NULL
			i += 1
        loop
        if i = NUMLISTS then
            i -= 1
		end if
        aList(i) = n
        n = nn
    loop
	
    n = NULL
    for i as integer = 0 to NUMLISTS-1
        n = mergeLists(aList(i), n)
	next
    
	function = n
	
end function

''''''''
function EfdSpedImport.lerTipo(bf as bfile, tipo as zstring ptr) as TipoRegistro

	if bf.peek1 <> asc("|") then
		onError("Erro: fora de sincronia na linha:" & nroLinha)
	else
		bf.char1 ' pular |
	end if
	
	*tipo = bf.char4
	var subtipo = valint(right(*tipo, 3))

	var tp = DESCONHECIDO
	
	select case as const tipo[0]
	case asc("0")
		select case subtipo
		case 150
			tp = PARTICIPANTE
		case 200
			tp = ITEM_ID
		case 300
			tp = BEM_CIAP
		case 305
			tp = BEM_CIAP_INFO
		case 450
			tp = INFO_COMPL
		case 460
			tp = OBS_LANCAMENTO
		case 500
			tp = CONTA_CONTAB
		case 600
			tp = CENTRO_CUSTO
		case 000
			tp = MESTRE
		end select
	case asc("C")
		select case subtipo
		case 100
			tp = DOC_NF
		case 110
			tp = DOC_NF_INFO
		case 170
			tp = DOC_NF_ITEM
		case 176
			tp = DOC_NF_ITEM_RESSARC_ST
		case 190
			tp = DOC_NF_ANAL
		case 195
			tp = DOC_NF_OBS
		case 197
			tp = DOC_NF_OBS_AJUSTE
		case 101
			tp = DOC_NF_DIFAL
		case 460
			tp = DOC_ECF
		case 470
			tp = DOC_ECF_ITEM
		case 490
			tp = DOC_ECF_ANAL
		case 400
			tp = EQUIP_ECF
		case 405
			tp = ECF_REDUCAO_Z
		case 500
			tp = DOC_NF_ELETRIC
		case 590
			tp = DOC_NF_ELETRIC_ANAL
		case 800
			tp = DOC_SAT
		case 850
			tp = DOC_SAT_ANAL
		end select
	case asc("D")
		select case subtipo
		case 100
			tp = DOC_CT
		case 190
			tp = DOC_CT_ANAL
		case 101
			tp = DOC_CT_DIFAL
		case 500
			tp = DOC_NFSCT
		case 590
			tp = DOC_NFSCT_ANAL
		end select
	case asc("E")	
		select case subtipo
		case 100
			tp = APURACAO_ICMS_PERIODO
		case 110
			tp = APURACAO_ICMS_PROPRIO
		case 111
			tp = APURACAO_ICMS_AJUSTE
		case 200
			tp = APURACAO_ICMS_ST_PERIODO
		case 210
			tp = APURACAO_ICMS_ST
		end select
	case asc("G")
		select case subtipo
		case 110
			tp = CIAP_TOTAL
		case 125
			tp = CIAP_ITEM
		case 130
			tp = CIAP_ITEM_DOC
		case 140
			tp = CIAP_ITEM_DOC_ITEM
		end select
	case asc("H")	
		select case subtipo
		case 005
			tp =  INVENTARIO_TOTAIS
		case 010
			tp =  INVENTARIO_ITEM
		end select
	case asc("K")
		select case subtipo
		case 100
			tp = ESTOQUE_PERIODO
		case 200
			tp = ESTOQUE_ITEM
		case 230
			tp = ESTOQUE_ORDEM_PROD
		end select
	case asc("9")
		select case subtipo
		case 999
			tp = FIM_DO_ARQUIVO
		end select
	end select
	
	if tp = DESCONHECIDO then
		if customLuaCbDict->lookup(*tipo) <> null then
			tp = LUA_CUSTOM
		end if
	end if
	
	function = tp

end function

''''''''
function EfdSpedImport.lerRegMestre(bf as bfile, reg as TRegistro ptr) as Boolean
   
	bf.char1		'pular |

	reg->mestre.versaoLayout= bf.varint
	reg->mestre.original 	= (bf.int1 = 0)
	bf.char1		'pular |
	reg->mestre.dataIni		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->mestre.dataFim		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->mestre.nome	   	= bf.varchar
	reg->mestre.cnpj	   	= bf.varchar
	reg->mestre.cpf	   		= bf.varint
	reg->mestre.uf			= bf.varchar
	reg->mestre.ie			= bf.varchar
	reg->mestre.municip		= bf.varint
	reg->mestre.im  		= bf.varchar
	reg->mestre.suframa  	= bf.varchar
	reg->mestre.perfil  	= bf.char1
	bf.char1		'pular |
	reg->mestre.atividade	= bf.int1
	bf.char1		'pular |

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function EfdSpedImport.lerRegParticipante(bf as bfile, reg as TRegistro ptr) as Boolean
   
	bf.char1		'pular |

	reg->part.id		= bf.varchar
	reg->part.nome		= bf.varchar
	reg->part.pais	   	= bf.varint
	reg->part.cnpj	   	= bf.varchar
	reg->part.cpf	   	= bf.varchar
	reg->part.ie		= bf.varchar
	reg->part.municip	= bf.varint
	reg->part.suframa  	= bf.varchar
	reg->part.ender	   	= bf.varchar
	reg->part.num		= bf.varchar
	reg->part.compl	   	= bf.varchar
	reg->part.bairro	= bf.varchar
   
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocNF(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->nf.operacao		= bf.int1
	bf.char1		'pular |
	reg->nf.emitente		= bf.int1
	bf.char1		'pular |
	reg->nf.idParticipante	= bf.varchar
	reg->nf.modelo			= bf.int2
	bf.char1		'pular |
	reg->nf.situacao		= bf.int2
	bf.char1		'pular |
	reg->nf.serie			= bf.varchar
	reg->nf.numero			= bf.varint
	reg->nf.chave			= bf.varchar
	reg->nf.dataEmi			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.valorTotal		= bf.vardbl
	reg->nf.pagamento		= bf.varint
	reg->nf.valorDesconto	= bf.vardbl
	reg->nf.valorAbatimento	= bf.vardbl
	reg->nf.valorMerc		= bf.vardbl
	reg->nf.frete			= bf.varint
	reg->nf.valorFrete		= bf.vardbl
	reg->nf.valorSeguro		= bf.vardbl
	reg->nf.valorAcessorias	= bf.vardbl
	reg->nf.bcICMS			= bf.vardbl
	reg->nf.ICMS			= bf.vardbl
	reg->nf.bcICMSST		= bf.vardbl
	reg->nf.ICMSST			= bf.vardbl
	reg->nf.IPI				= bf.vardbl
	reg->nf.PIS				= bf.vardbl
	reg->nf.COFINS			= bf.vardbl
	reg->nf.PISST			= bf.vardbl
	reg->nf.COFINSST		= bf.vardbl
	reg->nf.nroItens		= 0

	reg->nf.itemAnalListHead = null
	reg->nf.itemAnalListTail = null
	reg->nf.infoComplListHead = null
	reg->nf.infoComplListTail = null

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocNFInfo(bf as bfile, reg as TRegistro ptr, pai as TDocNF ptr) as Boolean

	bf.char1		'pular |

	reg->docInfoCompl.idCompl			= bf.varchar
	reg->docInfoCompl.extra				= bf.varchar
	reg->docInfoCompl.next_				= null
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocObs(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->docObs.idLanc			= bf.varchar
	reg->docObs.extra			= bf.varchar
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocObsAjuste(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->docObsAjuste.idAjuste		= bf.varchar
	reg->docObsAjuste.extra			= bf.varchar
	reg->docObsAjuste.idItem		= bf.varchar
	reg->docObsAjuste.bcICMS		= bf.vardbl
	reg->docObsAjuste.aliqICMS		= bf.vardbl
	reg->docObsAjuste.icms			= bf.vardbl
	reg->docObsAjuste.outros		= bf.vardbl
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocNFItem(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNF ptr) as Boolean

	bf.char1		'pular |

	reg->itemNF.documentoPai	= documentoPai
   
	reg->itemNF.numItem			= bf.varint
	reg->itemNF.itemId			= bf.varchar
	reg->itemNF.descricao		= bf.varchar
	reg->itemNF.qtd				= bf.vardbl
	reg->itemNF.unidade			= bf.varchar
	reg->itemNF.valor			= bf.vardbl
	reg->itemNF.desconto		= bf.vardbl
	reg->itemNF.indMovFisica	= bf.varint
	reg->itemNF.cstICMS			= bf.varint
	reg->itemNF.cfop			= bf.varint
	reg->itemNF.codNatureza		= bf.varchar
	reg->itemNF.bcICMS			= bf.vardbl
	reg->itemNF.aliqICMS		= bf.vardbl
	reg->itemNF.ICMS			= bf.vardbl
	reg->itemNF.bcICMSST		= bf.vardbl
	reg->itemNF.aliqICMSST		= bf.vardbl
	reg->itemNF.ICMSST			= bf.vardbl
	reg->itemNF.indApuracao		= bf.varint
	reg->itemNF.cstIPI			= bf.varint
	reg->itemNF.codEnqIPI		= bf.varchar
	reg->itemNF.bcIPI			= bf.vardbl
	reg->itemNF.aliqIPI			= bf.vardbl
	reg->itemNF.IPI				= bf.vardbl
	reg->itemNF.cstPIS			= bf.varint
	reg->itemNF.bcPIS			= bf.vardbl
	reg->itemNF.aliqPISPerc		= bf.vardbl
	reg->itemNF.qtdBcPIS		= bf.vardbl
	reg->itemNF.aliqPISMoed		= bf.vardbl
	reg->itemNF.PIS				= bf.vardbl
	reg->itemNF.cstCOFINS		= bf.varint
	reg->itemNF.bcCOFINS		= bf.vardbl
	reg->itemNF.aliqCOFINSPerc 	= bf.vardbl
	reg->itemNF.qtdBcCOFINS		= bf.vardbl
	reg->itemNF.aliqCOFINSMoed 	= bf.vardbl
	reg->itemNF.COFINS			= bf.vardbl
	bf.varchar					'' pular código da conta
	if regMestre->mestre.versaoLayout >= 013 then
		bf.vardbl				'' pular VL_ABAT_NT
	end if

	documentoPai->nroItens 		+= 1
	
	reg->itemNF.itemRessarcStListHead = null
	reg->itemNF.itemRessarcStListTail = null

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocNFItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= documentoPai
	reg->itemAnal.num		= documentoPai->nf.itemAnalCnt
	documentoPai->nf.itemAnalCnt += 1
	
	reg->itemAnal.cst		= bf.varint
	reg->itemAnal.cfop		= bf.varint
	reg->itemAnal.aliq		= bf.vardbl
	reg->itemAnal.valorOp	= bf.vardbl
	reg->itemAnal.bc		= bf.vardbl
	reg->itemAnal.ICMS		= bf.vardbl
	reg->itemAnal.bcST		= bf.vardbl
	reg->itemAnal.ICMSST	= bf.vardbl
	reg->itemAnal.redBC		= bf.vardbl
	reg->itemAnal.IPI		= bf.vardbl
	bf.varchar					'' pular código de observação

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function EfdSpedImport.lerRegDocNFItemRessarcSt(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNFItem ptr) as Boolean

	bf.char1		'pular |

	reg->itemRessarcSt.documentoPai	= documentoPai
	
	reg->itemRessarcSt.modeloUlt 			= bf.int2
	bf.char1		'pular |
	reg->itemRessarcSt.numeroUlt 			= bf.varint
	reg->itemRessarcSt.serieUlt  			= bf.varchar
	reg->itemRessarcSt.dataUlt				= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->itemRessarcSt.idParticipanteUlt	= bf.varchar
	reg->itemRessarcSt.qtdUlt				= bf.vardbl
	reg->itemRessarcSt.valorUlt				= bf.vardbl
	reg->itemRessarcSt.valorBcST			= bf.vardbl
	
	if bf.peek1 <> 13 then
		reg->itemRessarcSt.chaveNFeUlt		= bf.varchar
		reg->itemRessarcSt.numItemNFeUlt	= bf.varint
		reg->itemRessarcSt.bcIcmsUlt		= bf.vardbl
		reg->itemRessarcSt.aliqIcmsUlt		= bf.vardbl
		reg->itemRessarcSt.limiteBcIcmsUlt	= bf.vardbl
		reg->itemRessarcSt.icmsUlt			= bf.vardbl
		reg->itemRessarcSt.aliqIcmsStUlt	= bf.vardbl
		reg->itemRessarcSt.res				= bf.vardbl
		reg->itemRessarcSt.responsavelRet	= bf.int1
		bf.char1		'pular |
		reg->itemRessarcSt.motivo			= bf.int1
		bf.char1		'pular |
		reg->itemRessarcSt.chaveNFeRet		= bf.varchar
		reg->itemRessarcSt.idParticipanteRet= bf.varchar
		reg->itemRessarcSt.serieRet			= bf.varchar
		reg->itemRessarcSt.numeroRet		= bf.varint
		reg->itemRessarcSt.numItemNFeRet 	= bf.varint
		reg->itemRessarcSt.tipDocArrecadacao= bf.int1
		bf.char1		'pular |
		reg->itemRessarcSt.numDocArrecadacao= bf.varchar
	end if
   
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function EfdSpedImport.lerRegDocNFDifal(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNF ptr) as Boolean

	bf.char1		'pular |

	documentoPai->difal.fcp			= bf.vardbl
	documentoPai->difal.icmsDest	= bf.vardbl
	documentoPai->difal.icmsOrigem	= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocCT(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->ct.operacao		= bf.int1
	bf.char1		'pular |
	reg->ct.emitente		= bf.int1
	bf.char1		'pular |
	reg->ct.idParticipante	= bf.varchar
	reg->ct.modelo			= bf.int2
	bf.char1		'pular |
	reg->ct.situacao		= bf.int2
	bf.char1		'pular |
	reg->ct.serie			= bf.varchar
	bf.varchar		'pular sub-série
	reg->ct.numero			= bf.varint
	reg->ct.chave			= bf.varchar
	reg->ct.dataEmi			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ct.dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ct.tipoCTe			= bf.varint
	reg->ct.chaveRef		= bf.varchar
	reg->ct.valorTotal		= bf.vardbl
	reg->ct.valorDesconto	= bf.vardbl
	reg->ct.frete			= bf.varint
	reg->ct.valorServico	= bf.vardbl
	reg->ct.bcICMS			= bf.vardbl
	reg->ct.ICMS			= bf.vardbl
	reg->ct.valorNaoTributado = bf.vardbl
	reg->ct.codInfComplementar= bf.varchar
	bf.varchar		'pular código Conta Analitica
	
	'' códigos dos municípios de origem e de destino não aparecem em layouts antigos
	if bf.peek1 <> 13 and bf.peek1 <> 10 then 
		reg->ct.municipioOrigem	= bf.varint
		reg->ct.municipioDestino= bf.varint
	end if
	
	reg->ct.itemAnalListHead = null
	reg->ct.itemAnalListTail = null
	reg->ct.itemAnalCnt = 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocCTItemAnal(bf as bfile, reg as TRegistro ptr, docPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= docPai

	reg->itemAnal.cst			= bf.varint
	reg->itemAnal.cfop			= bf.varint
	reg->itemAnal.aliq			= bf.vardbl
	reg->itemAnal.valorOp		= bf.vardbl
	reg->itemAnal.bc			= bf.vardbl
	reg->itemAnal.ICMS			= bf.vardbl
	reg->itemAnal.redBc			= bf.vardbl
	bf.varchar					'' pular cod obs
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocCTDifal(bf as bfile, reg as TRegistro ptr, docPai as TDocCT ptr) as Boolean

	bf.char1		'pular |

	docPai->difal.fcp		= bf.vardbl
	docPai->difal.icmsDest	= bf.vardbl
	docPai->difal.icmsOrigem= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegEquipECF(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	var modelo = bf.varchar
	reg->equipECF.modelo	= iif(modelo = "2D", &h2D, valint(modelo))
	reg->equipECF.modeloEquip = bf.varchar
	reg->equipECF.numSerie 	= bf.varchar
	reg->equipECF.numCaixa	= bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocECF(bf as bfile, reg as TRegistro ptr, equipECF as TEquipECF ptr) as Boolean

	bf.char1		'pular |

	reg->ecf.equipECF		= equipECF
	reg->ecf.operacao		= SAIDA
	var modelo = bf.varchar
	reg->ecf.modelo			= iif(modelo = "2D", &h2D, valint(modelo))
	reg->ecf.situacao		= bf.int2
	bf.char1		'pular |
	reg->ecf.numero			= bf.varint
	reg->ecf.dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ecf.dataEntSaida	= reg->ecf.dataEmi
	reg->ecf.valorTotal		= bf.vardbl
	reg->ecf.PIS			= bf.vardbl
	reg->ecf.COFINS			= bf.vardbl
	reg->ecf.cpfCnpjAdquirente = bf.varchar
	reg->ecf.nomeAdquirente = bf.varchar
	reg->ecf.nroItens		= 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegECFReducaoZ(bf as bfile, reg as TRegistro ptr, equipECF as TEquipECF ptr) as Boolean

	bf.char1		'pular |

	reg->ecfRedZ.equipECF	= equipECF
	reg->ecfRedZ.dataMov	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ecfRedZ.cro		= bf.varint
	reg->ecfRedZ.crz		= bf.varint
	reg->ecfRedZ.numOrdem	= bf.varint
	reg->ecfRedZ.valorFinal	= bf.vardbl
	reg->ecfRedZ.valorBruto	= bf.vardbl

	reg->ecfRedZ.numIni		= 2^20
	reg->ecfRedZ.numFim		= -1
	reg->ecfRedZ.itemAnalListHead = null
	reg->ecfRedZ.itemAnalListTail = null

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocECFItem(bf as bfile, reg as TRegistro ptr, documentoPai as TDocECF ptr) as Boolean

	bf.char1		'pular |

	reg->itemECF.documentoPai	= documentoPai
   
	documentoPai->nroItens 		+= 1

	reg->itemECF.numItem		= documentoPai->nroItens
	reg->itemECF.itemId			= bf.varchar
	reg->itemECF.qtd			= bf.vardbl
	reg->itemECF.qtdCancelada	= bf.vardbl
	reg->itemECF.unidade		= bf.varchar
	reg->itemECF.valor			= bf.vardbl
	reg->itemECF.cstICMS		= bf.varint
	reg->itemECF.cfop			= bf.varint
	reg->itemECF.aliqICMS		= bf.vardbl
	reg->itemECF.PIS			= bf.vardbl
	reg->itemECF.COFINS			= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocECFItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= documentoPai
   
	reg->itemAnal.cst		= bf.varint
	reg->itemAnal.cfop		= bf.varint
	reg->itemAnal.aliq		= bf.vardbl
	reg->itemAnal.valorOp	= bf.vardbl
	reg->itemAnal.bc		= bf.vardbl
	reg->itemAnal.ICMS		= bf.vardbl
	bf.varchar					'' pular código de observação

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function EfdSpedImport.lerRegDocSAT(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->sat.operacao		= SAIDA
	reg->sat.modelo			= valint(bf.varchar)
	reg->sat.situacao		= bf.int2
	bf.char1		'pular |
	reg->sat.numero			= bf.varint
	reg->sat.dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->sat.valorTotal		= bf.vardbl
	reg->sat.PIS			= bf.vardbl
	reg->sat.COFINS			= bf.vardbl
	reg->sat.cpfCnpjAdquirente = bf.varchar
	reg->sat.serieEquip		= bf.varchar
	reg->sat.chave 			= bf.varchar
	reg->sat.descontos		= bf.vardbl
	reg->sat.valorMerc 		= bf.vardbl
	reg->sat.despesasAcess	= bf.vardbl
	reg->sat.icms			= bf.vardbl
	reg->sat.pisST			= bf.vardbl
	reg->sat.cofinsST		= bf.vardbl
	reg->sat.nroItens		= 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocSATItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= documentoPai
   
	reg->itemAnal.cst		= bf.varint
	reg->itemAnal.cfop		= bf.varint
	reg->itemAnal.aliq		= bf.vardbl
	reg->itemAnal.valorOp	= bf.vardbl
	reg->itemAnal.bc		= bf.vardbl
	reg->itemAnal.ICMS		= bf.vardbl
	bf.varchar					'' pular código de observação

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function


''''''''
function EfdSpedImport.lerRegDocNFSCT(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->nf.operacao		= bf.int1
	bf.char1		'pular |
	reg->nf.emitente		= bf.int1
	bf.char1		'pular |
	reg->nf.idParticipante	= bf.varchar
	reg->nf.modelo			= bf.int2
	bf.char1		'pular |
	reg->nf.situacao		= bf.int2
	bf.char1		'pular |
	reg->nf.serie			= bf.varchar
	reg->nf.subserie		= bf.varchar
	reg->nf.numero			= bf.varint
	reg->nf.dataEmi			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.valorTotal		= bf.vardbl
	reg->nf.valorDesconto	= bf.vardbl
	bf.vardbl		'pular valorServico
	bf.vardbl 		'pular valorServicoNT
	bf.vardbl 		'pular reg->nf.valorTerceiro
	bf.vardbl 		'pular reg->nf.valorDesp
	reg->nf.bcICMS			= bf.vardbl
	reg->nf.ICMS			= bf.vardbl
	bf.varchar		'pular cod_inf
	reg->nf.PIS				= bf.vardbl
	reg->nf.COFINS			= bf.vardbl
	bf.varchar		'pular cod_cta
	bf.varint		'pular tp_assinante
	reg->nf.nroItens		= 0

	reg->nf.itemAnalListHead = null
	reg->nf.itemAnalListTail = null
	reg->nf.itemAnalCnt = 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocNFSCTItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= documentoPai
   
	reg->itemAnal.cst		= bf.varint
	reg->itemAnal.cfop		= bf.varint
	reg->itemAnal.aliq		= bf.vardbl
	reg->itemAnal.valorOp	= bf.vardbl
	reg->itemAnal.bc		= bf.vardbl
	reg->itemAnal.ICMS		= bf.vardbl
	bf.vardbl		'pular VL_BC_ICMS_UF
	bf.vardbl		'pular VL_ICMS_UF
	reg->itemAnal.redBC		= bf.vardbl
	bf.varchar		'pular COD_OBS

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function EfdSpedImport.lerRegDocNFElet(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->nf.operacao		= bf.int1
	bf.char1		'pular |
	reg->nf.emitente		= bf.int1
	bf.char1		'pular |
	reg->nf.idParticipante	= bf.varchar
	reg->nf.modelo			= bf.int2
	bf.char1		'pular |
	reg->nf.situacao		= bf.int2
	bf.char1		'pular |
	reg->nf.serie			= bf.varchar
	reg->nf.subserie		= bf.varchar
	bf.varchar		'pular cod_cons
	reg->nf.numero			= bf.varint
	reg->nf.dataEmi			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.valorTotal		= bf.vardbl
	reg->nf.valorDesconto	= bf.vardbl
	bf.varchar		'pular vl_forn
	bf.varchar 		'pular vl_serv_nt
	bf.varchar		'pular vl_terc
	bf.varchar		'pular vl_da
	reg->nf.bcICMS			= bf.vardbl
	reg->nf.ICMS			= bf.vardbl
	reg->nf.bcICMSST		= bf.vardbl
	reg->nf.ICMSST			= bf.vardbl
	bf.varchar		'pular cod_inf
	reg->nf.PIS				= bf.vardbl
	reg->nf.COFINS			= bf.vardbl
	bf.varchar		'pular tp_ligacao
	bf.varchar		'pular cod_grupo_tensao
	if regMestre->mestre.versaoLayout >= 014 then
		reg->nf.chave		= bf.varchar		
		bf.varchar		'pular fin_doce
		bf.varchar		'pular chv_doce_ref
		bf.varchar		'pular ind_dest
		bf.varchar		'pular cod_mun_dest
		bf.varchar		'pular cod_cta
	end if
	reg->nf.nroItens		= 0

	reg->nf.itemAnalListHead = null
	reg->nf.itemAnalListTail = null
	reg->nf.itemAnalCnt = 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegDocNFEletItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= documentoPai
   
	reg->itemAnal.cst		= bf.varint
	reg->itemAnal.cfop		= bf.varint
	reg->itemAnal.aliq		= bf.vardbl
	reg->itemAnal.valorOp	= bf.vardbl
	reg->itemAnal.bc		= bf.vardbl
	reg->itemAnal.ICMS		= bf.vardbl
	reg->itemAnal.bcST		= bf.vardbl
	reg->itemAnal.ICMSST	= bf.vardbl
	reg->itemAnal.redBC		= bf.vardbl
	bf.varchar		'pular COD_OBS
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function EfdSpedImport.lerRegItemId(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemId.id			  	= bf.varchar
	reg->itemId.descricao	  	= bf.varchar
	reg->itemId.codBarra		= bf.varchar
	reg->itemId.codAnterior	  	= bf.varchar
	reg->itemId.unidInventario 	= bf.varchar
	reg->itemId.tipoItem		= bf.varint
	reg->itemId.ncm			  	= bf.varint
	reg->itemId.exIPI		  	= bf.varchar
	reg->itemId.codGenero	  	= bf.varint
	reg->itemId.codServico	  	= bf.varchar
	reg->itemId.aliqICMSInt	  	= bf.vardbl
	'CEST só é obrigatório a partir de 2017
	if bf.peek1 <> 13 and bf.peek1 <> 10 then 
	  reg->itemId.CEST		  	= bf.varint
	end if

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegBemCiap(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->bemCiap.id			  	= bf.varchar
	reg->bemCiap.tipoMerc		= bf.varint
	reg->bemCiap.descricao	  	= bf.varchar
	reg->bemCiap.principal		= bf.varchar
	reg->bemCiap.codAnal	  	= bf.varchar
	reg->bemCiap.parcelas		= bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegBemCiapInfo(bf as bfile, reg as TBemCiap ptr) as Boolean

	bf.char1		'pular |

	reg->codCusto		= bf.varchar
	reg->funcao	  		= bf.varchar
	reg->vidaUtil		= bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegObsLancamento(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->obsLanc.id				= bf.varchar
	reg->obsLanc.descricao	  	= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegContaContab(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->contaContab.dataInc		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->contaContab.codNat			= bf.varchar
	reg->contaContab.ind			= bf.varchar
	reg->contaContab.nivel			= bf.varint
	reg->contaContab.id			 	= bf.varchar
	reg->contaContab.descricao	  	= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegCentroCusto(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->centroCusto.dataInc		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->centroCusto.id			 	= bf.varchar
	reg->centroCusto.descricao	  	= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegInfoCompl(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->infoCompl.id				= bf.varchar
	reg->infoCompl.descricao	  	= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegApuIcmsPeriodo(bf as bfile, reg as TRegistro ptr) as Boolean

   bf.char1		'pular |

   reg->apuIcms.dataIni		  = ddMmYyyy2YyyyMmDd(bf.varchar)
   reg->apuIcms.dataFim		  = ddMmYyyy2YyyyMmDd(bf.varchar)

   'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

   function = true

end function

''''''''
function EfdSpedImport.lerRegApuIcmsProprio(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->apuIcms.totalDebitos			= bf.vardbl
	reg->apuIcms.ajustesDebitos			= bf.vardbl
	reg->apuIcms.totalAjusteDeb			= bf.vardbl
	reg->apuIcms.estornosCredito		= bf.vardbl
	reg->apuIcms.totalCreditos			= bf.vardbl
	reg->apuIcms.ajustesCreditos		= bf.vardbl
	reg->apuIcms.totalAjusteCred		= bf.vardbl
	reg->apuIcms.estornoDebitos			= bf.vardbl
	reg->apuIcms.saldoCredAnterior		= bf.vardbl
	reg->apuIcms.saldoDevedorApurado	= bf.vardbl
	reg->apuIcms.totalDeducoes			= bf.vardbl
	reg->apuIcms.icmsRecolher			= bf.vardbl
	reg->apuIcms.saldoCredTransportar	= bf.vardbl
	reg->apuIcms.debExtraApuracao		= bf.vardbl

	reg->apuIcms.ajustesListHead 		= null
	reg->apuIcms.ajustesListTail 		= null
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegApuIcmsAjuste(bf as bfile, reg as TRegistro ptr, pai as TApuracaoIcmsPeriodo ptr) as Boolean

	bf.char1		'pular |
	
	reg->apuIcmsAjust.codigo 	= bf.varchar
	reg->apuIcmsAjust.descricao = bf.varchar
	reg->apuIcmsAjust.valor 	= bf.vardbl
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegApuIcmsSTPeriodo(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->apuIcmsST.UF		 	 = bf.varchar
	reg->apuIcmsST.dataIni		 = ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->apuIcmsST.dataFim		 = ddMmYyyy2YyyyMmDd(bf.varchar)

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegApuIcmsST(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->apuIcmsST.mov						= bf.varint
	reg->apuIcmsST.saldoCredAnterior		= bf.vardbl
	reg->apuIcmsST.devolMercadorias			= bf.vardbl
	reg->apuIcmsST.totalRessarciment		= bf.vardbl
	reg->apuIcmsST.totalOutrosCred			= bf.vardbl
	reg->apuIcmsST.ajustesCreditos			= bf.vardbl
	reg->apuIcmsST.totalRetencao			= bf.vardbl
	reg->apuIcmsST.totalOutrosDeb			= bf.vardbl
	reg->apuIcmsST.ajustesDebitos			= bf.vardbl
	reg->apuIcmsST.saldoAntesDed			= bf.vardbl
	reg->apuIcmsST.totalDeducoes			= bf.vardbl
	reg->apuIcmsST.icmsRecolher				= bf.vardbl
	reg->apuIcmsST.saldoCredTransportar		= bf.vardbl
	reg->apuIcmsST.debExtraApuracao			= bf.vardbl

	reg->apuIcmsST.ajustesListHead 			= null
	reg->apuIcmsST.ajustesListTail 			= null
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegInventarioTotais(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->invTotais.dataInventario 	 = ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->invTotais.valorTotalEstoque = bf.vardbl
	reg->invTotais.motivoInventario	 = bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegInventarioItem(bf as bfile, reg as TRegistro ptr, inventarioPai as TInventarioTotais ptr) as Boolean

	bf.char1		'pular |

	reg->invItem.dataInventario 	= inventarioPai->dataInventario
	reg->invItem.itemId 	 		= bf.varchar
	reg->invItem.unidade 			= bf.varchar
	reg->invItem.qtd	 			= bf.vardbl
	reg->invItem.valorUnitario		= bf.vardbl
	reg->invItem.valorItem			= bf.vardbl
	reg->invItem.indPropriedade		= bf.varint
	reg->invItem.idParticipante		= bf.varchar
	reg->invItem.txtComplementar	= bf.varchar
	reg->invItem.codConta			= bf.varchar
	reg->invItem.valorItemIR		= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegCiapTotal(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->ciapTotal.dataIni 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ciapTotal.dataFim 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ciapTotal.saldoInicialICMS = bf.vardbl
	reg->ciapTotal.parcelasSoma 	= bf.vardbl
	reg->ciapTotal.valorTributExpSoma = bf.vardbl
	reg->ciapTotal.valorTotalSaidas = bf.vardbl
	reg->ciapTotal.indicePercSaidas = bf.vardbl
	reg->ciapTotal.valorIcmsAprop 	= bf.vardbl
	reg->ciapTotal.valorOutrosCred 	= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegCiapItem(bf as bfile, reg as TRegistro ptr, pai as TCiapTotal ptr) as Boolean

	bf.char1		'pular |

	reg->ciapItem.pai				= pai
	reg->ciapItem.bemId 	 		= bf.varchar
	reg->ciapItem.dataMov 			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ciapItem.tipoMov 			= bf.varchar
	reg->ciapItem.valorIcms	 		= bf.vardbl
	reg->ciapItem.valorIcmsSt		= bf.vardbl
	reg->ciapItem.valorIcmsFrete	= bf.vardbl
	reg->ciapItem.valorIcmsDifal	= bf.vardbl
	reg->ciapItem.parcela			= bf.varint
	reg->ciapItem.valorParcela		= bf.vardbl
	reg->ciapItem.docCnt			= 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegCiapItemDoc(bf as bfile, reg as TRegistro ptr, pai as TCiapItem ptr) as Boolean

	bf.char1		'pular |

	reg->ciapItemDoc.pai			= pai
	reg->ciapItemDoc.indEmi 		= bf.varint
	reg->ciapItemDoc.idParticipante = bf.varchar
	reg->ciapItemDoc.modelo			= bf.varint
	reg->ciapItemDoc.serie			= bf.varchar
	reg->ciapItemDoc.numero			= bf.varint
	reg->ciapItemDoc.chaveNFe		= bf.varchar
	reg->ciapItemDoc.dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	if bf.peek1 <> 13 andalso bf.peek1 <> 10 then 
		bf.varchar '' pular NUM_DA
	end if
	pai->docCnt += 1

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegCiapItemDocItem(bf as bfile, reg as TRegistro ptr, pai as TCiapItemDoc ptr) as Boolean

	bf.char1		'pular |

	reg->ciapItemDocItem.pai			= pai
	reg->ciapItemDocItem.num			= bf.varint
	reg->ciapItemDocItem.itemId 		= bf.varchar
	if bf.peek1 <> 13 andalso bf.peek1 <> 10 then 
		bf.vardbl 		'' pular QTDE
		bf.varchar 		'' pular UNID
		bf.vardbl 		'' pular VL_ICMS_OP
		bf.vardbl 		'' pular VL_ICMS_ST
		bf.vardbl 		'' pular VL_ICMS_FRT
		bf.vardbl 		'' pular VL_ICMS_DIF
	end if

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegEstoquePeriodo(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->estPeriod.dataIni 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->estPeriod.dataFim 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegEstoqueItem(bf as bfile, reg as TRegistro ptr, pai as TEstoquePeriodo ptr) as Boolean

	bf.char1		'pular |

	reg->estItem.pai				= pai
	bf.varchar		'pular DT_EST (é a mesma do DT_FIN do K100)
	reg->estItem.itemId 	 		= bf.varchar
	reg->estItem.qtd 				= bf.vardbl
	reg->estItem.tipoEst			= bf.varint
	reg->estItem.idParticipante		= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function EfdSpedImport.lerRegEstoqueOrdemProd(bf as bfile, reg as TRegistro ptr, pai as TEstoquePeriodo ptr) as Boolean

	bf.char1		'pular |

	reg->estOrdem.pai			= pai
	reg->estOrdem.dataIni 	 	= ddMmYyyy2YyyyMmDd(bf.varchar)
	var dtFim = bf.varchar
	reg->estOrdem.dataFim 	 	= iif(len(dtFim) > 0, ddMmYyyy2YyyyMmDd(dtFim), "99991231")
	reg->estOrdem.idOrdem		= bf.varchar
	reg->estOrdem.itemId 	 	= bf.varchar
	reg->estOrdem.qtd 			= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
private sub EfdSpedImport.lerAssinatura(bf as bfile)

	'' verificar header
	var header = bf.nchar(len(ASSINATURA_P7K_HEADER))
	if header <> ASSINATURA_P7K_HEADER then
		onError("Erro: header da assinatura P7K não reconhecido")
	end if
	
	var lgt = (bf.tamanho - bf.posicao) + 1
	
	redim this.assinaturaP7K_DER(0 to lgt-1)
	
	bf.ler(assinaturaP7K_DER(), lgt)

end sub

''''''''
function EfdSpedImport.lerRegistro(bf as bfile, reg as TRegistro ptr) as Boolean
	static as zstring * 4+1 tipo
	
	reg->tipo = lerTipo(bf, @tipo)
	reg->linha = nroLinha

	select case as const reg->tipo
	case DOC_NF
		if not lerRegDocNF(bf, reg) then
			return false
		end if
		
		ultimoReg = reg

	case DOC_NF_INFO
		if( ultimoReg <> null ) then
			if not lerRegDocNFInfo(bf, reg, @ultimoReg->nf) then
				return false
			end if
			
			var node = @reg->docInfoCompl
			var parent = @ultimoReg->nf
			
			if parent->infoComplListHead = null then
				parent->infoComplListHead = node
			else
				parent->infoComplListTail->next_ = node
			end if
			
			parent->infoComplListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
		
	case DOC_NF_ITEM
		if( ultimoReg <> null ) then
			if not lerRegDocNFItem(bf, reg, @ultimoReg->nf) then
				return false
			end if
			
			ultimoDocNFItem = @reg->itemNF
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_NF_ANAL
		if( ultimoReg <> null ) then
			if not lerRegDocNFItemAnal(bf, reg, ultimoReg) then
				return false
			end if
			
			var node = @reg->itemAnal
			var parent = @ultimoReg->nf
			
			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_NF_OBS
		if( ultimoReg <> null ) then
			if not lerRegDocObs(bf, reg) then
				return false
			end if
			
			ultimoDocObs = @reg->docObs
			var node = ultimoDocObs
			var parent = @ultimoReg->nf
			
			if parent->obsListHead = null then
				parent->obsListHead = node
			else
				parent->obsListTail->next_ = node
			end if
			
			parent->obsListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_NF_OBS_AJUSTE
		if( ultimoDocObs <> null ) then
			if not lerRegDocObsAjuste(bf, reg) then
				return false
			end if
			
			var node = @reg->docObsAjuste
			var parent = ultimoDocObs
			
			if parent->ajusteListHead = null then
				parent->ajusteListHead = node
			else
				parent->ajusteListTail->next_ = node
			end if
			
			parent->ajusteListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
		
	case DOC_NF_DIFAL
		if( ultimoReg <> null ) then
			if not lerRegDocNFDifal(bf, reg, @ultimoReg->nf) then
				return false
			end if
			
			reg->tipo = DESCONHECIDO			'' deletar registro, já que vamos reusar o registro pai
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
		
	case DOC_NF_ITEM_RESSARC_ST
		if( ultimoDocNFItem <> null ) then
			if not lerRegDocNFItemRessarcSt(bf, reg, ultimoDocNFItem) then
				return false
			end if
			
			var node = @reg->itemRessarcSt
			var parent = ultimoDocNFItem
			
			if parent->itemRessarcStListHead = null then
				parent->itemRessarcStListHead = node
			else
				parent->itemRessarcStListTail->next_ = node
			end if
			
			parent->itemRessarcStListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_CT
		if not lerRegDocCT(bf, reg) then
			return false
		end if

		ultimoReg = reg

	case DOC_CT_ANAL
		if( ultimoReg <> null ) then
			if not lerRegDocCTItemAnal(bf, reg, ultimoReg) then
				return false
			end if

			var node = @reg->itemAnal
			var parent = @ultimoReg->ct
			
			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
		
	case DOC_CT_DIFAL
		if( ultimoReg <> null ) then
			if not lerRegDocCTDifal(bf, reg, @reg->ct) then
				return false
			end if
			
			reg->tipo = DESCONHECIDO			'' deletar registro, já que vamos reusar o registro pai
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_ECF
		if( ultimoEquipECF <> null ) then
			if not lerRegDocECF(bf, reg, ultimoEquipECF) then
				return false
			end if

			ultimoReg = reg
			
			if ultimoECFRedZ->ecfRedZ.numIni > reg->ecf.numero then
				ultimoECFRedZ->ecfRedZ.numIni = reg->ecf.numero
			end if

			if ultimoECFRedZ->ecfRedZ.numFim < reg->ecf.numero then
				ultimoECFRedZ->ecfRedZ.numFim = reg->ecf.numero
			end if
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
		
	case ECF_REDUCAO_Z
		if( ultimoEquipECF <> null ) then
			if not lerRegECFReducaoZ(bf, reg, ultimoEquipECF) then
				return false
			end if

			ultimoECFRedZ = reg
		else
			pularLinha(bf)
			ultimoECFRedZ = null
			reg->tipo = DESCONHECIDO
		end if
		
	case DOC_ECF_ITEM
		if( ultimoReg <> null ) then
			if not lerRegDocECFItem(bf, reg, @ultimoReg->ecf) then
				return false
			end if
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_ECF_ANAL
		if( ultimoECFRedZ <> null ) then
			if not lerRegDocECFItemAnal(bf, reg, ultimoECFRedZ) then
				return false
			end if
			
			var node = @reg->itemAnal
			var parent = @ultimoECFRedZ->ecfRedZ
			
			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case EQUIP_ECF
		if not lerRegEquipECF(bf, reg) then
			return false
		end if
		
		ultimoEquipECF = @reg->equipECF

	case DOC_SAT
		if not lerRegDocSAT(bf, reg) then
			return false
		end if

		ultimoReg = reg

	case DOC_SAT_ANAL
		if( ultimoReg <> null ) then
			if not lerRegDocSATItemAnal(bf, reg, ultimoReg) then
				return false
			end if
			
			var node = @reg->itemAnal
			var parent = @ultimoReg->sat

			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_NFSCT
		if not lerRegDocNFSCT(bf, reg) then
			return false
		end if

		ultimoReg = reg

	case DOC_NFSCT_ANAL
		if( ultimoReg <> null ) then
			if not lerRegDocNFSCTItemAnal(bf, reg, ultimoReg) then
				return false
			end if
			
			var node = @reg->itemAnal
			var parent = @ultimoReg->nf

			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
	
	case DOC_NF_ELETRIC
		if not lerRegDocNFElet(bf, reg) then
			return false
		end if

		ultimoReg = reg

	case DOC_NF_ELETRIC_ANAL
		if( ultimoReg <> null ) then
			if not lerRegDocNFEletItemAnal(bf, reg, ultimoReg) then
				return false
			end if

			var node = @reg->itemAnal
			var parent = @ultimoReg->nf

			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
	
	case ITEM_ID
		if not lerRegItemId(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		if itemIdDict->lookup(reg->itemId.id) = null then
			itemIdDict->add(reg->itemId.id, @reg->itemId)
		end if

	case BEM_CIAP
		if not lerRegBemCiap(bf, reg) then
			return false
		end if
		
		ultimoBemCiap = @reg->bemCiap

		'adicionar ao dicionário
		if bemCiapDict->lookup(reg->bemCiap.id) = null then
			bemCiapDict->add(reg->bemCiap.id, @reg->bemCiap)
		end if

	case BEM_CIAP_INFO
		if not lerRegBemCiapInfo(bf, ultimoBemCiap) then
			return false
		end if
		
		'' deletar registro, já que vamos reusar o registro anterior
		reg->tipo = DESCONHECIDO

	case INFO_COMPL
		if not lerRegInfoCompl(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		if infoComplDict->lookup(reg->infoCompl.id) = null then
			infoComplDict->add(reg->infoCompl.id, @reg->infoCompl)
		end if

	case PARTICIPANTE
		if not lerRegParticipante(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		if participanteDict->lookup(reg->part.id) = null then
			participanteDict->add(reg->part.id, @reg->part)
		end if

	case APURACAO_ICMS_PERIODO
		if not lerRegApuIcmsPeriodo(bf, reg) then
			return false
		end if

		ultimoReg = reg
		
	case APURACAO_ICMS_PROPRIO
		if not lerRegApuIcmsProprio(bf, ultimoReg) then
			return false
		end if
		
		reg->tipo = DESCONHECIDO			'' deletar registro, já que vamos reusar o registro pai

	case APURACAO_ICMS_AJUSTE
		'' nota: como apuIcms e apuIcmsST estendem a mesma classe, pode-se acessar os campos comuns de qualquer classe filha
		if not lerRegApuIcmsAjuste(bf, reg, @ultimoReg->apuIcms) then
			return false
		end if
		
		var node = @reg->apuIcmsAjust
		var parent = @ultimoReg->apuIcms

		if parent->ajustesListHead = null then
			parent->ajustesListHead = node
		else
			parent->ajustesListTail->next_ = node
		end if

		parent->ajustesListTail = node
		node->next_ = null

	case APURACAO_ICMS_ST_PERIODO
		if not lerRegApuIcmsSTPeriodo(bf, reg) then
			return false
		end if

		ultimoReg = reg
		
	case APURACAO_ICMS_ST
		if not lerRegApuIcmsST(bf, ultimoReg) then
			return false
		end if
		
		reg->tipo = DESCONHECIDO			'' deletar registro, já que vamos reusar o registro pai

	case INVENTARIO_TOTAIS
		if not lerRegInventarioTotais(bf, reg) then
			return false
		end if
		
		ultimoInventario = @reg->invTotais
	
	case INVENTARIO_ITEM
		if not lerRegInventarioItem(bf, reg, ultimoInventario) then
			return false
		end if
	
	case CIAP_TOTAL
		if not lerRegCiapTotal(bf, reg) then
			return false
		end if
		
		ultimoCiap = @reg->ciapTotal
	
	case CIAP_ITEM
		if not lerRegCiapItem(bf, reg, ultimoCiap) then
			return false
		end if
	
		ultimoCiapItem = @reg->ciapItem
		var parent = ultimoCiap
		
		if parent->itemListHead = null then
			parent->itemListHead = ultimoCiapItem
		else
			parent->itemListTail->next_ = ultimoCiapItem
		end if

		parent->itemListTail = ultimoCiapItem
		ultimoCiapItem->next_ = null

	case CIAP_ITEM_DOC
		if not lerRegCiapItemDoc(bf, reg, ultimoCiapItem) then
			return false
		end if
		
		ultimoCiapItemDoc = @reg->ciapItemDoc
		var node = ultimoCiapItemDoc
		var parent = ultimoCiapItem

		if parent->docListHead = null then
			parent->docListHead = node
		else
			parent->docListTail->next_ = node
		end if

		parent->docListTail = node
		node->next_ = null

	case CIAP_ITEM_DOC_ITEM
		if not lerRegCiapItemDocItem(bf, reg, ultimoCiapItemDoc) then
			return false
		end if
		
		var node = @reg->ciapItemDocItem
		var parent = ultimoCiapItemDoc

		if parent->itemListHead = null then
			parent->itemListHead = node
		else
			parent->itemListTail->next_ = node
		end if

		parent->itemListTail = node
		node->next_ = null

	case ESTOQUE_PERIODO
		if not lerRegEstoquePeriodo(bf, reg) then
			return false
		end if
		
		ultimoEstoque = @reg->estPeriod
	
	case ESTOQUE_ITEM
		if not lerRegEstoqueItem(bf, reg, ultimoEstoque) then
			return false
		end if
		
	case ESTOQUE_ORDEM_PROD
		if not lerRegEstoqueOrdemProd(bf, reg, ultimoEstoque) then
			return false
		end if
	
	case MESTRE
		if not lerRegMestre(bf, reg) then
			return false
		end if
		
		regMestre = reg

	case OBS_LANCAMENTO
		if not lerRegObsLancamento(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		if obsLancamentoDict->lookup(reg->obsLanc.id) = null then
			obsLancamentoDict->add(reg->obsLanc.id, @reg->obsLanc)
		end if

	case CONTA_CONTAB
		if not lerRegContaContab(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		if contaContabDict->lookup(reg->contaContab.id) = null then
			contaContabDict->add(reg->contaContab.id, @reg->contaContab)
		end if

	case CENTRO_CUSTO
		if not lerRegCentroCusto(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		if centroCustoDict->lookup(reg->centroCusto.id) = null then
			centroCustoDict->add(reg->centroCusto.id, @reg->centroCusto)
		end if

	case FIM_DO_ARQUIVO
		pularLinha(bf)
		
		lerAssinatura(bf)
	
	case LUA_CUSTOM
		
		var luaFunc = cast(customLuaCb ptr, customLuaCbDict->lookup(tipo))->reader
		
		if luaFunc <> null then
			lua_getglobal(lua, luaFunc)
			lua_pushlightuserdata(lua, @bf)
			lua_newtable(lua)
			reg->lua.table = luaL_ref(lua, LUA_REGISTRYINDEX)
			lua_rawgeti(lua, LUA_REGISTRYINDEX, reg->lua.table)
			lua_call(lua, 2, 1)

			reg->tipo = LUA_CUSTOM
			reg->lua.tipo = tipo
			
		end if
	
	case else
		pularLinha(bf)
	end select

	function = true

end function

''''''''
function EfdSpedImport.carregar(nomeArquivo as string) as boolean

	dim bf as bfile
   
	if not bf.abrir( nomeArquivo ) then
		return false
	end if

	tipoArquivo = TIPO_ARQUIVO_EFD
	regListHead = null
	nroRegs = 0
	
	try
		var fsize = bf.tamanho - 6500 			'' descontar certificado digital no final do arquivo
		nroLinha = 1
		
		dim as TRegistro ptr tail = null

		do while bf.temProximo()		 
			var reg = new TRegistro

			if not onProgress(null, (bf.posicao / fsize) * 0.66) then
				exit do
			end if
			
			if lerRegistro( bf, reg ) then 
				if reg->tipo <> DESCONHECIDO then
					select case as const reg->tipo
					'' fim de arquivo?
					case FIM_DO_ARQUIVO
						delete reg
						onProgress(null, 1)
						exit do

					'' adicionar ao DB
					case DOC_NF, _
						 DOC_NF_ITEM, _
						 DOC_NF_ANAL, _
						 DOC_CT, _
						 ECF_REDUCAO_Z, _
						 DOC_SAT, _
						 DOC_NF_ITEM_RESSARC_ST, _
						 ITEM_ID, _
						 MESTRE
						addRegistroAoDB(reg)
					end select
					
					'' adicionar ao fim da lista
					if tail = null then
						regListHead = reg
						tail = reg
					else
						tail->next_ = reg
						tail = reg
					end if

					nroRegs += 1
				else
					delete reg
				end if
			 
				nroLinha += 1
			else
				exit do
			end if
		loop
	
	catch
		onError(!"\r\nErro ao carregar o registro da linha (" & nroLinha & !") do arquivo\r\n")
	endtry
	
	regListHead = ordenarRegistrosPorData(regListHead)
	
	onProgress(null, 1)

	function = true
  
	bf.fechar()
   
end function

''''''''
type HashCtx
	bf				as bfile ptr
	tamanhoSemSign	as longint
	bytesLidosTotal	as longint
end type

private function HashReadCB cdecl(ctx_ as any ptr, buffer as ubyte ptr, maxLen as long) as long
	var ctx = cast(HashCtx ptr, ctx_)
	
	if ctx->bytesLidosTotal + maxLen > ctx->tamanhoSemSign then
		maxLen = ctx->tamanhoSemSign - ctx->bytesLidosTotal
	end if
	
	var bytesLidos = ctx->bf->ler(buffer, maxLen)
	ctx->bytesLidosTotal += bytesLidos
	
	function = bytesLidos
	
end function

''''''''
function EfdSpedImport.lerInfoAssinatura(nomeArquivo as string) as InfoAssinatura ptr
	
	try
		var res = new InfoAssinatura
		
		var sh = new SSL_Helper
		var tamanhoAssinatura = ubound(assinaturaP7K_DER)+1
		var p7k = sh->Load_P7K(@assinaturaP7K_DER(0), tamanhoAssinatura)
		
		''
		var s = sh->Get_CommonName(p7k)
		if s <> null then
			res->assinante = *s
			deallocate s
		end if
		
		''
		s = sh->Get_AttributeFromAltName(p7k, AN_ATT_CPF)
		if s <> null then
			res->cpf = *s
			deallocate s
		else
			res->cpf = "00000000000"
		end if

		''
		var bf = new bfile()
		bf->abrir(nomeArquivo)
		var ctx = new HashCtx
		ctx->bf = bf
		ctx->tamanhoSemSign = bf->tamanho() - (tamanhoAssinatura + len(ASSINATURA_P7K_HEADER))
		ctx->bytesLidosTotal = 0
		
		s = sh->Compute_SHA1(@HashReadCB, ctx)
		if s <> null then
			res->hashDoArquivo = *s
			deallocate s
		end if
		
		bf->fechar()

		''
		sh->Free(p7k)
		delete sh
		
		function = res
	catch
		onError("Erro ao ler assinatura digital. As informações relativas à assinatura estarão em branco nos relatórios gerados")
		function = null
	endtry
	
end function

''''''''
function EfdSpedImport.adicionarMestre(reg as TMestre ptr) as long

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
		return 0
	end if
	
	return db->lastId()

end function

''''''''
function EfdSpedImport.adicionarDocEscriturado(doc as TDocDF ptr) as long
	
	if ISREGULAR(doc->situacao) then
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
				return 0
			end if
			
			return db->lastId()
			
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
				return 0
			end if
			
			return db->lastId()
		end if
	
	else
		'' !!!TODO!!! inserir em outra tabela para fazermos análises posteriores
	end if
	
	return 0
	
end function

''''''''
function EfdSpedImport.adicionarDocEscriturado(doc as TDocECF ptr) as long
	
	if ISREGULAR(doc->situacao) then
	
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
		
			if not db->execNonQuery(db_LRSInsertStmt) then
				onError("Erro ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
				return 0
			end if
			
			return db->lastId()
		end if
	
	else
		'' !!!TODO!!! inserir em outra tabela para fazermos análises posteriores
	end if

	return 0
end function

''''''''
function EfdSpedImport.adicionarDocEscriturado(doc as TDocSAT ptr) as long
	
	if ISREGULAR(doc->situacao) then
	
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
		
			if not db->execNonQuery(db_LRSInsertStmt) then
				onError("Erro ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
				return 0
			end if
			
			return db->lastId()
		end if
	
	else
		'' !!!TODO!!! inserir em outra tabela para fazermos análises posteriores
	end if
	
	return 0
end function

''''''''
function EfdSpedImport.adicionarItemNFEscriturado(item as TDocNFItem ptr) as long
	
	var doc = item->documentoPai
	if ISREGULAR(doc->situacao) then
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
		if opcoes->manterDb then
			db_itensNfLRInsertStmt->bind(21, item->itemId)
		else
			db_itensNfLRInsertStmt->bind(21, null)
		end if
		
		if not db->execNonQuery(db_itensNfLRInsertStmt) then
			onError("Erro ao inserir registro na EFD_Itens: " & *db->getErrorMsg())
			return 0
		end if
		
		return db->lastId()
	end if
	
	return 0
	
end function

''''''''
function EfdSpedImport.adicionarRessarcStEscriturado(doc as TDocNFItemRessarcSt ptr) as long

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
		return 0
	end if
	
	return db->lastId()
	
end function

''''''''
function EfdSpedImport.adicionarItemEscriturado(item as TItemId ptr) as long

	'' (id, descricao, ncm, cest, aliqInt)
	db_itensIdInsertStmt->reset()
	db_itensIdInsertStmt->bind(1, item->id)
	db_itensIdInsertStmt->bind(2, item->descricao)
	db_itensIdInsertStmt->bind(3, item->ncm)
	db_itensIdInsertStmt->bind(4, item->CEST)
	db_itensIdInsertStmt->bind(5, item->aliqICMSInt)
	
	if not db->execNonQuery(db_itensIdInsertStmt) then
		onError("Erro ao inserir registro na EFD_ItensId: " & *db->getErrorMsg())
		return 0
	end if
	
	return db->lastId()

end function

''''''''
function EfdSpedImport.adicionarAnalEscriturado(anal as TDocItemAnal ptr) as long

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
		return 0
	end if
	
	return db->lastId()

end function

''''''''
function EfdSpedImport.addRegistroAoDB(reg as TRegistro ptr) as long

	if opcoes->pularResumos andalso opcoes->pularAnalises andalso not opcoes->manterDb then
		return 0
	end if

	select case as const reg->tipo
	case DOC_NF
		return adicionarDocEscriturado(@reg->nf)
	case DOC_NF_ITEM
		return adicionarItemNFEscriturado(@reg->itemNF)
	case DOC_NF_ANAL
		return adicionarAnalEscriturado(@reg->itemAnal)
	case DOC_CT
		return adicionarDocEscriturado(@reg->ct)
	case DOC_ECF
		return adicionarDocEscriturado(@reg->ecf)
	case DOC_SAT
		return adicionarDocEscriturado(@reg->sat)
	case DOC_NF_ITEM_RESSARC_ST
		return adicionarRessarcStEscriturado(@reg->itemRessarcSt)
	case ITEM_ID
		if opcoes->manterDb then
			return adicionarItemEscriturado(@reg->itemId)
		end if
	case MESTRE
		return adicionarMestre(@reg->mestre)
	end select
	
	return 0
	
end function

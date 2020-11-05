#include once "EfdExt.bi"
#include once "EfdSpedImport.bi"
#include once "libs/ssl_helper.bi"
#include once "libs/trycatch.bi"

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
	lreInsertStmt as SQLiteStmt ptr, _
	itensNfLRInsertStmt as SQLiteStmt ptr, _
	lrsInsertStmt as SQLiteStmt ptr, _
	analInsertStmt as SQLiteStmt ptr, _
	ressarcStItensNfLRSInsertStmt as SQLiteStmt ptr, _
	itensIdInsertStmt as SQLiteStmt ptr, _
	mestreInsertStmt as SQLiteStmt ptr _
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
			var nf = cast(TDocNF ptr, pSrc1)
			isReg = ISREGULAR(nf->situacao)
			dReg = @nf->dataEntSaida
			nro = nf->numero
		case DOC_CT
			var ct = cast(TDocCT ptr, pSrc1)
			isReg = ISREGULAR(ct->situacao)
			dReg = @ct->dataEntSaida
			nro = ct->numero
		case DOC_NF_ITEM
			var doc = cast(TDocNFItem ptr, pSrc1)->documentoPai
			isReg = ISREGULAR(doc->situacao)
			dReg = @doc->dataEntSaida
			nro = doc->numero
		case ECF_REDUCAO_Z
			var redz = cast(TECFReducaoZ ptr, pSrc1)
			isReg = true
			dReg = @redz->dataMov
			nro = redz->numIni
		case DOC_SAT
			var sat = cast(TDocSAT ptr, pSrc1)
			isReg = true
			dReg = @sat->dataEntSaida
			nro = sat->numero
		case else
			isReg = false
			dReg = null
			nro = 0
		end select
		
		var date1 = iif(isReg, yyyyMmDd2Days(dReg) shl 32, 0) + nro

		select case as const pSrc2->tipo
		case DOC_NF, DOC_NFSCT, DOC_NF_ELETRIC
			var nf = cast(TDocNF ptr, pSrc2)
			isReg = ISREGULAR(nf->situacao)
			dReg = @nf->dataEntSaida
			nro = nf->numero
		case DOC_CT
			var ct = cast(TDocCT ptr, pSrc2)
			isReg = ISREGULAR(ct->situacao)
			dReg = @ct->dataEntSaida
			nro = ct->numero
		case DOC_NF_ITEM
			var doc = cast(TDocNFItem ptr, pSrc2)->documentoPai
			isReg = ISREGULAR(doc->situacao)
			dReg = @doc->dataEntSaida
			nro = doc->numero
		case ECF_REDUCAO_Z
			var redz = cast(TECFReducaoZ ptr, pSrc2)
			isReg = true
			dReg = @redz->dataMov
			nro = redz->numIni
		case DOC_SAT
			var sat = cast(TDocSAT ptr, pSrc2)
			isReg = true
			dReg = @sat->dataEntSaida
			nro = sat->numero
		case else
			isReg = false
			dReg = null
			nro = 0
		end select

		var date2 = iif(isReg, yyyyMmDd2Days(dReg) shl 32, 0) + nro

		if date2 < date1 then
			*ppDst = pSrc2
			ppDst = @pSrc2->prox
			pSrc2 = *ppDst
			if pSrc2 = NULL then
				*ppDst = pSrc1
				exit do
			end if
		else
			*ppDst = pSrc1
			ppDst = @pSrc1->prox
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
        var nn = n->prox
        n->prox = NULL
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
		onError("fora de sincronia na linha:" & nroLinha)
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
function EfdSpedImport.lerRegMestre(bf as bfile) as TMestre ptr
   
	var reg = new TMestre
	
	bf.char1		'pular |

	reg->versaoLayout= bf.varint
	reg->original 	= (bf.int1 = 0)
	bf.char1		'pular |
	reg->dataIni	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->dataFim	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nome	   	= bf.varchar
	reg->cnpj	   	= bf.varchar
	reg->cpf	   	= bf.varint
	reg->uf			= bf.varchar
	reg->ie			= bf.varchar
	reg->municip	= bf.varint
	reg->im  		= bf.varchar
	reg->suframa  	= bf.varchar
	reg->perfil  	= bf.char1
	bf.char1		'pular |
	reg->atividade	= bf.int1
	bf.char1		'pular |

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	return reg

end function

''''''''
function EfdSpedImport.lerRegParticipante(bf as bfile) as TParticipante ptr
   
	var reg = new TParticipante
	
	bf.char1		'pular |

	reg->id		= bf.varchar
	reg->nome	= bf.varchar
	reg->pais	= bf.varint
	reg->cnpj	= bf.varchar
	reg->cpf	= bf.varchar
	reg->ie		= bf.varchar
	reg->municip= bf.varint
	reg->suframa= bf.varchar
	reg->ender	= bf.varchar
	reg->num	= bf.varchar
	reg->compl	= bf.varchar
	reg->bairro	= bf.varchar
   
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocNF(bf as bfile) as TDocNF ptr

	var reg = new TDocNF
	bf.char1		'pular |

	reg->operacao		= bf.int1
	bf.char1		'pular |
	reg->emitente		= bf.int1
	bf.char1		'pular |
	reg->idParticipante	= bf.varchar
	reg->modelo			= bf.int2
	bf.char1		'pular |
	reg->situacao		= bf.int2
	bf.char1		'pular |
	reg->serie			= bf.varchar
	reg->numero			= bf.varint
	reg->chave			= bf.varchar
	reg->dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->valorTotal		= bf.vardbl
	reg->pagamento		= bf.varint
	reg->valorDesconto	= bf.vardbl
	reg->valorAbatimento= bf.vardbl
	reg->valorMerc		= bf.vardbl
	reg->frete			= bf.varint
	reg->valorFrete		= bf.vardbl
	reg->valorSeguro	= bf.vardbl
	reg->valorAcessorias= bf.vardbl
	reg->bcICMS			= bf.vardbl
	reg->ICMS			= bf.vardbl
	reg->bcICMSST		= bf.vardbl
	reg->ICMSST			= bf.vardbl
	reg->IPI			= bf.vardbl
	reg->PIS			= bf.vardbl
	reg->COFINS			= bf.vardbl
	reg->PISST			= bf.vardbl
	reg->COFINSST		= bf.vardbl
	reg->nroItens		= 0

	reg->itemAnalListHead = null
	reg->itemAnalListTail = null
	reg->infoComplListHead = null
	reg->infoComplListTail = null

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocNFInfo(bf as bfile, pai as TDocNF ptr) as TDocInfoCompl ptr

	var reg = new TDocInfoCompl
	
	bf.char1		'pular |

	reg->idCompl			= bf.varchar
	reg->extra				= bf.varchar
	reg->next_				= null
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocObs(bf as bfile) as TDocObs ptr

	var reg = new TDocObs

	bf.char1		'pular |

	reg->idLanc			= bf.varchar
	reg->extra			= bf.varchar
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocObsAjuste(bf as bfile) as TDocObsAjuste ptr

	var reg = new TDocObsAjuste
	
	bf.char1		'pular |

	reg->idAjuste	= bf.varchar
	reg->extra		= bf.varchar
	reg->idItem		= bf.varchar
	reg->bcICMS		= bf.vardbl
	reg->aliqICMS	= bf.vardbl
	reg->icms		= bf.vardbl
	reg->outros		= bf.vardbl
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocNFItem(bf as bfile, documentoPai as TDocNF ptr) as TDocNFItem ptr

	var reg = new TDocNFItem
	
	bf.char1		'pular |

	reg->documentoPai	= documentoPai
   
	reg->numItem		= bf.varint
	reg->itemId			= bf.varchar
	reg->descricao		= bf.varchar
	reg->qtd			= bf.vardbl
	reg->unidade		= bf.varchar
	reg->valor			= bf.vardbl
	reg->desconto		= bf.vardbl
	reg->indMovFisica	= bf.varint
	reg->cstICMS		= bf.varint
	reg->cfop			= bf.varint
	reg->codNatureza	= bf.varchar
	reg->bcICMS			= bf.vardbl
	reg->aliqICMS		= bf.vardbl
	reg->ICMS			= bf.vardbl
	reg->bcICMSST		= bf.vardbl
	reg->aliqICMSST		= bf.vardbl
	reg->ICMSST			= bf.vardbl
	reg->indApuracao	= bf.varint
	reg->cstIPI			= bf.varint
	reg->codEnqIPI		= bf.varchar
	reg->bcIPI			= bf.vardbl
	reg->aliqIPI		= bf.vardbl
	reg->IPI			= bf.vardbl
	reg->cstPIS			= bf.varint
	reg->bcPIS			= bf.vardbl
	reg->aliqPISPerc	= bf.vardbl
	reg->qSQLitecPIS		= bf.vardbl
	reg->aliqPISMoed	= bf.vardbl
	reg->PIS			= bf.vardbl
	reg->cstCOFINS		= bf.varint
	reg->bcCOFINS		= bf.vardbl
	reg->aliqCOFINSPerc = bf.vardbl
	reg->qSQLitecCOFINS	= bf.vardbl
	reg->aliqCOFINSMoed = bf.vardbl
	reg->COFINS			= bf.vardbl
	bf.varchar					'' pular código da conta
	if regMestre->versaoLayout >= 013 then
		bf.vardbl				'' pular VL_ABAT_NT
	end if

	documentoPai->nroItens 		+= 1
	
	reg->itemRessarcStListHead = null
	reg->itemRessarcStListTail = null

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocNFItemAnal(bf as bfile, documentoPai as TDocNF ptr) as TDocItemAnal ptr

	var reg = new TDocItemAnal

	bf.char1		'pular |

	reg->documentoPai= documentoPai
	reg->num		= documentoPai->itemAnalCnt
	documentoPai->itemAnalCnt += 1
	
	reg->cst		= bf.varint
	reg->cfop		= bf.varint
	reg->aliq		= bf.vardbl
	reg->valorOp	= bf.vardbl
	reg->bc			= bf.vardbl
	reg->ICMS		= bf.vardbl
	reg->bcST		= bf.vardbl
	reg->ICMSST		= bf.vardbl
	reg->redBC		= bf.vardbl
	reg->IPI		= bf.vardbl
	bf.varchar					'' pular código de observação

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	return reg

end function

''''''''
function EfdSpedImport.lerRegDocNFItemRessarcSt(bf as bfile, documentoPai as TDocNFItem ptr) as TDocNFItemRessarcSt ptr

	var reg = new TDocNFItemRessarcSt

	bf.char1		'pular |

	reg->documentoPai	= documentoPai
	
	reg->modeloUlt 			= bf.int2
	bf.char1		'pular |
	reg->numeroUlt 			= bf.varint
	reg->serieUlt  			= bf.varchar
	reg->dataUlt			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->idParticipanteUlt	= bf.varchar
	reg->qtdUlt				= bf.vardbl
	reg->valorUlt			= bf.vardbl
	reg->valorBcST			= bf.vardbl
	
	if bf.peek1 <> 13 then
		reg->chaveNFeUlt	= bf.varchar
		reg->numItemNFeUlt	= bf.varint
		reg->bcIcmsUlt		= bf.vardbl
		reg->aliqIcmsUlt	= bf.vardbl
		reg->limiteBcIcmsUlt= bf.vardbl
		reg->icmsUlt		= bf.vardbl
		reg->aliqIcmsStUlt	= bf.vardbl
		reg->res			= bf.vardbl
		reg->responsavelRet	= bf.int1
		bf.char1		'pular |
		reg->motivo			= bf.int1
		bf.char1		'pular |
		reg->chaveNFeRet	= bf.varchar
		reg->idParticipanteRet= bf.varchar
		reg->serieRet		= bf.varchar
		reg->numeroRet		= bf.varint
		reg->numItemNFeRet 	= bf.varint
		reg->tipDocArrecadacao= bf.int1
		bf.char1		'pular |
		reg->numDocArrecadacao= bf.varchar
	end if
   
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	return reg

end function

''''''''
function EfdSpedImport.lerRegDocNFDifal(bf as bfile, reg as TDocNF ptr) as TDocNF ptr

	bf.char1		'pular |

	reg->difal.fcp		= bf.vardbl
	reg->difal.icmsDest	= bf.vardbl
	reg->difal.icmsOrigem= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocCT(bf as bfile) as TDocCT ptr

	var reg = new TDocCT
	
	bf.char1		'pular |

	reg->operacao		= bf.int1
	bf.char1		'pular |
	reg->emitente		= bf.int1
	bf.char1		'pular |
	reg->idParticipante	= bf.varchar
	reg->modelo			= bf.int2
	bf.char1		'pular |
	reg->situacao		= bf.int2
	bf.char1		'pular |
	reg->serie			= bf.varchar
	bf.varchar		'pular sub-série
	reg->numero			= bf.varint
	reg->chave			= bf.varchar
	reg->dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->tipoCTe		= bf.varint
	reg->chaveRef		= bf.varchar
	reg->valorTotal		= bf.vardbl
	reg->valorDesconto	= bf.vardbl
	reg->frete			= bf.varint
	reg->valorServico	= bf.vardbl
	reg->bcICMS			= bf.vardbl
	reg->ICMS			= bf.vardbl
	reg->valorNaoTributado = bf.vardbl
	reg->codInfComplementar= bf.varchar
	bf.varchar		'pular código Conta Analitica
	
	'' códigos dos municípios de origem e de destino não aparecem em layouts antigos
	if bf.peek1 <> 13 and bf.peek1 <> 10 then 
		reg->municipioOrigem= bf.varint
		reg->municipioDestino= bf.varint
	end if
	
	reg->itemAnalListHead = null
	reg->itemAnalListTail = null
	reg->itemAnalCnt = 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocCTItemAnal(bf as bfile, docPai as TDocCT ptr) as TDocItemAnal ptr

	var reg = new TDocItemAnal
	
	bf.char1		'pular |

	reg->documentoPai= docPai

	reg->cst		= bf.varint
	reg->cfop		= bf.varint
	reg->aliq		= bf.vardbl
	reg->valorOp	= bf.vardbl
	reg->bc			= bf.vardbl
	reg->ICMS		= bf.vardbl
	reg->redBc		= bf.vardbl
	bf.varchar					'' pular cod obs
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocCTDifal(bf as bfile, reg as TDocCT ptr) as TDocCT ptr

	bf.char1		'pular |

	reg->difal.fcp		= bf.vardbl
	reg->difal.icmsDest	= bf.vardbl
	reg->difal.icmsOrigem= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegEquipECF(bf as bfile) as TEquipECF ptr

	var reg = new TEquipECF
	
	bf.char1		'pular |

	var modelo 		= bf.varchar
	reg->modelo		= iif(modelo = "2D", &h2D, valint(modelo))
	reg->modeloEquip= bf.varchar
	reg->numSerie 	= bf.varchar
	reg->numCaixa	= bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocECF(bf as bfile, equipECF as TEquipECF ptr) as TDocECF ptr

	var reg = new TDocECF
	
	bf.char1		'pular |

	reg->equipECF		= equipECF
	reg->operacao		= SAIDA
	var modelo = bf.varchar
	reg->modelo			= iif(modelo = "2D", &h2D, valint(modelo))
	reg->situacao		= bf.int2
	bf.char1		'pular |
	reg->numero			= bf.varint
	reg->dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->dataEntSaida	= reg->dataEmi
	reg->valorTotal		= bf.vardbl
	reg->PIS			= bf.vardbl
	reg->COFINS			= bf.vardbl
	reg->cpfCnpjAdquirente = bf.varchar
	reg->nomeAdquirente = bf.varchar
	reg->nroItens		= 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegECFReducaoZ(bf as bfile, equipECF as TEquipECF ptr) as TECFReducaoZ ptr

	var reg = new TECFReducaoZ
	
	bf.char1		'pular |

	reg->equipECF	= equipECF
	reg->dataMov	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->cro		= bf.varint
	reg->crz		= bf.varint
	reg->numOrdem	= bf.varint
	reg->valorFinal	= bf.vardbl
	reg->valorBruto	= bf.vardbl

	reg->numIni		= 2^20
	reg->numFim		= -1
	reg->itemAnalListHead = null
	reg->itemAnalListTail = null

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocECFItem(bf as bfile, documentoPai as TDocECF ptr) as TDocECFItem ptr

	var reg = new TDocECFItem
	
	bf.char1		'pular |

	reg->documentoPai	= documentoPai
   
	documentoPai->nroItens 		+= 1

	reg->numItem		= documentoPai->nroItens
	reg->itemId			= bf.varchar
	reg->qtd			= bf.vardbl
	reg->qtdCancelada	= bf.vardbl
	reg->unidade		= bf.varchar
	reg->valor			= bf.vardbl
	reg->cstICMS		= bf.varint
	reg->cfop			= bf.varint
	reg->aliqICMS		= bf.vardbl
	reg->PIS			= bf.vardbl
	reg->COFINS			= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocECFItemAnal(bf as bfile, documentoPai as TRegistro ptr) as TDocItemAnal ptr

	var reg = new TDocItemAnal

	bf.char1		'pular |

	reg->documentoPai= documentoPai
   
	reg->cst		= bf.varint
	reg->cfop		= bf.varint
	reg->aliq		= bf.vardbl
	reg->valorOp	= bf.vardbl
	reg->bc			= bf.vardbl
	reg->ICMS		= bf.vardbl
	bf.varchar					'' pular código de observação

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	return reg

end function

''''''''
function EfdSpedImport.lerRegDocSAT(bf as bfile) as TDocSAT ptr

	var reg = new TDocSAT
	
	bf.char1		'pular |

	reg->operacao		= SAIDA
	reg->modelo			= valint(bf.varchar)
	reg->situacao		= bf.int2
	bf.char1		'pular |
	reg->numero			= bf.varint
	reg->dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->valorTotal		= bf.vardbl
	reg->PIS			= bf.vardbl
	reg->COFINS			= bf.vardbl
	reg->cpfCnpjAdquirente = bf.varchar
	reg->serieEquip		= bf.varchar
	reg->chave 			= bf.varchar
	reg->descontos		= bf.vardbl
	reg->valorMerc 		= bf.vardbl
	reg->despesasAcess	= bf.vardbl
	reg->icms			= bf.vardbl
	reg->pisST			= bf.vardbl
	reg->cofinsST		= bf.vardbl
	reg->nroItens		= 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocSATItemAnal(bf as bfile, documentoPai as TRegistro ptr) as TDocItemAnal ptr

	var reg = new TDocItemAnal
	
	bf.char1		'pular |

	reg->documentoPai	= documentoPai
   
	reg->cst		= bf.varint
	reg->cfop		= bf.varint
	reg->aliq		= bf.vardbl
	reg->valorOp	= bf.vardbl
	reg->bc		= bf.vardbl
	reg->ICMS		= bf.vardbl
	bf.varchar					'' pular código de observação

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	return reg

end function


''''''''
function EfdSpedImport.lerRegDocNFSCT(bf as bfile) as TDocNF ptr

	var reg = new TDocNF
	
	bf.char1		'pular |

	reg->operacao		= bf.int1
	bf.char1		'pular |
	reg->emitente		= bf.int1
	bf.char1		'pular |
	reg->idParticipante	= bf.varchar
	reg->modelo			= bf.int2
	bf.char1		'pular |
	reg->situacao		= bf.int2
	bf.char1		'pular |
	reg->serie			= bf.varchar
	reg->subserie		= bf.varchar
	reg->numero			= bf.varint
	reg->dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->valorTotal		= bf.vardbl
	reg->valorDesconto	= bf.vardbl
	bf.vardbl		'pular valorServico
	bf.vardbl 		'pular valorServicoNT
	bf.vardbl 		'pular reg->valorTerceiro
	bf.vardbl 		'pular reg->valorDesp
	reg->bcICMS			= bf.vardbl
	reg->ICMS			= bf.vardbl
	bf.varchar		'pular cod_inf
	reg->PIS			= bf.vardbl
	reg->COFINS			= bf.vardbl
	bf.varchar		'pular cod_cta
	bf.varint		'pular tp_assinante
	reg->nroItens		= 0

	reg->itemAnalListHead = null
	reg->itemAnalListTail = null
	reg->itemAnalCnt = 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocNFSCTItemAnal(bf as bfile, documentoPai as TRegistro ptr) as TDocItemAnal ptr

	var reg = new TDocItemAnal
	
	bf.char1		'pular |

	reg->documentoPai= documentoPai
   
	reg->cst		= bf.varint
	reg->cfop		= bf.varint
	reg->aliq		= bf.vardbl
	reg->valorOp	= bf.vardbl
	reg->bc			= bf.vardbl
	reg->ICMS		= bf.vardbl
	bf.vardbl		'pular VL_BC_ICMS_UF
	bf.vardbl		'pular VL_ICMS_UF
	reg->redBC		= bf.vardbl
	bf.varchar		'pular COD_OBS

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	return reg

end function

''''''''
function EfdSpedImport.lerRegDocNFElet(bf as bfile) as TDocNF ptr

	var reg = new TDocNF
	
	bf.char1		'pular |

	reg->operacao		= bf.int1
	bf.char1		'pular |
	reg->emitente		= bf.int1
	bf.char1		'pular |
	reg->idParticipante	= bf.varchar
	reg->modelo			= bf.int2
	bf.char1		'pular |
	reg->situacao		= bf.int2
	bf.char1		'pular |
	reg->serie			= bf.varchar
	reg->subserie		= bf.varchar
	bf.varchar		'pular cod_cons
	reg->numero			= bf.varint
	reg->dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->valorTotal		= bf.vardbl
	reg->valorDesconto	= bf.vardbl
	bf.varchar		'pular vl_forn
	bf.varchar 		'pular vl_serv_nt
	bf.varchar		'pular vl_terc
	bf.varchar		'pular vl_da
	reg->bcICMS			= bf.vardbl
	reg->ICMS			= bf.vardbl
	reg->bcICMSST		= bf.vardbl
	reg->ICMSST			= bf.vardbl
	bf.varchar		'pular cod_inf
	reg->PIS			= bf.vardbl
	reg->COFINS			= bf.vardbl
	bf.varchar		'pular tp_ligacao
	bf.varchar		'pular cod_grupo_tensao
	if regMestre->versaoLayout >= 014 then
		reg->chave		= bf.varchar		
		bf.varchar		'pular fin_doce
		bf.varchar		'pular chv_doce_ref
		bf.varchar		'pular ind_dest
		bf.varchar		'pular cod_mun_dest
		bf.varchar		'pular cod_cta
	end if
	reg->nroItens		= 0

	reg->itemAnalListHead = null
	reg->itemAnalListTail = null
	reg->itemAnalCnt = 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegDocNFEletItemAnal(bf as bfile, documentoPai as TRegistro ptr) as TDocItemAnal ptr

	var reg = new TDocItemAnal
	
	bf.char1		'pular |

	reg->documentoPai	= documentoPai
   
	reg->cst		= bf.varint
	reg->cfop		= bf.varint
	reg->aliq		= bf.vardbl
	reg->valorOp	= bf.vardbl
	reg->bc			= bf.vardbl
	reg->ICMS		= bf.vardbl
	reg->bcST		= bf.vardbl
	reg->ICMSST	= bf.vardbl
	reg->redBC		= bf.vardbl
	bf.varchar		'pular COD_OBS
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	return reg

end function

''''''''
function EfdSpedImport.lerRegItemId(bf as bfile) as TItemId ptr

	var reg = new TItemId
	
	bf.char1		'pular |

	reg->id			  	= bf.varchar
	reg->descricao	  	= bf.varchar
	reg->codBarra		= bf.varchar
	reg->codAnterior	= bf.varchar
	reg->unidInventario = bf.varchar
	reg->tipoItem		= bf.varint
	reg->ncm			= bf.varint
	reg->exIPI		  	= bf.varchar
	reg->codGenero	  	= bf.varint
	reg->codServico	  	= bf.varchar
	reg->aliqICMSInt	= bf.vardbl
	'CEST só é obrigatório a partir de 2017
	if bf.peek1 <> 13 and bf.peek1 <> 10 then 
	  reg->CEST		  	= bf.varint
	end if

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegBemCiap(bf as bfile) as TBemCiap ptr

	var reg = new TBemCiap
	
	bf.char1		'pular |

	reg->id			  	= bf.varchar
	reg->tipoMerc		= bf.varint
	reg->descricao	  	= bf.varchar
	reg->principal		= bf.varchar
	reg->codAnal	  	= bf.varchar
	reg->parcelas		= bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegBemCiapInfo(bf as bfile, reg as TBemCiap ptr) as TBemCiap ptr

	bf.char1		'pular |

	reg->codCusto		= bf.varchar
	reg->funcao	  		= bf.varchar
	reg->vidaUtil		= bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegObsLancamento(bf as bfile) as TObsLancamento ptr

	var reg = new TObsLancamento
	
	bf.char1		'pular |

	reg->id				= bf.varchar
	reg->descricao	  	= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegContaContab(bf as bfile) as TContaContab ptr

	var reg = new TContaContab
	
	bf.char1		'pular |

	reg->dataInc		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->codNat			= bf.varchar
	reg->ind			= bf.varchar
	reg->nivel			= bf.varint
	reg->id			 	= bf.varchar
	reg->descricao	  	= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegCentroCusto(bf as bfile) as TCentroCusto ptr

	var reg = new TCentroCusto
	
	bf.char1		'pular |

	reg->dataInc		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->id			 	= bf.varchar
	reg->descricao	  	= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegInfoCompl(bf as bfile) as TInfoCompl ptr

	var reg = new TInfoCompl
	
	bf.char1		'pular |

	reg->id				= bf.varchar
	reg->descricao	  	= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegApuIcmsPeriodo(bf as bfile) as TApuracaoIcmsPropPeriodo ptr

   var reg = new TApuracaoIcmsPropPeriodo
   
   bf.char1		'pular |

   reg->dataIni		  = ddMmYyyy2YyyyMmDd(bf.varchar)
   reg->dataFim		  = ddMmYyyy2YyyyMmDd(bf.varchar)

   'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

   return reg

end function

''''''''
function EfdSpedImport.lerRegApuIcmsProprio(bf as bfile, reg as TApuracaoIcmsPropPeriodo ptr) as TApuracaoIcmsPropPeriodo ptr

	bf.char1		'pular |

	reg->totalDebitos			= bf.vardbl
	reg->ajustesDebitos			= bf.vardbl
	reg->totalAjusteDeb			= bf.vardbl
	reg->estornosCredito		= bf.vardbl
	reg->totalCreditos			= bf.vardbl
	reg->ajustesCreditos		= bf.vardbl
	reg->totalAjusteCred		= bf.vardbl
	reg->estornoDebitos			= bf.vardbl
	reg->saldoCredAnterior		= bf.vardbl
	reg->saldoDevedorApurado	= bf.vardbl
	reg->totalDeducoes			= bf.vardbl
	reg->icmsRecolher			= bf.vardbl
	reg->saldoCredTransportar	= bf.vardbl
	reg->debExtraApuracao		= bf.vardbl

	reg->ajustesListHead 		= null
	reg->ajustesListTail 		= null
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegApuIcmsAjuste(bf as bfile, pai as TApuracaoIcmsPeriodo ptr) as TApuracaoIcmsAjuste ptr

	var reg = new TApuracaoIcmsAjuste
	
	bf.char1		'pular |
	
	reg->codigo 	= bf.varchar
	reg->descricao = bf.varchar
	reg->valor 	= bf.vardbl
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegApuIcmsSTPeriodo(bf as bfile) as TApuracaoIcmsSTPeriodo ptr

	var reg = new TApuracaoIcmsSTPeriodo
	
	bf.char1		'pular |

	reg->UF		 	= bf.varchar
	reg->dataIni	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->dataFim	= ddMmYyyy2YyyyMmDd(bf.varchar)

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegApuIcmsST(bf as bfile, reg as TApuracaoIcmsSTPeriodo ptr) as TApuracaoIcmsSTPeriodo ptr

	bf.char1		'pular |

	reg->mov					= bf.varint
	reg->saldoCredAnterior		= bf.vardbl
	reg->devolMercadorias		= bf.vardbl
	reg->totalRessarciment		= bf.vardbl
	reg->totalOutrosCred		= bf.vardbl
	reg->ajustesCreditos		= bf.vardbl
	reg->totalRetencao			= bf.vardbl
	reg->totalOutrosDeb			= bf.vardbl
	reg->ajustesDebitos			= bf.vardbl
	reg->saldoAntesDed			= bf.vardbl
	reg->totalDeducoes			= bf.vardbl
	reg->icmsRecolher			= bf.vardbl
	reg->saldoCredTransportar	= bf.vardbl
	reg->debExtraApuracao		= bf.vardbl

	reg->ajustesListHead 		= null
	reg->ajustesListTail 		= null
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegInventarioTotais(bf as bfile) as TInventarioTotais ptr

	var reg = new TInventarioTotais
	
	bf.char1		'pular |

	reg->dataInventario 	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->valorTotalEstoque 	= bf.vardbl
	reg->motivoInventario	= bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegInventarioItem(bf as bfile, inventarioPai as TInventarioTotais ptr) as TInventarioItem ptr

	var reg = new TInventarioItem
	
	bf.char1		'pular |

	reg->dataInventario 	= inventarioPai->dataInventario
	reg->itemId 	 		= bf.varchar
	reg->unidade 			= bf.varchar
	reg->qtd	 			= bf.vardbl
	reg->valorUnitario		= bf.vardbl
	reg->valorItem			= bf.vardbl
	reg->indPropriedade		= bf.varint
	reg->idParticipante		= bf.varchar
	reg->txtComplementar	= bf.varchar
	reg->codConta			= bf.varchar
	reg->valorItemIR		= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegCiapTotal(bf as bfile) as TCiapTotal ptr

	var reg = new TCiapTotal
	
	bf.char1		'pular |

	reg->dataIni 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->dataFim 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->saldoInicialICMS 	= bf.vardbl
	reg->parcelasSoma 		= bf.vardbl
	reg->valorTributExpSoma = bf.vardbl
	reg->valorTotalSaidas 	= bf.vardbl
	reg->indicePercSaidas 	= bf.vardbl
	reg->valorIcmsAprop 	= bf.vardbl
	reg->valorOutrosCred	= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegCiapItem(bf as bfile, pai as TCiapTotal ptr) as TCiapItem ptr

	var reg = new TCiapItem
	
	bf.char1		'pular |

	reg->pai			= pai
	reg->bemId 	 		= bf.varchar
	reg->dataMov 		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->tipoMov 		= bf.varchar
	reg->valorIcms	 	= bf.vardbl
	reg->valorIcmsSt	= bf.vardbl
	reg->valorIcmsFrete	= bf.vardbl
	reg->valorIcmsDifal	= bf.vardbl
	reg->parcela		= bf.varint
	reg->valorParcela	= bf.vardbl
	reg->docCnt			= 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegCiapItemDoc(bf as bfile, pai as TCiapItem ptr) as TCiapItemDoc ptr

	var reg = new TCiapItemDoc
	
	bf.char1		'pular |

	reg->pai			= pai
	reg->indEmi 		= bf.varint
	reg->idParticipante = bf.varchar
	reg->modelo			= bf.varint
	reg->serie			= bf.varchar
	reg->numero			= bf.varint
	reg->chaveNFe		= bf.varchar
	reg->dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	if bf.peek1 <> 13 andalso bf.peek1 <> 10 then 
		bf.varchar '' pular NUM_DA
	end if
	pai->docCnt += 1

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegCiapItemDocItem(bf as bfile, pai as TCiapItemDoc ptr) as TCiapItemDocItem ptr

	var reg = new TCiapItemDocItem
	
	bf.char1		'pular |

	reg->pai			= pai
	reg->num			= bf.varint
	reg->itemId 		= bf.varchar
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
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegEstoquePeriodo(bf as bfile) as TEstoquePeriodo ptr

	var reg = new TEstoquePeriodo
	
	bf.char1		'pular |

	reg->dataIni 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->dataFim 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegEstoqueItem(bf as bfile, pai as TEstoquePeriodo ptr) as TEstoqueItem ptr

	var reg = new TEstoqueItem
	
	bf.char1		'pular |

	reg->pai				= pai
	bf.varchar		'pular DT_EST (é a mesma do DT_FIN do K100)
	reg->itemId 	 		= bf.varchar
	reg->qtd 				= bf.vardbl
	reg->tipoEst			= bf.varint
	reg->idParticipante		= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
function EfdSpedImport.lerRegEstoqueOrdemProd(bf as bfile, pai as TEstoquePeriodo ptr) as TEstoqueOrdemProd ptr

	var reg = new TEstoqueOrdemProd
	
	bf.char1		'pular |

	reg->pai			= pai
	reg->dataIni 	 	= ddMmYyyy2YyyyMmDd(bf.varchar)
	var dtFim = bf.varchar
	reg->dataFim 	 	= iif(len(dtFim) > 0, ddMmYyyy2YyyyMmDd(dtFim), "99991231")
	reg->idOrdem		= bf.varchar
	reg->itemId 	 	= bf.varchar
	reg->qtd 			= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	return reg

end function

''''''''
private sub EfdSpedImport.lerAssinatura(bf as bfile)

	'' verificar header
	var header = bf.nchar(len(ASSINATURA_P7K_HEADER))
	if header <> ASSINATURA_P7K_HEADER then
		onError("header da assinatura P7K não reconhecido")
	end if
	
	var lgt = (bf.tamanho - bf.posicao) + 1
	
	redim this.assinaturaP7K_DER(0 to lgt-1)
	
	bf.ler(assinaturaP7K_DER(), lgt)

end sub

''''''''
function EfdSpedImport.lerRegistro(bf as bfile) as TRegistro ptr
	static as zstring * 4+1 tipostr
	
	var tipo = lerTipo(bf, @tipostr)
	var reg = cast(TRegistro ptr, null)

	select case as const tipo
	case DOC_NF
		reg = lerRegDocNF(bf)
		if reg = null then
			return null
		end if
		
		ultimoReg = reg

	case DOC_NF_INFO
		if( ultimoReg <> null ) then
			var node = lerRegDocNFInfo(bf, cast(TDocNF ptr, ultimoReg))
			if node = null then
				return null
			end if
			reg = node
			
			var parent = cast(TDocNF ptr, ultimoReg)
			
			if parent->infoComplListHead = null then
				parent->infoComplListHead = node
			else
				parent->infoComplListTail->next_ = node
			end if
			
			parent->infoComplListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			return null
		end if
		
	case DOC_NF_ITEM
		if( ultimoReg <> null ) then
			reg = lerRegDocNFItem(bf, cast(TDocNF ptr, ultimoReg))
			if reg = null then
				return null
			end if
			
			ultimoDocNFItem = cast(TDocNFItem ptr, reg)
		else
			pularLinha(bf)
			return null
		end if

	case DOC_NF_ANAL
		if( ultimoReg <> null ) then
			var node = lerRegDocNFItemAnal(bf, cast(TDocNF ptr, ultimoReg))
			if node = null  then
				return null
			end if
			reg = node
			
			var parent = cast(TDocNF ptr, ultimoReg)
			
			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			return null
		end if

	case DOC_NF_OBS
		if( ultimoReg <> null ) then
			var node = lerRegDocObs(bf)
			if node = null then
				return null
			end if
			reg = node
			
			ultimoDocObs = node
			
			var parent = cast(TDocNF ptr, ultimoReg)
			
			if parent->obsListHead = null then
				parent->obsListHead = node
			else
				parent->obsListTail->next_ = node
			end if
			
			parent->obsListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			return null
		end if

	case DOC_NF_OBS_AJUSTE
		if( ultimoDocObs <> null ) then
			var node = lerRegDocObsAjuste(bf)
			if node = null  then
				return null
			end if
			reg = node
			
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
			return null
		end if
		
	case DOC_NF_DIFAL
		if( ultimoReg <> null ) then
			lerRegDocNFDifal(bf, cast(TDocNF ptr, ultimoReg))
			return null
		else
			pularLinha(bf)
			return null
		end if
		
	case DOC_NF_ITEM_RESSARC_ST
		if( ultimoDocNFItem <> null ) then
			var node = lerRegDocNFItemRessarcSt(bf, ultimoDocNFItem)
			if node = null then
				return null
			end if
			reg = node
			
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
			return null
		end if

	case DOC_CT
		reg = lerRegDocCT(bf)
		if reg = null then
			return null
		end if

		ultimoReg = reg

	case DOC_CT_ANAL
		if( ultimoReg <> null ) then
			var node = lerRegDocCTItemAnal(bf, cast(TDocCT ptr, ultimoReg))
			if node = null then
				return null
			end if
			reg = node

			var parent = cast(TDocCT ptr, ultimoReg)
			
			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			return null
		end if
		
	case DOC_CT_DIFAL
		if( ultimoReg <> null ) then
			lerRegDocCTDifal(bf, cast(TDocCT ptr, ultimoReg))
			return null
		else
			pularLinha(bf)
			return null
		end if

	case DOC_ECF
		if( ultimoEquipECF <> null ) then
			var node = lerRegDocECF(bf, ultimoEquipECF)
			if node = null then
				return null
			end if
			reg = node

			ultimoReg = reg
			
			if ultimoECFRedZ->numIni > node->numero then
				ultimoECFRedZ->numIni = node->numero
			end if

			if ultimoECFRedZ->numFim < node->numero then
				ultimoECFRedZ->numFim = node->numero
			end if
		else
			pularLinha(bf)
			return null
		end if
		
	case ECF_REDUCAO_Z
		if( ultimoEquipECF <> null ) then
			reg = lerRegECFReducaoZ(bf, ultimoEquipECF)
			if reg = null then
				return null
			end if

			ultimoECFRedZ = cast(TECFReducaoZ ptr, reg)
		else
			pularLinha(bf)
			ultimoECFRedZ = null
			return null
		end if
		
	case DOC_ECF_ITEM
		if( ultimoReg <> null ) then
			reg = lerRegDocECFItem(bf, cast(TDocECF ptr, ultimoReg))
			if reg = null then
				return null
			end if
		else
			pularLinha(bf)
			return null
		end if

	case DOC_ECF_ANAL
		if( ultimoECFRedZ <> null ) then
			var node = lerRegDocECFItemAnal(bf, ultimoECFRedZ)
			if node = null then
				return null
			end if
			reg = node
			
			var parent = ultimoECFRedZ
			
			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			return null
		end if

	case EQUIP_ECF
		reg = lerRegEquipECF(bf)
		if reg = null  then
			return null
		end if
		
		ultimoEquipECF = cast(TEquipECF ptr, reg)

	case DOC_SAT
		reg = lerRegDocSAT(bf)
		if reg = null then
			return null
		end if

		ultimoReg = reg

	case DOC_SAT_ANAL
		if( ultimoReg <> null ) then
			var node = lerRegDocSATItemAnal(bf, cast(TDocSAT ptr, ultimoReg))
			if reg = null then
				return null
			end if
			reg = node
			
			var parent = cast(TDocSAT ptr, ultimoReg)

			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			return null
		end if

	case DOC_NFSCT
		reg = lerRegDocNFSCT(bf)
		if reg = null then
			return null
		end if

		ultimoReg = reg

	case DOC_NFSCT_ANAL
		if( ultimoReg <> null ) then
			var node = lerRegDocNFSCTItemAnal(bf, cast(TDocNF ptr, ultimoReg))
			if node = null then
				return null
			end if
			reg = node
			
			var parent = cast(TDocNF ptr, ultimoReg)

			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			return null
		end if
	
	case DOC_NF_ELETRIC
		reg = lerRegDocNFElet(bf)
		if reg = null then
			return null
		end if

		ultimoReg = reg

	case DOC_NF_ELETRIC_ANAL
		if( ultimoReg <> null ) then
			var node = lerRegDocNFEletItemAnal(bf, cast(TDocNF ptr, ultimoReg))
			if node = null then
				return null
			end if
			reg = node

			var parent = cast(TDocNF ptr, ultimoReg)

			if parent->itemAnalListHead = null then
				parent->itemAnalListHead = node
			else
				parent->itemAnalListTail->next_ = node
			end if
			
			parent->itemAnalListTail = node
			node->next_ = null
		else
			pularLinha(bf)
			return null
		end if
	
	case ITEM_ID
		var node = lerRegItemId(bf)
		if node = null  then
			return null
		end if
		reg = node

		'adicionar ao dicionário
		if itemIdDict->lookup(node->id) = null then
			itemIdDict->add(node->id, node)
		end if

	case BEM_CIAP
		var node = lerRegBemCiap(bf)
		if node = null then
			return null
		end if
		reg = node
		
		ultimoBemCiap = node

		'adicionar ao dicionário
		if bemCiapDict->lookup(node->id) = null then
			bemCiapDict->add(node->id, node)
		end if

	case BEM_CIAP_INFO
		lerRegBemCiapInfo(bf, ultimoBemCiap)
		return null

	case INFO_COMPL
		var node = lerRegInfoCompl(bf)
		if node = null then
			return null
		end if
		reg = node

		'adicionar ao dicionário
		if infoComplDict->lookup(node->id) = null then
			infoComplDict->add(node->id, node)
		end if

	case PARTICIPANTE
		var node = lerRegParticipante(bf)
		if node = null then
			return null
		end if
		reg = node

		'adicionar ao dicionário
		if participanteDict->lookup(node->id) = null then
			participanteDict->add(node->id, node)
		end if

	case APURACAO_ICMS_PERIODO
		reg = lerRegApuIcmsPeriodo(bf)
		if reg = null then
			return null
		end if

		ultimoReg = reg
		
	case APURACAO_ICMS_PROPRIO
		lerRegApuIcmsProprio(bf, cast(TApuracaoIcmsPropPeriodo ptr, ultimoReg))
		return null

	case APURACAO_ICMS_AJUSTE
		var node = lerRegApuIcmsAjuste(bf, cast(TApuracaoIcmsPropPeriodo ptr, ultimoReg))
		if node = null then
			return null
		end if
		reg = node
		
		var parent = cast(TApuracaoIcmsPropPeriodo ptr, ultimoReg)

		if parent->ajustesListHead = null then
			parent->ajustesListHead = node
		else
			parent->ajustesListTail->next_ = node
		end if

		parent->ajustesListTail = node
		node->next_ = null

	case APURACAO_ICMS_ST_PERIODO
		reg = lerRegApuIcmsSTPeriodo(bf)
		if reg = null then
			return null
		end if

		ultimoReg = reg
		
	case APURACAO_ICMS_ST
		lerRegApuIcmsST(bf, cast(TApuracaoIcmsSTPeriodo ptr, ultimoReg))
		return null

	case INVENTARIO_TOTAIS
		reg = lerRegInventarioTotais(bf)
		if reg = null then
			return null
		end if
		
		ultimoInventario = cast(TInventarioTotais ptr, reg)
	
	case INVENTARIO_ITEM
		reg = lerRegInventarioItem(bf, ultimoInventario)
		if reg = null then
			return null
		end if
	
	case CIAP_TOTAL
		reg = lerRegCiapTotal(bf)
		if reg = null then
			return null
		end if
		
		ultimoCiap = cast(TCiapTotal ptr, reg)
	
	case CIAP_ITEM
		var node = lerRegCiapItem(bf, ultimoCiap)
		if node = null then
			return null
		end if
		reg = node
	
		ultimoCiapItem = node
		var parent = ultimoCiap
		
		if parent->itemListHead = null then
			parent->itemListHead = node
		else
			parent->itemListTail->next_ = node
		end if

		parent->itemListTail = node
		node->next_ = null

	case CIAP_ITEM_DOC
		var node = lerRegCiapItemDoc(bf, ultimoCiapItem)
		if node = null then
			return null
		end if
		reg = node
		
		ultimoCiapItemDoc = node
		var parent = ultimoCiapItem

		if parent->docListHead = null then
			parent->docListHead = node
		else
			parent->docListTail->next_ = node
		end if

		parent->docListTail = node
		node->next_ = null

	case CIAP_ITEM_DOC_ITEM
		var node = lerRegCiapItemDocItem(bf, ultimoCiapItemDoc)
		if node = null then
			return null
		end if
		reg = node
		
		var parent = ultimoCiapItemDoc

		if parent->itemListHead = null then
			parent->itemListHead = node
		else
			parent->itemListTail->next_ = node
		end if

		parent->itemListTail = node
		node->next_ = null

	case ESTOQUE_PERIODO
		reg = lerRegEstoquePeriodo(bf)
		if reg = null then
			return null
		end if
		
		ultimoEstoque = cast(TEstoquePeriodo ptr, reg)
	
	case ESTOQUE_ITEM
		reg = lerRegEstoqueItem(bf, ultimoEstoque)
		if reg = null then
			return null
		end if
		
	case ESTOQUE_ORDEM_PROD
		reg = lerRegEstoqueOrdemProd(bf, ultimoEstoque)
		if reg = null then
			return null
		end if
	
	case MESTRE
		reg = lerRegMestre(bf)
		if reg = null then
			return null
		end if
		
		regMestre = cast(TMestre ptr, reg)

	case OBS_LANCAMENTO
		var node = lerRegObsLancamento(bf)
		if node = null then
			return null
		end if
		reg = node

		'adicionar ao dicionário
		if obsLancamentoDict->lookup(node->id) = null then
			obsLancamentoDict->add(node->id, node)
		end if

	case CONTA_CONTAB
		var node = lerRegContaContab(bf)
		if node = null then
			return null
		end if
		reg = node

		'adicionar ao dicionário
		if contaContabDict->lookup(node->id) = null then
			contaContabDict->add(node->id, node)
		end if

	case CENTRO_CUSTO
		var node = lerRegCentroCusto(bf)
		if node = null then
			return null
		end if
		reg = node

		'adicionar ao dicionário
		if centroCustoDict->lookup(node->id) = null then
			centroCustoDict->add(node->id, node)
		end if

	case FIM_DO_ARQUIVO
		pularLinha(bf)
		lerAssinatura(bf)
		reg = new TRegistro
	
	case LUA_CUSTOM
		
		var luaFunc = cast(customLuaCb ptr, customLuaCbDict->lookup(tipo))->reader
		
		if luaFunc <> null then
			lua_getglobal(lua, luaFunc)
			lua_pushlightuserdata(lua, @bf)
			lua_newtable(lua)
			var node = new TLuaReg
			node->table = luaL_ref(lua, LUA_REGISTRYINDEX)
			lua_rawgeti(lua, LUA_REGISTRYINDEX, node->table)
			lua_call(lua, 2, 1)
			reg = node
		end if
	
	case else
		pularLinha(bf)
		return null
	end select

	reg->tipo = tipo
	reg->linha = nroLinha

	return reg

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
			if not onProgress(null, (bf.posicao / fsize) * 0.66) then
				exit do
			end if
			
			var reg = lerRegistro( bf )
			if reg <> null then 
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
						tail->prox = reg
						tail = reg
					end if

					nroRegs += 1
				else
					delete reg
				end if
			end if
			
			nroLinha += 1
		loop
	
	catch
		onError(!"\r\nErro ao carregar o registro da linha (" & nroLinha & !") do arquivo\r\n")
	endtry
	
	regListHead = ordenarRegistrosPorData(regListHead)
	
	onProgress(null, 1)

	return true
  
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
		onError("ao ler assinatura digital. As informações relativas à assinatura estarão em branco nos relatórios gerados")
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
		onError("ao inserir registro na EFD_Mestre: " & *db->getErrorMsg())
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
			db_LREInsertStmt->bind(1, valint(regMestre->dataIni))
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
				onError("ao inserir registro na EFD_LRE: " & *db->getErrorMsg())
				return 0
			end if
			
			return db->lastId()
			
		else
			'' (periodo, cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp, IE)
			db_LRSInsertStmt->reset()
			db_LRSInsertStmt->bind(1, valint(regMestre->dataIni))
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
				onError("ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
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
			db_LRSInsertStmt->bind(1, valint(regMestre->dataIni))
			db_LRSInsertStmt->bind(2, doc->cpfCnpjAdquirente)
			db_LRSInsertStmt->bind(3, 35)
			db_LRSInsertStmt->bind(4, 0)
			db_LRSInsertStmt->bind(5, doc->numero)
			db_LRSInsertStmt->bind(6, doc->modelo)
			db_LRSInsertStmt->bind(7, doc->chave)
			db_LRSInsertStmt->bind(8, doc->dataEmi)
			db_LRSInsertStmt->bind(9, doc->valorTotal)
		
			if not db->execNonQuery(db_LRSInsertStmt) then
				onError("ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
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
			db_LRSInsertStmt->bind(1, valint(regMestre->dataIni))
			db_LRSInsertStmt->bind(2, 0) '' não é possível usar doc->cpfCnpjAdquirente, porque relatório do BO vem sem essa info
			db_LRSInsertStmt->bind(3, 35)
			db_LRSInsertStmt->bind(4, 0)
			db_LRSInsertStmt->bind(5, doc->numero)
			db_LRSInsertStmt->bind(6, doc->modelo)
			db_LRSInsertStmt->bind(7, doc->chave)
			db_LRSInsertStmt->bind(8, doc->dataEmi)
			db_LRSInsertStmt->bind(9, doc->valorTotal)
		
			if not db->execNonQuery(db_LRSInsertStmt) then
				onError("ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
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
		db_itensNfLRInsertStmt->bind(1, valint(regMestre->dataIni))
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
			onError("ao inserir registro na EFD_Itens: " & *db->getErrorMsg())
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
	db_ressarcStItensNfLRSInsertStmt->bind(1, valint(regMestre->dataIni))
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
		onError("ao inserir registro na EFD_Ressarc_Itens: " & *db->getErrorMsg())
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
		onError("ao inserir registro na EFD_ItensId: " & *db->getErrorMsg())
		return 0
	end if
	
	return db->lastId()

end function

''''''''
function EfdSpedImport.adicionarAnalEscriturado(anal as TDocItemAnal ptr) as long

	var doc = cast(TDocNF ptr, anal->documentoPai)
	var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipante) )
	
	var uf = iif(part->municip >= 1100000 and part->municip <= 5399999, part->municip \ 100000, 99)

	'' (operacao, periodo, cnpj, uf, serie, numero, modelo, numReg, cst, cst_origem, cst_tribut, cfop, aliq, valorOp, bc, icms, bcIcmsST, icmsST, redBC, ipi)
	db_analInsertStmt->reset()
	db_analInsertStmt->bind(1, doc->operacao)
	db_analInsertStmt->bind(2, valint(regMestre->dataIni))
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
		onError("ao inserir registro na EDF_Anal: " & *db->getErrorMsg())
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
		return adicionarDocEscriturado(cast(TDocDF ptr, reg))
	case DOC_NF_ITEM
		return adicionarItemNFEscriturado(cast(TDocNFItem ptr, reg))
	case DOC_NF_ANAL
		return adicionarAnalEscriturado(cast(TDocItemAnal ptr, reg))
	case DOC_CT
		return adicionarDocEscriturado(cast(TDocDF ptr, reg))
	case DOC_ECF
		return adicionarDocEscriturado(cast(TDocECF ptr, reg))
	case DOC_SAT
		return adicionarDocEscriturado(cast(TDocSAT ptr, reg))
	case DOC_NF_ITEM_RESSARC_ST
		return adicionarRessarcStEscriturado(cast(TDocNFItemRessarcSt ptr, reg))
	case ITEM_ID
		if opcoes->manterDb then
			return adicionarItemEscriturado(cast(TItemId ptr, reg))
		end if
	case MESTRE
		return adicionarMestre(cast(TMestre ptr, reg))
	end select
	
	return 0
	
end function

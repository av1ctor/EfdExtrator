
#include once "efd.bi"
#include once "bfile.bi"
#include once "hash.bi"
#include once "ExcelWriter.bi"
#include once "vbcompat.bi"
#include once "ssl_helper.bi"
#include once "DocxFactoryDyn.bi"
#include once "DB.bi"

dim shared as string codUF2Sigla(11 to 53)
dim shared as string situacao2String(0 to __TipoSituacao__LEN__-1)

const ASSINATURA_P7K_HEADER = "SBRCAAEPDR"

private sub tablesCtor constructor
	codUF2Sigla(11)="RO"
	codUF2Sigla(12)="AC"
	codUF2Sigla(13)="AM"
	codUF2Sigla(14)="RR"
	codUF2Sigla(15)="PA"
	codUF2Sigla(16)="AP"
	codUF2Sigla(17)="TO"
	codUF2Sigla(21)="MA"
	codUF2Sigla(22)="PI"
	codUF2Sigla(23)="CE"
	codUF2Sigla(24)="RN"
	codUF2Sigla(25)="PB"
	codUF2Sigla(26)="PE"
	codUF2Sigla(27)="AL"
	codUF2Sigla(28)="SE"
	codUF2Sigla(29)="BA"
	codUF2Sigla(31)="MG"
	codUF2Sigla(32)="ES"
	codUF2Sigla(33)="RJ"
	codUF2Sigla(35)="SP"
	codUF2Sigla(41)="PR"
	codUF2Sigla(42)="SC"
	codUF2Sigla(43)="RS"
	codUF2Sigla(50)="MS"
	codUF2Sigla(51)="MT"
	codUF2Sigla(52)="GO"
	codUF2Sigla(53)="DF"

	situacao2String(REGULAR) = "REG"
	situacao2String(EXTEMPORANEO) = "EXTEMP"
	situacao2String(CANCELADO) = "CANC"
	situacao2String(CANCELADO_EXT) = "CANC EXTEMP"
	situacao2String(DENEGADO) = "DENEG"
	situacao2String(INUTILIZADO) = "INUT"
	situacao2String(COMPLEMENTAR) = "COMPL"
	situacao2String(COMPLEMENTAR_EXT) = "COMPL EXTEMP"
	situacao2String(REGIME_ESPECIAL) = "REG ESP"
	situacao2String(SUBSTITUIDO) = "SUBST"
end sub

''''''''
constructor Efd()
	''
	chaveDFeDict.init(2^20)
	nfeDestSafiFornecido = false
	nfeEmitSafiFornecido = false
	itemNFeSafiFornecido = false
	cteSafiFornecido = false
	dfeListHead = null
	dfeListTail = null
	
	''
	efdDFeDict.init(2^20)
	efdDfeListHead = null
	efdDfeListTail = null
	
	''
	baseTemplatesDir = ExePath + "\templates\"
	
	dfwd = new DocxFactoryDyn
	
	municipDict.init(2^10, true, true, true)
	
	''
	dbConfig = new TDb
	dbConfig->open(ExePath + "\db\config.db")
	
end constructor

destructor Efd()
	
	''
	dbConfig->close()
	delete dbConfig
	
	''
	municipDict.end_()
	
	delete dfwd
	
	''
	efdDFeDict.end_()

	do while efdDfeListHead <> null
		var next_ = efdDfeListHead->next_
		delete efdDfeListHead
		efdDfeListHead = next_
	loop
	
	''
	chaveDFeDict.end_()
	
	do while dfeListHead <> null
		var next_ = dfeListHead->next_
		if dfeListHead->modelo = NFE then
			do while dfeListHead->nfe.itemListHead <> null
				var next_ = dfeListHead->nfe.itemListHead->next_
				delete dfeListHead->nfe.itemListHead
				dfeListHead->nfe.itemListHead = next_
			loop
		end if
		delete dfeListHead
		dfeListHead = next_
	loop
end destructor

''''''''
private sub pularLinha(bf as bfile) 

	'ler até \r
	do
	  if bf.char1 = 13 then
			exit do
		end if
	loop
	
	'pular \n
	bf.char1 
	
end sub

''''''''
private function lerLinha(bf as bfile) as string

	var res = ""
	var c1 = " "
	
	'ler até \r
	do
		c1[0] = bf.char1
		if c1[0] = 13 then
			exit do
		end if
		
		res += c1
	loop
	
	'pular \n
	bf.char1

	function = res
	
end function

''''''''
private function lerTipo(bf as bfile) as TipoRegistro

	bf.char1 ' pular |
	
	var tipo = bf.char4

	function = DESCONHECIDO
	
	select case as const tipo[0]
	case asc("0")
		select case tipo
		case "0150"
			function = PARTICIPANTE
		case "0200"
			function = ITEM_ID
		case "0000"
			function = MESTRE
		end select
	case asc("C")
		select case tipo
		case "C100"
			function = DOC_NFE
		case "C170"
			function = DOC_NFE_ITEM
		case "C190"
			function = DOC_NFE_ANAL
		case "C101"
			function = DOC_NFE_DIFAL
		end select
	case asc("D")
		select case tipo
		case "D100"
			function = DOC_CTE
		case "D190"
			function = DOC_CTE_ANAL
		case "D101"
			function = DOC_CTE_DIFAL
		end select
	case asc("E")	
		select case tipo
		case "E100"
			function = APURACAO_ICMS_PERIODO
		case "E110"
			function = APURACAO_ICMS_PROPRIO
		case "E200"
			function = APURACAO_ICMS_ST_PERIODO
		case "E210"
			function = APURACAO_ICMS_ST
		end select
	case asc("9")
		select case tipo
		case "9999"
			function = EOF_
		end select
	end select

end function

''''''''
private function lerRegMestre(bf as bfile, reg as TRegistro ptr) as Boolean
   
	bf.char1		'pular |

	reg->mestre.versaoLayout= bf.varint
	reg->mestre.original 	= (bf.int1 = 0)
	bf.char1		'pular |
	reg->mestre.dataIni		= bf.varchar
	reg->mestre.dataFim		= bf.varchar
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
	bf.char1
	bf.char1

	function = true

end function

''''''''
private function lerRegParticipante(bf as bfile, reg as TRegistro ptr) as Boolean
   
	bf.char1		'pular |

	reg->part.id		= bf.varchar
	reg->part.nome		= bf.varchar
	reg->part.pais	   	= bf.varint
	reg->part.cnpj	   	= bf.varchar
	reg->part.cpf	   	= bf.varint
	reg->part.ie		= bf.varchar
	reg->part.municip	= bf.varint
	reg->part.suframa  	= bf.varchar
	reg->part.ender	   	= bf.varchar
	reg->part.num		= bf.varchar
	reg->part.compl	   	= bf.varchar
	reg->part.bairro	= bf.varchar
   
	'pular \r\n
	bf.char1
	bf.char1

	function = true

end function

''''''''
private function lerRegDocNFe(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->nfe.operacao		= bf.int1
	bf.char1		'pular |
	reg->nfe.emitente		= bf.int1
	bf.char1		'pular |
	reg->nfe.idParticipante	= bf.varchar
	reg->nfe.modelo			= bf.int2
	bf.char1		'pular |
	reg->nfe.situacao		= bf.int2
	bf.char1		'pular |
	reg->nfe.serie			= bf.varint
	reg->nfe.numero			= bf.varint
	reg->nfe.chave			= bf.varchar
	reg->nfe.dataEmi		= bf.varchar
	reg->nfe.dataEntSaida	= bf.varchar
	reg->nfe.valorTotal		= bf.vardbl
	reg->nfe.pagamento		= bf.int1
	bf.char1		'pular |
	reg->nfe.valorDesconto	= bf.vardbl
	reg->nfe.valorAbatimento= bf.vardbl
	reg->nfe.valorMerc		= bf.vardbl
	reg->nfe.frete			= bf.int1
	bf.char1		'pular |
	reg->nfe.valorFrete		= bf.vardbl
	reg->nfe.valorSeguro	= bf.vardbl
	reg->nfe.valorAcessorias= bf.vardbl
	reg->nfe.bcICMS			= bf.vardbl
	reg->nfe.ICMS			= bf.vardbl
	reg->nfe.bcICMSST		= bf.vardbl
	reg->nfe.ICMSST			= bf.vardbl
	reg->nfe.IPI			= bf.vardbl
	reg->nfe.PIS			= bf.vardbl
	reg->nfe.COFINS			= bf.vardbl
	reg->nfe.PISST			= bf.vardbl
	reg->nfe.COFINSST		= bf.vardbl
	reg->nfe.nroItens		= 0

	reg->nfe.itemAnalListHead = null
	reg->nfe.itemAnalListTail = null

	'pular \r\n
	bf.char1
	bf.char1

	function = true

end function

''''''''
private function lerRegDocNFeItem(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNFe ptr) as Boolean

	bf.char1		'pular |

	reg->itemNFe.documentoPai	= documentoPai
   
	reg->itemNFe.numItem		= bf.varint
	reg->itemNFe.itemId			= bf.varchar
	reg->itemNFe.descricao		= bf.varchar
	reg->itemNFe.qtd			= bf.vardbl
	reg->itemNFe.unidade		= bf.varchar
	reg->itemNFe.valor			= bf.vardbl
	reg->itemNFe.desconto		= bf.vardbl
	reg->itemNFe.indMovFisica	= bf.varint
	reg->itemNFe.cstICMS		= bf.varint
	reg->itemNFe.cfop			= bf.varint
	reg->itemNFe.codNatureza	= bf.varchar
	reg->itemNFe.bcICMS			= bf.vardbl
	reg->itemNFe.aliqICMS		= bf.vardbl
	reg->itemNFe.ICMS			= bf.vardbl
	reg->itemNFe.bcICMSST		= bf.vardbl
	reg->itemNFe.aliqICMSST		= bf.vardbl
	reg->itemNFe.ICMSST			= bf.vardbl
	reg->itemNFe.indApuracao	= bf.varint
	reg->itemNFe.cstIPI			= bf.varint
	reg->itemNFe.codEnqIPI		= bf.varchar
	reg->itemNFe.bcIPI			= bf.vardbl
	reg->itemNFe.aliqIPI		= bf.vardbl
	reg->itemNFe.IPI			= bf.vardbl
	reg->itemNFe.cstPIS			= bf.varint
	reg->itemNFe.bcPIS			= bf.vardbl
	reg->itemNFe.aliqPISPerc	= bf.vardbl
	reg->itemNFe.qtdBcPIS		= bf.vardbl
	reg->itemNFe.aliqPISMoed	= bf.vardbl
	reg->itemNFe.PIS			= bf.vardbl
	reg->itemNFe.cstCOFINS		= bf.varint
	reg->itemNFe.bcCOFINS		= bf.vardbl
	reg->itemNFe.aliqCOFINSPerc = bf.vardbl
	reg->itemNFe.qtdBcCOFINS	= bf.vardbl
	reg->itemNFe.aliqCOFINSMoed = bf.vardbl
	reg->itemNFe.COFINS			= bf.vardbl
	bf.varchar					'' pular código da conta

	documentoPai->nroItens 		+= 1

	'pular \r\n
	bf.char1
	bf.char1

	function = true

end function

''''''''
private function lerRegDocNFeItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

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
	reg->itemAnal.IPI		= bf.vardbl
	bf.varchar					'' pular código de observação

	'pular \r\n
	bf.char1
	bf.char1
	
	function = true

end function

''''''''
private function lerRegDocNFeDifal(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNFe ptr) as Boolean

	bf.char1		'pular |

	documentoPai->difal.fcp			= bf.vardbl
	documentoPai->difal.icmsDest	= bf.vardbl
	documentoPai->difal.icmsOrigem	= bf.vardbl

	'pular \r\n
	bf.char1
	bf.char1

	function = true

end function

''''''''
private function lerRegDocCTe(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->cte.operacao		= bf.int1
	bf.char1		'pular |
	reg->cte.emitente		= bf.int1
	bf.char1		'pular |
	reg->cte.idParticipante	= bf.varchar
	reg->cte.modelo			= bf.int2
	bf.char1		'pular |
	reg->cte.situacao		= bf.int2
	bf.char1		'pular |
	reg->cte.serie			= bf.varint
	bf.varchar		'pular sub-série
	reg->cte.numero			= bf.varint
	reg->cte.chave			= bf.varchar
	reg->cte.dataEmi		= bf.varchar
	reg->cte.dataEntSaida	= bf.varchar
	reg->cte.tipoCTe		= bf.int1
	bf.char1		'pular |
	reg->cte.chaveRef		= bf.varchar
	reg->cte.valorTotal		= bf.vardbl
	reg->cte.valorDesconto	= bf.vardbl
	reg->cte.frete			= bf.int1
	bf.char1		'pular |
	reg->cte.valorServico	= bf.vardbl
	reg->cte.bcICMS			= bf.vardbl
	reg->cte.ICMS			= bf.vardbl
	reg->cte.valorNaoTributado = bf.vardbl
	reg->cte.codInfComplementar	= bf.varchar
	bf.varchar		'pular código Conta Analitica
	
	'' códigos dos municípios de origem e de destino não aparecem em layouts antigos
	if bf.peek1 <> 13 then 
		reg->cte.municipioOrigem 	= bf.varint
		reg->cte.municipioDestino	= bf.varint
	end if
	
	reg->cte.itemAnalListHead = null
	reg->cte.itemAnalListTail = null

	'pular \r\n
	bf.char1
	bf.char1

	function = true

end function

''''''''
private function lerRegDocCTeItemAnal(bf as bfile, reg as TRegistro ptr, docPai as TRegistro ptr) as Boolean

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
	bf.char1
	bf.char1

	function = true

end function

''''''''
private function lerRegDocCTeDifal(bf as bfile, reg as TRegistro ptr, docPai as TDocCTe ptr) as Boolean

	bf.char1		'pular |

	docPai->difal.fcp		= bf.vardbl
	docPai->difal.icmsDest	= bf.vardbl
	docPai->difal.icmsOrigem= bf.vardbl

	'pular \r\n
	bf.char1
	bf.char1

	function = true

end function

''''''''
private function lerRegItemId(bf as bfile, reg as TRegistro ptr) as Boolean

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
	if bf.peek1 <> 13 then 
	  reg->itemId.CEST		  	= bf.varint
	end if

	'pular \r\n
	bf.char1
	bf.char1

	function = true

end function

''''''''
private function lerRegApuIcmsPeriodo(bf as bfile, reg as TRegistro ptr) as Boolean

   bf.char1		'pular |

   reg->apuIcms.dataIni		  = bf.varchar
   reg->apuIcms.dataFim		  = bf.varchar

   'pular \r\n
   bf.char1
   bf.char1

   function = true

end function

''''''''
private function lerRegApuIcmsProprio(bf as bfile, reg as TRegistro ptr) as Boolean

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

	'pular \r\n
	bf.char1
	bf.char1

	function = true

end function

''''''''
private function lerRegApuIcmsSTPeriodo(bf as bfile, reg as TRegistro ptr) as Boolean

   bf.char1		'pular |

   reg->apuIcmsST.UF		 	 = bf.varchar
   reg->apuIcmsST.dataIni		 = bf.varchar
   reg->apuIcmsST.dataFim		 = bf.varchar

   'pular \r\n
   bf.char1
   bf.char1

   function = true

end function

''''''''
private function lerRegApuIcmsST(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->apuIcmsST.mov						= bf.varint
	reg->apuIcmsST.saldoCredAnterior		= bf.vardbl
	reg->apuIcmsST.devolMercadorias			= bf.vardbl
	reg->apuIcmsST.totalRessarciment		= bf.vardbl
	reg->apuIcmsST.totalOutrosCred			= bf.vardbl
	reg->apuIcmsST.ajusteCred				= bf.vardbl
	reg->apuIcmsST.totalRetencao			= bf.vardbl
	reg->apuIcmsST.totalOutrosDeb			= bf.vardbl
	reg->apuIcmsST.ajusteDeb				= bf.vardbl
	reg->apuIcmsST.saldoAntesDed			= bf.vardbl
	reg->apuIcmsST.totalDeducoes			= bf.vardbl
	reg->apuIcmsST.icmsRecolher				= bf.vardbl
	reg->apuIcmsST.saldoCredTransportar		= bf.vardbl
	reg->apuIcmsST.debExtraApuracao			= bf.vardbl

	'pular \r\n
	bf.char1
	bf.char1

	function = true

end function

''''''''
private sub Efd.lerAssinatura(bf as bfile)

	'' verificar header
	var header = bf.nchar(len(ASSINATURA_P7K_HEADER))
	if header <> ASSINATURA_P7K_HEADER then
		print "Erro: header da assinatura P7K não reconhecido"
	end if
	
	var lgt = (bf.tamanho - bf.posicao) + 1
	
	redim this.assinaturaP7K_DER(0 to lgt-1)
	
	bf.ler(assinaturaP7K_DER(), lgt)

end sub

''''''''
private function Efd.lerRegistro(bf as bfile, reg as TRegistro ptr) as Boolean

	reg->tipo = lerTipo(bf)

	select case reg->tipo
	case DOC_NFE
		if not lerRegDocNFe(bf, reg) then
			return false
		end if

		ultimoReg = reg

	case DOC_NFE_ITEM
		if not lerRegDocNFeItem(bf, reg, @ultimoReg->nfe) then
			return false
		end if

	case DOC_NFE_ANAL
		if not lerRegDocNFeItemAnal(bf, reg, ultimoReg) then
			return false
		end if
		
		if ultimoReg->nfe.itemAnalListHead = null then
			ultimoReg->nfe.itemAnalListHead = @reg->itemAnal
		else
			ultimoReg->nfe.itemAnalListTail->next_ = @reg->itemAnal
		end if
		
		ultimoReg->nfe.itemAnalListTail = @reg->itemAnal
		reg->itemAnal.next_ = null
		
	case DOC_NFE_DIFAL
		if not lerRegDocNFeDifal(bf, reg, @ultimoReg->nfe) then
			return false
		end if
		
		reg->tipo = DESCONHECIDO			'' deletar registro, já que vamos reusar o registro pai

	case DOC_CTE
		if not lerRegDocCTe(bf, reg) then
			return false
		end if

		ultimoReg = reg

	case DOC_CTE_ANAL
		if not lerRegDocCTeItemAnal(bf, reg, ultimoReg) then
			return false
		end if

		if ultimoReg->cte.itemAnalListHead = null then
			ultimoReg->cte.itemAnalListHead = @reg->itemAnal
		else
			ultimoReg->cte.itemAnalListTail->next_ = @reg->itemAnal
		end if
		
		ultimoReg->cte.itemAnalListTail = @reg->itemAnal
		reg->itemAnal.next_ = null
		
	case DOC_CTE_DIFAL
		if not lerRegDocCTeDifal(bf, reg, @reg->cte) then
			return false
		end if
		
		reg->tipo = DESCONHECIDO			'' deletar registro, já que vamos reusar o registro pai

	case ITEM_ID
		if not lerRegItemId(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		if itemIdDict.lookup(reg->itemId.id) = null then
			itemIdDict.add(reg->itemId.id, @reg->itemId)
		end if

	case PARTICIPANTE
		if not lerRegParticipante(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		if participanteDict.lookup(reg->part.id) = null then
			participanteDict.add(reg->part.id, @reg->part)
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

	case MESTRE
		if not lerRegMestre(bf, reg) then
			return false
		end if

	case EOF_
		pularLinha(bf)
		
		lerAssinatura(bf)
	
	case else
		pularLinha(bf)
	end select

	function = true

end function

''''''''
private function situacaoSintegra2SituacaoEfd(sit as byte) as TipoSituacao
	select case sit
	case asc("N")
		return REGULAR
	case asc("S")
		return CANCELADO
	case asc("E")
		return EXTEMPORANEO
	case asc("X")
		return CANCELADO_EXT
	case asc("2")
		return DENEGADO
	case asc("4")
		return INUTILIZADO
	case else
		return REGULAR
	end select

end function

''''''''
private function lerRegSintegraDocumento(bf as bfile, reg as TRegistro ptr) as Boolean

	reg->docSint.cnpj 		= bf.nchar(14)
	reg->docSint.ie 		= bf.nchar(14)
	reg->docSint.dataEmi 	= bf.char8
	reg->docSint.uf 		= bf.char2
	reg->docSint.modelo 	= bf.int2
	reg->docSint.serie 		= valint(bf.nchar(3))
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("¨") then
		bf.char1
		reg->docSint.numero = bf.int9
	else
		reg->docSint.numero = bf.int6
	end if
	reg->docSint.cfop 		= bf.int4
	reg->docSint.operacao 	= iif( bf.char1 = asc("T"), ENTRADA, SAIDA )
	reg->docSint.valorTotal = bf.dbl13_2
	reg->docSint.bcICMS 	= bf.dbl13_2
	reg->docSint.ICMS 		= bf.dbl13_2
	reg->docSint.valorIsento= bf.dbl13_2
	reg->docSint.valorOutras= bf.dbl13_2
	reg->docSint.aliqICMS 	= bf.dbl4_2
	reg->docSint.situacao 	=	situacaoSintegra2SituacaoEfd( bf.char1 )

	'' ler chave NF-e no final da linha, se for um sintegra convertido pelo SAFI
	if bf.peek1 <> 13 then
		reg->docSint.chave 	= bf.nchar(44)
	end if

	'pular \r\n
	bf.char1
	bf.char1

	function = true
end function

''''''''
private function lerRegSintegraDocumentoST(bf as bfile, reg as TRegistro ptr) as Boolean

	reg->docSint.cnpj 		= bf.nchar(14)
	reg->docSint.ie 		= bf.nchar(14)
	reg->docSint.dataEmi	= bf.char8
	reg->docSint.uf 		= bf.char2
	reg->docSint.modelo 	= bf.int2
	reg->docSint.serie 		= valint(bf.nchar(3))
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("¨") then
		bf.char1
		reg->docSint.numero = bf.int9
	else
		reg->docSint.numero = bf.int6
	end if
	reg->docSint.cfop 		= bf.int4
	reg->docSint.operacao 	= iif( bf.char1 = asc("T"), ENTRADA, SAIDA )
	reg->docSint.bcICMSST 	= bf.dbl13_2
	reg->docSint.ICMSST 	= bf.dbl13_2
	reg->docSint.despesasAcess = bf.dbl13_2
	reg->docSint.situacao 	= situacaoSintegra2SituacaoEfd( bf.char1 )
	bf.nchar(30)

	'pular \r\n
	bf.char1
	bf.char1

	function = true
end function

''''''''
private function lerRegSintegraDocumentoIPI(bf as bfile, reg as TRegistro ptr) as Boolean

	reg->docSint.cnpj 		= bf.nchar(14)
	reg->docSint.ie 		= bf.nchar(14)
	reg->docSint.dataEmi 	= bf.char8
	reg->docSint.uf 		= bf.char2
	reg->docSint.serie 		= valint(bf.nchar(3))
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("¨") then
		bf.char1
		reg->docSint.numero = bf.int9
	else
		reg->docSint.numero = bf.int6
	end if
	reg->docSint.cfop 		= bf.int4
	reg->docSint.valorTotal = bf.dbl13_2
	reg->docSint.valorIPI 	= bf.dbl13_2
	reg->docSint.valorIsentoIPI = bf.dbl13_2
	reg->docSint.valorOutrasIPI = bf.dbl13_2
	bf.nchar(1+20)

	'pular \r\n
	bf.char1
	bf.char1

	function = true
end function

#define GENSINTEGRAKEY(r) (r->docSint.cnpj + r->docSint.ie + r->docSint.dataEmi + r->docSint.uf + str(r->docSint.serie) + str(r->docSint.numero))
  
''''''''
private function Efd.lerRegistroSintegra(bf as bfile, reg as TRegistro ptr) as Boolean

	var tipo = bf.int2

	select case tipo
	case SINTEGRA_DOCUMENTO
		reg->tipo = SINTEGRA_DOCUMENTO
		if not lerRegSintegraDocumento(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		reg->docSint.chaveHash = GENSINTEGRAKEY(reg)
		var antReg = cast(TRegistro ptr, sintegraDict.lookup(reg->docSint.chaveHash))
		if antReg = null then
			sintegraDict.add(reg->docSint.chaveHash, reg)
		else
			'' para cada alíquota diferente há um novo registro 50, mas nós só queremos os valores totais
			antReg->docSint.valorTotal	+= reg->docSint.valorTotal
			antReg->docSint.bcICMS		+= reg->docSint.bcICMS
			antReg->docSint.ICMS		+= reg->docSint.ICMS
			antReg->docSint.valorIsento += reg->docSint.valorIsento
			antReg->docSint.valorOutras += reg->docSint.valorOutras

			reg->tipo = DESCONHECIDO 
		end if

	case SINTEGRA_DOCUMENTO_ST
		reg->tipo = SINTEGRA_DOCUMENTO
		if not lerRegSintegraDocumentoST(bf, reg) then
			return false
		end if

		reg->docSint.chaveHash = GENSINTEGRAKEY(reg)
		var antReg = cast(TRegistro ptr, sintegraDict.lookup(reg->docSint.chaveHash))
		'' NOTA: pode existir registro 53 sem o correspondente 50, para quando só há ICMS ST, sem destaque ICMS próprio
		if antReg = null then
			sintegraDict.add(reg->docSint.chaveHash, reg)
		else
			antReg->docSint.bcICMSST		+= reg->docSint.bcICMSST
			antReg->docSint.ICMSST			+= reg->docSint.ICMSST
			antReg->docSint.despesasAcess	+= reg->docSint.despesasAcess
			reg->tipo = DESCONHECIDO
		end if
	  
	case SINTEGRA_DOCUMENTO_IPI
		reg->tipo = SINTEGRA_DOCUMENTO
		if not lerRegSintegraDocumentoIPI(bf, reg) then
			return false
		end if

		reg->docSint.chaveHash = GENSINTEGRAKEY(reg)
		var antReg = cast(TRegistro ptr, sintegraDict.lookup(reg->docSint.chaveHash))
		if antReg = null then
			print "ERRO: Sintegra 53 sem 50: "; reg->docSint.chaveHash
		else
			antReg->docSint.valorIPI		= reg->docSint.valorIPI
			antReg->docSint.valorIsentoIPI	= reg->docSint.valorIsentoIPI
			antReg->docSint.valorOutrasIPI	= reg->docSint.valorOutrasIPI
		end if

		reg->tipo = DESCONHECIDO 

	case else
		pularLinha(bf)
		reg->tipo = DESCONHECIDO
	end select

	function = true

end function

''''''''
function Efd.carregarSintegra(bf as bfile, mostrarProgresso as sub(porCompleto as double)) as Boolean
	
   var fsize = bf.tamanho

   do while bf.temProximo()		 
	  var reg = new TRegistro
	  
	  if lerRegistroSintegra( bf, reg ) then 
		 if mostrarProgresso <> NULL then
			mostrarProgresso(bf.posicao / fsize)
		 end if 
			
		 if reg->tipo <> DESCONHECIDO then
			if regListHead = null then
			   regListHead = reg
			   regListTail = reg
			else
			   regListTail->next_ = reg
			   regListTail = reg
			end if
			
			nroRegs += 1
		 else
			delete reg
		 end if
		 
	  else
		 exit do
	  end if
   loop
	   
   function = true

end function

private sub STR2YYYYMMDD(s as const zstring ptr, d as zstring ptr)
	'(mid(s,5) + mid(s,3,2) + left(s,2))
	static as integer ltb(0 to 7) = { 4, 5, 6, 7, 2, 3, 0, 1 }
	
	if len(*s) = 0 then
		(*d)[0] = 0
	else
		for i as integer = 0 to 7
			(*d)[i] = (*s)[ltb(i)]
		next
		(*d)[8] = 0
	end if
	
end sub

''''''''
sub Efd.addRegistroOrdenadoPorData(reg as TRegistro ptr)

	if regListHead = null then
		regListHead = reg
		regListTail = reg
		return
	end if

	dim as zstring * 8+1 demi
	
	select case reg->tipo
	case DOC_NFE
		STR2YYYYMMDD(reg->nfe.dataEmi, demi)
	case DOC_CTE
		STR2YYYYMMDD(reg->cte.dataEmi, demi)
	case DOC_NFE_ITEM
		STR2YYYYMMDD(reg->itemNFe.documentoPai->dataEmi, demi)
	end select
	
	var n = regListHead
	dim as TRegistro ptr p = null
	dim as zstring * 8+1 n_demi
	do 
		var dotest = true
		select case n->tipo
		case DOC_NFE
			STR2YYYYMMDD(n->nfe.dataEmi, n_demi)
		case DOC_CTE
			STR2YYYYMMDD(n->cte.dataEmi, n_demi)
		case DOC_NFE_ITEM
			STR2YYYYMMDD(n->itemNFe.documentoPai->dataEmi, n_demi)
		case else
			dotest = false
		end select
		
		if dotest then
			if n_demi > demi then
				reg->next_ = n
				if p <> null then
					p->next_ = reg
				else
					regListHead = reg
				end if
				exit do
			end if
		end if
		
		p = n
		n = n->next_
	loop until n = null
	
	if n = null then
		regListTail->next_ = reg
		regListTail = reg
	end if

end sub

''''''''
function Efd.carregarTxt(nomeArquivo as String, mostrarProgresso as sub(porCompleto as double)) as Boolean

	dim bf as bfile
   
	if not bf.abrir( nomeArquivo ) then
		return false
	end if

	participanteDict.init(2^20)
	itemIdDict.init(2^20)	 
	sintegraDict.init(2^20)

	regListHead = null
	regListTail = null
	nroRegs = 0

	if bf.peek1 <> asc("|") then
		tipoArquivo = TIPO_ARQUIVO_SINTEGRA
		function = carregarSintegra(bf, mostrarProgresso)
	else
		tipoArquivo = TIPO_ARQUIVO_EFD
		var fsize = bf.tamanho - 6500 			'' descontar certificado digital no final do arquivo

		do while bf.temProximo()		 
			var reg = new TRegistro

			if mostrarProgresso <> NULL then
				mostrarProgresso(bf.posicao / fsize)
			end if 
				
			if lerRegistro( bf, reg ) then 
				if reg->tipo <> DESCONHECIDO then
					select case reg->tipo
					'' fim de arquivo?
					case EOF_
						delete reg
						if mostrarProgresso <> NULL then
							mostrarProgresso(1)
						end if					
						exit do

					'' ordernar por data emi
					case DOC_NFE, DOC_NFE_ITEM, DOC_CTE
						addRegistroOrdenadoPorData(reg)
					
					'' registro sem data, adicionar ao fim da fila
					case else
						if regListHead = null then
							regListHead = reg
							regListTail = reg
						else
							regListTail->next_ = reg
							regListTail = reg
						end if
					end select

					nroRegs += 1
				else
					delete reg
				end if
			 
			else
				exit do
			end if
		loop

		function = true
	  
	end if

	bf.fechar()
   
end function

''''''''
function Efd.carregarCsvNFeDest(bf as bfile, emModoOutrasUFs as boolean) as TDFe ptr
	
	var dfe = new TDFe
	
	dfe->operacao			= ENTRADA
	
	if not emModoOutrasUFs then
		dfe->chave				= bf.charCsv
		dfe->dataEmi			= bf.charCsv
		dfe->nfe.cnpjEmit		= bf.charCsv
		dfe->nfe.nomeEmit		= bf.charCsv
		dfe->nfe.ieEmit			= bf.charCsv
		dfe->nfe.cnpjDest		= bf.charCsv
		dfe->nfe.ufDest			= bf.charCsv
		dfe->nfe.nomeDest		= bf.charCsv
		dfe->nfe.bcICMSTotal	= bf.dblCsv
		dfe->nfe.ICMSTotal		= bf.dblCsv
		dfe->nfe.bcICMSSTTotal	= bf.dblCsv
		dfe->nfe.ICMSSTTotal	= bf.dblCsv
		dfe->nfe.valorNotaTotal	= bf.dblCsv
		dfe->nfe.ufEmit			= bf.charCsv
		dfe->nfe.numero			= bf.intCsv
		dfe->nfe.serie			= bf.intCsv
		dfe->modelo				= bf.intCsv
	else
		dfe->chave				= bf.charCsv
		dfe->nfe.cnpjDest		= bf.charCsv
		dfe->nfe.nomeDest		= bf.charCsv
		dfe->dataEmi			= bf.charCsv
		dfe->nfe.ufDest			= "SP"
		dfe->nfe.cnpjEmit		= bf.charCsv
		dfe->nfe.nomeEmit		= bf.charCsv
		dfe->nfe.ufEmit			= bf.charCsv
		dfe->nfe.bcICMSTotal	= bf.dblCsv
		dfe->nfe.ICMSTotal		= bf.dblCsv
		dfe->nfe.bcICMSSTTotal	= bf.dblCsv
		dfe->nfe.ICMSSTTotal	= bf.dblCsv
		dfe->nfe.valorNotaTotal	= bf.dblCsv
		dfe->modelo				= bf.intCsv
		dfe->nfe.serie			= bf.intCsv
		dfe->nfe.numero			= bf.intCsv
	end if

	dfe->valorOperacao			= dfe->nfe.valorNotaTotal
	
	'' pular \r\n
	bf.char1
	bf.char1
	
	function = dfe
	
end function

''''''''
function Efd.carregarCsvNFeEmit(bf as bfile) as TDFe ptr
	
	var chave = bf.charCsv
	var dfe = cast(TDFe ptr, chaveDFeDict.lookup(chave))	
	if dfe = null then
		dfe = new TDFe
	end if
	
	dfe->operacao			= SAIDA
	dfe->chave				= chave
	dfe->dataEmi			= bf.charCsv
	dfe->nfe.cnpjEmit		= bf.charCsv
	dfe->nfe.nomeEmit		= bf.charCsv
	dfe->nfe.ieEmit			= bf.charCsv
	dfe->nfe.ufEmit			= "SP"
	dfe->nfe.cnpjDest		= bf.charCsv
	dfe->nfe.ufDest			= bf.charCsv
	dfe->nfe.nomeDest		= bf.charCsv
	dfe->nfe.bcICMSTotal	= bf.dblCsv
	dfe->nfe.ICMSTotal		= bf.dblCsv
	dfe->nfe.bcICMSSTTotal	= bf.dblCsv
	dfe->nfe.ICMSSTTotal	= bf.dblCsv
	dfe->nfe.valorNotaTotal	= bf.dblCsv
	bf.charCsv		'' pular "Saída"
	dfe->nfe.numero			= bf.intCsv
	dfe->nfe.serie			= bf.intCsv
	dfe->modelo				= bf.intCsv
	
	dfe->valorOperacao		= dfe->nfe.valorNotaTotal
	
	dfe->nfe.itemListHead	= null
	dfe->nfe.itemListTail	= null

	'' pular \r\n
	bf.char1
	bf.char1
	
	function = dfe
	
end function

''''''''
function Efd.carregarCsvNFeEmitItens(bf as bfile, chave as string) as TDFe_NFeItem ptr
	
	var item = new TDFe_NFeItem
	
	bf.charCsv				'' pular versão
	bf.charCsv				'' pular cnpj emitente
	bf.charCsv				'' pular ie emitente
	bf.charCsv				'' pular cnpj dest
	bf.charCsv				'' pular modelo
	bf.charCsv				'' pular serie
	bf.charCsv				'' pular número
	bf.charCsv				'' pular data emi
	item->cfop				= bf.intCsv
	item->nroItem			= bf.intCsv
	item->codProduto		= bf.charCsv
	item->descricao			= bf.charCsv
	item->qtd				= bf.dblCsv
	item->unidade			= bf.charCsv
	item->valorProduto		= bf.dblCsv
	item->desconto			= bf.dblCsv
	item->despesasAcess		= bf.dblCsv
	item->bcICMS			= bf.dblCsv
	item->aliqICMS			= bf.dblCsv
	item->ICMS				= bf.dblCsv
	item->bcICMSST			= bf.dblCsv
	item->IPI				= bf.dblCsv
	item->next_ = null
	
	chave = bf.charCsv
	
	'' pular \r\n
	bf.char1
	bf.char1
	
	function = item
end function

''''''''
function Efd.carregarCsvCTe(bf as bfile, emModoOutrasUFs as boolean) as TDFe ptr
	var dfe = new TDFe
	
	'' NOTA: só será possível saber se é operacação de entrada ou saída quando pegarmos 
	''       o CNPJ base do contribuinte, que só vem no final do arquivo.......
	dfe->operacao			= DESCONHECIDA			
	
	bf.charCsv				'' pular chave quebrada
	dfe->cte.serie			= bf.intCsv
	dfe->cte.numero			= bf.intCsv
	dfe->cte.cnpjEmit		= bf.charCsv
	dfe->dataEmi			= bf.charCsv
	dfe->cte.nomeEmit		= bf.charCsv
	dfe->cte.ufEmit			= bf.charCsv
	dfe->cte.cnpjToma		= bf.charCsv
	dfe->cte.nomeToma		= bf.charCsv
	dfe->cte.ufToma			= bf.charCsv
	dfe->cte.cnpjRem		= bf.charCsv
	dfe->cte.nomeRem		= bf.charCsv
	dfe->cte.ufRem			= bf.charCsv
	dfe->cte.cnpjDest		= bf.charCsv
	dfe->cte.nomeDest		= bf.charCsv
	dfe->cte.ufDest			= bf.charCsv
	dfe->cte.cnpjExp		= bf.charCsv
	dfe->cte.ufExp			= bf.charCsv
	dfe->cte.cnpjReceb		= bf.charCsv
	dfe->cte.ufReceb		= bf.charCsv
	dfe->cte.tipo			= valint(left(bf.charCsv,1))
	dfe->chave				= bf.charCsv
	dfe->cte.valorPrestacao	= bf.dblCsv
	dfe->cte.valorReceber	= bf.dblCsv
	dfe->cte.qtdCTe			= bf.dblCsv
	dfe->cte.cfop			= bf.intCsv
	dfe->cte.nomeMunicIni	= bf.charCsv
	dfe->cte.ufIni			= bf.charCsv
	dfe->cte.nomeMunicFim	= bf.charCsv
	dfe->cte.ufFim			= bf.charCsv
	dfe->modelo				= 57
	
	dfe->valorOperacao 		= dfe->cte.valorPrestacao
	
	'' pular \r\n
	bf.char1
	bf.char1
	
	''
	if cteListHead = null then
		cteListHead = @dfe->cte
	else
		cteListTail->next_ = @dfe->cte
	end if
	
	cteListTail = @dfe->cte
	dfe->cte.next_ = null
	dfe->cte.parent_ = dfe
	
	function = dfe
	
end function

''''''''
sub Efd.adicionarDFe(dfe as TDFe ptr)
	
	if dfeListHead = null then
		dfeListHead = dfe
	else
		dfeListTail->next_ = dfe
	end if
	
	dfeListTail = dfe
	dfe->next_ = null
	
	if chaveDFeDict.lookup(dfe->chave) = null then
		chaveDFeDict.add(dfe->chave, dfe)
	end if
	
	nroDfe =+ 1

end sub

''''''''
function Efd.carregarCsv(nomeArquivo as String, mostrarProgresso as sub(porCompleto as double)) as Boolean

	dim bf as bfile
   
	if not bf.abrir( nomeArquivo ) then
		return false
	end if
	
	dim as integer tipoArquivo
	if instr( nomeArquivo, "SAFI_NFe_Destinatario" ) > 0 then
		tipoArquivo = SAFI_NFe_Dest
		nfeDestSafiFornecido = true
	
	elseif instr( nomeArquivo, "SAFI_NFe_Emitente_Itens" ) > 0 then
		tipoArquivo = SAFI_NFe_Emit_Itens
		itemNFeSafiFornecido = true
	
	elseif instr( nomeArquivo, "SAFI_NFe_Emitente" ) > 0 then
		tipoArquivo = SAFI_NFe_Emit
		nfeEmitSafiFornecido = true
	
	elseif instr( nomeArquivo, "SAFI_CTe_CNPJ" ) > 0 then
		tipoArquivo = SAFI_CTe
		cteListHead = null
		cteListTail = null
		cteSafiFornecido = true
	else
		print "Erro: impossível resolver tipo de arquivo pelo nome"
		return false
	end if

	var fsize = bf.tamanho

	'' pular header
	pularLinha(bf)
	
	var emModoOutrasUFs = false
	
	do while bf.temProximo()		 
		if mostrarProgresso <> NULL then
			mostrarProgresso(bf.posicao / fsize)
		end if 
		
		'' outro header?
		if bf.peek1 <> asc("""") then
			'' final de arquivo?
			if lcase(left(lerLinha(bf), 22)) = "cnpj base contribuinte" then
				if mostrarProgresso <> NULL then
					mostrarProgresso(1)
				end if 
				
				'' se for CT-e, temos que ler o CNPJ base do contribuinte para fazer um 
				'' patch em todos os tipos de operação (saída ou entrada)
				if tipoArquivo = SAFI_CTe then
					var cnpjBase = bf.charCsv
					var cte = cteListHead
					do while cte <> null 
						if left(cte->cnpjEmit,8) = cnpjBase then
							cte->parent_->operacao = SAIDA
						elseif left(cte->cnpjDest,8) = cnpjBase then
							cte->parent_->operacao = ENTRADA
						end if
						cte = cte->next_
					loop
				end if
				exit do
			else
				emModoOutrasUFs = true
			end if
		end if
	
		select case as const tipoArquivo  
		case SAFI_NFe_Dest
			var dfe = carregarCsvNFeDest( bf, emModoOutrasUFs )
			if dfe <> null then
				adicionarDFe(dfe)
			end if
		
		case SAFI_NFe_Emit
			var dfe = carregarCsvNFeEmit( bf )
			if dfe <> null then
				adicionarDFe(dfe)
			end if
			
		case SAFI_NFe_Emit_Itens
			var chave = ""
			var nfeItem = carregarCsvNFeEmitItens( bf, chave )
			if nfeItem <> null then

				var dfe = cast(TDFe ptr, chaveDFeDict.lookup(chave))
				'' nf-e não encontrada? pode acontecer se processarmos o csv de itens antes do csv de nf-e
				if dfe = null then
					dfe = new TDFe
					'' só adicionar ao dicionário, depois será adicionado por adicionarDFe() no case acima
					dfe->chave = chave
					chaveDFeDict.add(dfe->chave, dfe)
				end if
				
				if dfe->nfe.itemListHead = null then
					dfe->nfe.itemListHead = nfeItem
				else
					dfe->nfe.itemListTail->next_ = nfeItem
				end if
				
				dfe->nfe.itemListTail = nfeItem
			end if
		
		case SAFI_CTe
			var dfe = carregarCsvCTe( bf, emModoOutrasUFs )
			if dfe <> null then
				adicionarDFe(dfe)
			end if
		end select
	loop
   
	bf.fechar()
	
	function = true
	
end function

''''''''
private sub adicionarColunasComuns(sheet as ExcelWorksheet ptr, ehEntrada as Boolean, itemNFeSafiFornecido as boolean)
	
	sheet->AddCellType(CT_STRING, "CNPJ " + iif(ehEntrada, "Emitente", "Destinatario"))
	sheet->AddCellType(CT_STRING, "IE " + iif(ehEntrada, "Emitente", "Destinatario"))
	sheet->AddCellType(CT_STRING, "UF " + iif(ehEntrada, "Emitente", "Destinatario"))
	sheet->AddCellType(CT_STRING, "Razao Social " + iif(ehEntrada, "Emitente", "Destinatario"))
	sheet->AddCellType(CT_INTNUMBER, "Modelo")
	sheet->AddCellType(CT_INTNUMBER, "Serie")
	sheet->AddCellType(CT_INTNUMBER, "Numero")
	sheet->AddCellType(CT_DATE, "Data Emissao")
	sheet->AddCellType(CT_DATE, "Data " + iif(ehEntrada, "Entrada", "Saida"))
	sheet->AddCellType(CT_STRING, "Chave")
	sheet->AddCellType(CT_STRING, "Situacao")

	if ehEntrada or itemNFeSafiFornecido then 
		sheet->AddCellType(CT_MONEY, "BC ICMS")
		sheet->AddCellType(CT_NUMBER, "Aliq ICMS")
		sheet->AddCellType(CT_MONEY, "Valor ICMS")
		sheet->AddCellType(CT_MONEY, "BC ICMS ST")
		sheet->AddCellType(CT_NUMBER, "Aliq ICMS ST")
		sheet->AddCellType(CT_MONEY, "Valor ICMS ST")
		sheet->AddCellType(CT_MONEY, "Valor IPI")
		sheet->AddCellType(CT_MONEY, "Valor Item")
		sheet->AddCellType(CT_INTNUMBER, "Nro Item")
		sheet->AddCellType(CT_NUMBER, "Qtd")
		sheet->AddCellType(CT_STRING, "Unidade")
		sheet->AddCellType(CT_INTNUMBER, "CFOP")
		sheet->AddCellType(CT_INTNUMBER, "CST")
		sheet->AddCellType(CT_INTNUMBER, "NCM")
		sheet->AddCellType(CT_STRING, "Codigo Item")
		sheet->AddCellType(CT_STRING, "Descricao Item")
	else
		if not itemNFeSafiFornecido then
			sheet->AddCellType(CT_MONEY, "BC ICMS")
			sheet->AddCellType(CT_MONEY, "Valor ICMS")
			sheet->AddCellType(CT_MONEY, "BC ICMS ST")
			sheet->AddCellType(CT_MONEY, "Valor ICMS ST")
			sheet->AddCellType(CT_MONEY, "Valor IPI")
			sheet->AddCellType(CT_MONEY, "Valor Total")
		end if
	end if
   
	if not ehEntrada then
		sheet->AddCellType(CT_MONEY, "DifAl FCP")
		sheet->AddCellType(CT_MONEY, "DifAl ICMS Orig")
		sheet->AddCellType(CT_MONEY, "DifAl ICMS Dest")
	end if
end sub
   
''''''''
sub Efd.iniciarExtracao(nomeArquivo as String)
	
	ew = new ExcelWriter
	ew->create(nomeArquivo)

	entradas = null
	saidas = null
	naoEscrituradas = null

end sub

''''''''
sub Efd.finalizarExtracao(mostrarProgresso as sub(porCompleto as double))

	ew->Flush(mostrarProgresso)

	ew->Close
	
	delete ew
	ew = null
   
end sub

''''''''
#define STR2DATA(s) (iif(len(s)<8, "1900-01-01T00:00:00.000", mid(s,5) + "-" + mid(s,3,2) + "-" + mid(s,1,2) + "T00:00:00.000"))

''''''''
#define STR2DATABR(s) (mid(s,1,2) + "/" + mid(s,3,2) + "/" + mid(s,5))

''''''''
#define STRSINT2DATA(s) (mid(s,1,4) + "-" + mid(s,5,2) + "-" + mid(s,7,2) + "T00:00:00.000")

''''''''
#define MUNICIPIO2SIGLA(m) (iif(m >= 1100000 and m <= 5399999, codUF2Sigla(m \ 100000), "EX"))

''''''''
sub Efd.adicionarEfdDfe(chave as zstring ptr, operacao as TipoOperacao, dataEmi as zstring ptr, valorOperacao as double)
	
	if len(chave) = 0 then
		return
	end if
	
	if efdDFeDict.lookup(chave) = null then
		var ed = new TEfd_Dfe
		ed->chave = *chave
		ed->operacao = operacao
		ed->dataEmi = *dataEmi
		ed->valorOperacao = valorOperacao
		if efdDfeListHead = null then
			efdDfeListHead = ed
		else
			efdDfeListTail->next_ = ed
		end if
		efdDfeListTail = ed
		ed->next_ = null
		
		efdDFeDict.add(ed->chave, @ed)
	end if
	
end sub

''''''''
sub Efd.criarPlanilhas()
	'' planilha de entradas
	entradas = ew->AddWorksheet("Entradas")
	adicionarColunasComuns(entradas, true, itemNFeSafiFornecido)

	'' planilha de saídas
	saidas = ew->AddWorksheet("Saidas")
	adicionarColunasComuns(saidas, false, itemNFeSafiFornecido)

	'' apuração do ICMS
	apuracaoIcms = ew->AddWorksheet("Apuracao ICMS")
	apuracaoIcms->AddCellType(CT_DATE, "Inicio")
	apuracaoIcms->AddCellType(CT_DATE, "Fim")
	apuracaoIcms->AddCellType(CT_MONEY, "Total Debitos")
	apuracaoIcms->AddCellType(CT_MONEY, "Ajustes Debitos")
	apuracaoIcms->AddCellType(CT_MONEY, "Total Ajuste Deb")
	apuracaoIcms->AddCellType(CT_MONEY, "Estornos Credito")
	apuracaoIcms->AddCellType(CT_MONEY, "Total Creditos")
	apuracaoIcms->AddCellType(CT_MONEY, "Ajustes Creditos")
	apuracaoIcms->AddCellType(CT_MONEY, "Total Ajuste Cred")
	apuracaoIcms->AddCellType(CT_MONEY, "Estornos Debito")
	apuracaoIcms->AddCellType(CT_MONEY, "Saldo Cred Anterior")
	apuracaoIcms->AddCellType(CT_MONEY, "Saldo Devedor Apurado")
	apuracaoIcms->AddCellType(CT_MONEY, "Total Deducoes")
	apuracaoIcms->AddCellType(CT_MONEY, "ICMS a Recolher")
	apuracaoIcms->AddCellType(CT_MONEY, "Saldo Credor a Transportar")
	apuracaoIcms->AddCellType(CT_MONEY, "Deb Extra Apuracao")
   
	'' apuração do ICMS ST
	apuracaoIcmsST = ew->AddWorksheet("Apuracao ICMS ST")
	apuracaoIcmsST->AddCellType(CT_DATE, "Inicio")
	apuracaoIcmsST->AddCellType(CT_DATE, "Fim")
	apuracaoIcmsST->AddCellType(CT_STRING, "UF")
	apuracaoIcmsST->AddCellType(CT_STRING, "Movimentacao")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Saldo Credor Anterior")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Total Devolucao Merc")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Total Ressarcimentos")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Total Ajustes Cred")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Total Ajustes Cred Docs")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Total Retencao")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Total Ajustes Deb")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Total Ajustes Deb Docs")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Saldo Devedor ant. Deducoes")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Total Deducoes")
	apuracaoIcmsST->AddCellType(CT_MONEY, "ICMS a Recolher")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Saldo Credor a Transportar")
	apuracaoIcmsST->AddCellType(CT_MONEY, "Deb Extra Apuracao")
			
end sub

''''''''
type HashCtx
	bf				as bfile ptr
	tamanhoSemSign	as longint
	bytesLidosTotal	as longint
end type

private function hashReadCB cdecl(ctx_ as any ptr, buffer as ubyte ptr, maxLen as integer) as integer
	var ctx = cast(HashCtx ptr, ctx_)
	
	if ctx->bytesLidosTotal + maxLen > ctx->tamanhoSemSign then
		maxLen = ctx->tamanhoSemSign - ctx->bytesLidosTotal
	end if
	
	var bytesLidos = ctx->bf->ler(buffer, maxLen)
	ctx->bytesLidosTotal += bytesLidos
	
	function = bytesLidos
	
end function

''''''''
function lerInfoAssinatura(nomeArquivo as string, assinaturaP7K_DER() as byte) as InfoAssinatura ptr
	
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
	
	s = sh->Compute_SHA1(@hashReadCB, ctx)
	if s <> null then
		res->hashDoArquivo = *s
		deallocate s
	end if
	
	bf->fechar()

	''
	sh->Free(p7k)
	delete sh
	
	function = res
	
end function

''''''''
function Efd.processar(nomeArquivo as string, mostrarProgresso as sub(porCompleto as double), gerarPDF as boolean) as Boolean
   
	gerarPlanilhas(nomeArquivo)
	
	if gerarPDF then
		if tipoArquivo = TIPO_ARQUIVO_EFD then
			infAssinatura = lerInfoAssinatura(nomeArquivo, assinaturaP7K_DER())
		
			gerarRelatorios(nomeArquivo)
			
			if infAssinatura <> NULL then
				delete infAssinatura
			end if
		end if
	end if
	
	do while regListHead <> null
		var next_ = regListHead->next_
		delete regListHead
		regListHead = next_
	loop

	regListHead = null
	regListTail = null

	sintegraDict.end_()
	itemIdDict.end_()
	participanteDict.end_()

	function = true
end function

''''''''
sub Efd.gerarPlanilhas(nomeArquivo as string)
	
	if entradas = null then
		criarPlanilhas
	end if
	
	var reg = regListHead
	do while reg <> null
		'para cada registro..
		select case reg->tipo
		'item de NF-e?
		case DOC_NFE_ITEM
			var doc = reg->itemNFe.documentoPai
			select case as const doc->situacao
			case REGULAR, EXTEMPORANEO
				'só existe item para entradas
				if doc->operacao = ENTRADA then
					if len(doc->chave) > 0 then
						adicionarEfdDfe(doc->chave, doc->operacao, doc->dataEmi, doc->valorTotal)
					end if

					var row = entradas->AddRow()

					var part = cast( TParticipante ptr, participanteDict.lookup(doc->idParticipante) )
					if part <> null then
						row->addCell(part->cnpj)
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
					row->addCell(STR2DATA(doc->dataEmi))
					row->addCell(STR2DATA(doc->dataEntSaida))
					row->addCell(doc->chave)
					row->addCell(situacao2String(doc->situacao))
					row->addCell(reg->itemNFe.bcICMS)
					row->addCell(reg->itemNFe.aliqICMS)
					row->addCell(reg->itemNFe.ICMS)
					row->addCell(reg->itemNFe.bcICMSST)
					row->addCell(reg->itemNFe.aliqICMSST)
					row->addCell(reg->itemNFe.ICMSST)
					row->addCell(reg->itemNFe.IPI)
					row->addCell(reg->itemNFe.valor)
					row->addCell(reg->itemNFe.numItem)
					row->addCell(reg->itemNFe.qtd)
					row->addCell(reg->itemNFe.unidade)
					row->addCell(reg->itemNFe.cfop)
					row->addCell(reg->itemNFe.cstICMS)
					var itemId = cast( TItemId ptr, itemIdDict.lookup(reg->itemNFe.itemId) )
					if itemId <> null then 
						row->addCell(itemId->ncm)
						row->addCell(itemId->id)
						row->addCell(itemId->descricao)
					else
						row->addCell("")
						row->addCell("")
						row->addCell("")
					end if
				end if
			end select

		'NF-e?
		case DOC_NFE
			select case as const reg->nfe.situacao
			case REGULAR, EXTEMPORANEO
				if len(reg->nfe.chave) > 0 then
					adicionarEfdDfe(reg->nfe.chave, reg->nfe.operacao, reg->nfe.dataEmi, reg->nfe.valorTotal)
				end if

				'' NOTA: não existe itemDoc para saídas, só temos informação básica do DF-e, 
				'' 	     a não ser que sejam carregados os relatórios .csv do SAFI vindos do infoview
				if reg->nfe.operacao = SAIDA then
					dim as TDFe_NFeItem ptr item = null
					if itemNFeSafiFornecido then
						if len(reg->nfe.chave) > 0 then
							var dfe = cast( TDFe ptr, chaveDFeDict.lookup(reg->nfe.chave) )
							if dfe <> null then
								item = dfe->nfe.itemListHead
							end if
						end if
					end if

					var part = cast( TParticipante ptr, participanteDict.lookup(reg->nfe.idParticipante) )

					do
						var row = saidas->AddRow()
						if part <> null then
							row->addCell(part->cnpj)
							row->addCell(part->ie)
							row->addCell(MUNICIPIO2SIGLA(part->municip))
							row->addCell(part->nome)
						else
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell("")
						end if
						row->addCell(reg->nfe.modelo)
						row->addCell(reg->nfe.serie)
						row->addCell(reg->nfe.numero)
						row->addCell(STR2DATA(reg->nfe.dataEmi))
						row->addCell(STR2DATA(reg->nfe.dataEntSaida))
						row->addCell(reg->nfe.chave)
						row->addCell(situacao2String(reg->nfe.situacao))

						if itemNFeSafiFornecido then
							if item <> null then
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
								row->addCell("")
								row->addCell("")
								row->addCell(item->codProduto)
								row->addCell(item->descricao)
							else
								for cell as integer = 1 to 13
									row->addCell("")
								next
							end if

						else
							row->addCell(reg->nfe.bcICMS)
							row->addCell(reg->nfe.ICMS)
							row->addCell(reg->nfe.bcICMSST)
							row->addCell(reg->nfe.ICMSST)
							row->addCell(reg->nfe.IPI)
							row->addCell(reg->nfe.valorTotal)
						end if

						row->addCell(reg->nfe.difal.fcp)
						row->addCell(reg->nfe.difal.icmsOrigem)
						row->addCell(reg->nfe.difal.icmsDest)
						
						if item = null then
							exit do
						end if

						item = item->next_
					loop while item <> null
				end if
		   
			case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
				if reg->nfe.operacao = SAIDA then
					var row = saidas->AddRow()

					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell(reg->nfe.modelo)
					row->addCell(reg->nfe.serie)
					row->addCell(reg->nfe.numero)
					'' NOTA: cancelados e inutilizados não vêm com a data preenchida, então retiramos a data da chave ou do registro mestre
					var dataEmi = iif( len(reg->nfe.chave) = 44, "01" + mid(reg->nfe.chave,5,2) + "20" + mid(reg->nfe.chave,3,2), regListHead->mestre.dataIni )
					if len(reg->nfe.chave) > 0 then
						adicionarEfdDfe(reg->nfe.chave, reg->nfe.operacao, dataEmi, 0)
					end if
					
					row->addCell(STR2DATA(dataEmi))
					row->addCell("")
					row->addCell(reg->nfe.chave)
					row->addCell(situacao2String(reg->nfe.situacao))
				end if

			end select

		'CT-e?
		case DOC_CTE
			select case as const reg->cte.situacao
			case REGULAR, EXTEMPORANEO
				if len(reg->cte.chave) > 0 then
					adicionarEfdDfe(reg->cte.chave, reg->cte.operacao, reg->cte.dataEmi, reg->cte.valorServico)
				end if
				
				var part = cast( TParticipante ptr, participanteDict.lookup(reg->cte.idParticipante) )

				dim as TDFe_CTe ptr cte = null
				if cteSafiFornecido then
					if len(reg->cte.chave) > 0 then
						var dfe = cast( TDFe ptr, chaveDFeDict.lookup(reg->cte.chave) )
						if dfe <> null then
							cte = @dfe->cte
						end if
					end if
				end if
				
				dim as TDocItemAnal ptr item = null
				if reg->cte.operacao = ENTRADA then
					item = reg->cte.itemAnalListHead
				end if
				
				var itemCnt = 1
				do
					dim as ExcelRow ptr row 
					if reg->cte.operacao = SAIDA then
						row = saidas->AddRow()
					else
						row = entradas->AddRow()
					end if
					
					if part <> null then
						row->addCell(part->cnpj)
						row->addCell(part->ie)
						row->addCell(MUNICIPIO2SIGLA(part->municip))
						row->addCell(part->nome)
					else
						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell("")
					end if
					row->addCell(reg->cte.modelo)
					row->addCell(reg->cte.serie)
					row->addCell(reg->cte.numero)
					row->addCell(STR2DATA(reg->cte.dataEmi))
					row->addCell(STR2DATA(reg->cte.dataEntSaida))
					row->addCell(reg->cte.chave)
					row->addCell(situacao2String(reg->cte.situacao))
					
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
						if reg->cte.operacao = ENTRADA or cint(itemNFeSafiFornecido) then
							row->addCell(reg->cte.bcICMS)
							row->addCell("")
							row->addCell(reg->cte.ICMS)
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell(reg->cte.valorServico)
							row->addCell(1)
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell("")
						else
							row->addCell(reg->cte.bcICMS)
							row->addCell(reg->cte.ICMS)
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell(reg->cte.valorServico)
						end if
						
					end if

					if reg->cte.operacao = SAIDA then
						row->addCell(reg->cte.difal.fcp)
						row->addCell(reg->cte.difal.icmsOrigem)
						row->addCell(reg->cte.difal.icmsDest)
					end if
					
				loop while item <> null
			
			case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
				if reg->cte.operacao = SAIDA then
					var row = saidas->AddRow()

					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell(reg->cte.modelo)
					row->addCell(reg->cte.serie)
					row->addCell(reg->cte.numero)
					'' NOTA: cancelados e inutilizados não vêm com a data preenchida, então retiramos a data da chave ou do registro mestre
					var dataEmi = iif( len(reg->cte.chave) = 44, "01" + mid(reg->cte.chave,5,2) + "20" + mid(reg->cte.chave,3,2), regListHead->mestre.dataIni )
					if len(reg->cte.chave) > 0 then
						adicionarEfdDfe(reg->cte.chave, reg->cte.operacao, dataEmi, 0)
					end if
					row->addCell(STR2DATA(dataEmi))
					row->addCell("")
					row->addCell(reg->cte.chave)
					row->addCell(situacao2String(reg->cte.situacao))
				end if
			
			end select
			
		case APURACAO_ICMS_PERIODO
			var row = apuracaoIcms->AddRow()

			row->addCell(STR2DATA(reg->apuIcms.dataIni))
			row->addCell(STR2DATA(reg->apuIcms.dataFim))
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
			
		case APURACAO_ICMS_ST_PERIODO
			var row = apuracaoIcmsST->AddRow()

			row->addCell(STR2DATA(reg->apuIcmsST.dataIni))
			row->addCell(STR2DATA(reg->apuIcmsST.dataFim))
			row->addCell(reg->apuIcmsST.UF)
			row->addCell(iif(reg->apuIcmsST.mov=0, "N", "S"))
			row->addCell(reg->apuIcmsST.saldoCredAnterior)
			row->addCell(reg->apuIcmsST.devolMercadorias)
			row->addCell(reg->apuIcmsST.totalRessarciment)
			row->addCell(reg->apuIcmsST.totalOutrosCred)
			row->addCell(reg->apuIcmsST.ajusteCred)
			row->addCell(reg->apuIcmsST.totalRetencao)
			row->addCell(reg->apuIcmsST.totalOutrosDeb)
			row->addCell(reg->apuIcmsST.ajusteDeb)
			row->addCell(reg->apuIcmsST.saldoAntesDed)
			row->addCell(reg->apuIcmsST.totalDeducoes)
			row->addCell(reg->apuIcmsST.icmsRecolher)
			row->addCell(reg->apuIcmsST.saldoCredTransportar)
			row->addCell(reg->apuIcmsST.debExtraApuracao)

		'documento do sintegra?
		case SINTEGRA_DOCUMENTO
			if reg->docSint.modelo = 55 then 
				select case as const reg->docSint.situacao
				case REGULAR, EXTEMPORANEO
					dim as ExcelRow ptr row 
					if reg->docSint.operacao = SAIDA then
						row = saidas->AddRow()
					else
						row = entradas->AddRow()
					end if
					  
					row->addCell(reg->docSint.cnpj)
					row->addCell(reg->docSint.ie)
					row->addCell(reg->docSint.uf)
					row->addCell("")
					row->addCell(reg->docSint.modelo)
					row->addCell(reg->docSint.serie)
					row->addCell(reg->docSint.numero)
					row->addCell(STRSINT2DATA(reg->docSint.dataEmi))
					row->addCell("")
					row->addCell("")
					row->addCell(situacao2String(reg->docSint.situacao))
					row->addCell(reg->docSint.bcICMS)
					if reg->docSint.operacao = ENTRADA or cint(itemNFeSafiFornecido) then
						row->addCell(reg->docSint.aliqICMS)
					end if
					row->addCell(reg->docSint.ICMS)
					row->addCell(reg->docSint.bcICMSST)
					if reg->docSint.operacao = ENTRADA or cint(itemNFeSafiFornecido) then
						row->addCell("")
					end if
					row->addCell(reg->docSint.ICMSST)
					row->addCell(reg->docSint.valorIPI)
					row->addCell(reg->docSint.valorTotal)
					if reg->docSint.operacao = ENTRADA or cint(itemNFeSafiFornecido) then
						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell(reg->docSint.cfop)
						row->addCell("")
					end if
					
				case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
					/'var row = canceladas->AddRow()
					row->addCell(reg->docSint.modelo)
					row->addCell(reg->docSint.serie)
					row->addCell(reg->docSint.numero)
					row->addCell("")'/

				end select
			end if

		end select

		reg = reg->next_
	loop
end sub

''''''''
sub Efd.gerarRelatorios(nomeArquivo as string)
	
	ultimoRelatorio = -1
	
	'' NOTA: por limitação do DocxFactory, que só consegue trabalhar com um template por vez, 
	''		 precisar processar entradas primeiro, depois saídas e por últimos os registros 
	''		 que são sequenciais (LRE e LRS vêm intercalados na EFD)
	
	'' LRE
	iniciarRelatorio(REL_LRE, "entradas", "LRE")
	
	var reg = regListHead
	do while reg <> null
		'para cada registro..
		select case reg->tipo
		'NF-e?
		case DOC_NFE
			if reg->nfe.operacao = ENTRADA then
				var part = cast( TParticipante ptr, participanteDict.lookup(reg->nfe.idParticipante) )
				adicionarDocRelatorioEntradas(@reg->nfe, part)
			end if
		
		'CT-e?
		case DOC_CTE
			if reg->cte.operacao = ENTRADA then
				var part = cast( TParticipante ptr, participanteDict.lookup(reg->cte.idParticipante) )
				adicionarDocRelatorioEntradas(@reg->cte, part)
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
		case DOC_NFE
			if reg->nfe.operacao = SAIDA then
				var part = cast( TParticipante ptr, participanteDict.lookup(reg->nfe.idParticipante) )
				adicionarDocRelatorioSaidas(@reg->nfe, part)
			end if

		'CT-e?
		case DOC_CTE
			if reg->cte.operacao = SAIDA then
				var part = cast( TParticipante ptr, participanteDict.lookup(reg->cte.idParticipante) )
				adicionarDocRelatorioSaidas(@reg->cte, part)
			end if
			
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

end sub

''''''''
function STRDFE2DATA(s as zstring ptr) as string 
	''         0123456789
	var res = "0000-00-00T00:00:00.000"
	
	var p = 0
	if s[0+1] = asc("/") then
		res[9] = s[0]
		p += 1+1
	else
		res[8] = s[0]
		res[9] = s[1]
		p += 2+1
	end if

	if s[p+1] = asc("/") then
		res[6] = s[p]
		p += 1+1
	else
		res[5] = s[p]
		res[6] = s[p+1]
		p += 2+1
	end if
	
	res[0] = s[p]
	res[1] = s[p+1]
	res[2] = s[p+2]
	res[3] = s[p+3]
	
	function = res
end function

''''''''
sub Efd.analisar(mostrarProgresso as sub(porCompleto as double)) 

	naoEscrituradas = ew->AddWorksheet("Nao Escrituradas")
	naoEscrituradas->AddCellType(CT_STRING, "Chave")
	naoEscrituradas->AddCellType(CT_DATE, "Data")
	naoEscrituradas->AddCellType(CT_INTNUMBER, "Modelo")
	naoEscrituradas->AddCellType(CT_INTNUMBER, "Serie")
	naoEscrituradas->AddCellType(CT_INTNUMBER, "Numero")
	naoEscrituradas->AddCellType(CT_STRING, "Operacao")
	naoEscrituradas->AddCellType(CT_MONEY, "Valor Operacao")
	
	var i = 0
	var dfe = dfeListHead
	if dfeListHead = null then
		var row = naoEscrituradas->AddRow()
		row->addCell("Nao foi possivel verificar falta de escrituracao porque os relatorios do SAFI nao foram fornecidos")
	else
		do while dfe <> null
			if efdDFeDict.lookup(dfe->chave) = null then
				var row = naoEscrituradas->AddRow()
				row->addCell(dfe->chave)
				row->addCell(STRDFE2DATA(dfe->dataEmi))
				row->addCell(dfe->modelo)
				row->addCell(mid(dfe->chave, 23, 3))
				row->addCell(mid(dfe->chave, 23+3, 9))
				row->addCell(iif(dfe->operacao < 2, iif(dfe->operacao = 0, "E", "S"), "-"))
				row->addCell(dfe->valorOperacao)
			end if

			i += 1
			if mostrarProgresso <> NULL then
				'mostrarProgresso(i / nroDFe)
			end if 
			
			dfe = dfe->next_
		loop
	end if

end sub

''''''''
function STR2IE(ie as string) as string
	var ie2 = right(string(12,"0") + ie, 12)
	function = left(ie2,3) + "." + mid(ie2,4,3) + "." + mid(ie2,4+3,3) + "." + right(ie2,3)
end function

''''''''
#define STR2CNPJ(s) (left(s,2) + "." + mid(s,3,3) + "." + mid(s,3+3,3) + "/" + mid(s,3+3+3,4) + "-" + right(s,2))

#define STR2CPF(s) (left(s,3) + "." + mid(s,4,3) + "." + mid(s,4+3,3) + "-" + right(s,2))

#define STR2YYYYMM(s) (mid(s,5) + "-" + mid(s,3,2))

#define DBL2MONEYBR(d) (format(d,"#,#,#.00"))

''''''''
sub Efd.gerarRelatorioApuracaoICMS(nomeArquivo as string, reg as TRegistro ptr)

	iniciarRelatorio(REL_RAICMS, "apuracao_icms", "RAICMS")
	
	dfwd->setClipboardValueByStrW("grid", "nome", regListHead->mestre.nome)
	dfwd->setClipboardValueByStr("grid", "cnpj", STR2CNPJ(regListHead->mestre.cnpj))
	dfwd->setClipboardValueByStr("grid", "ie", STR2IE(regListHead->mestre.ie))
	dfwd->setClipboardValueByStr("grid", "escrit", STR2DATABR(regListHead->mestre.dataIni) + " a " + STR2DATABR(regListHead->mestre.dataFim))
	dfwd->setClipboardValueByStr("grid", "apur", STR2DATABR(reg->apuIcms.dataIni) + " a " + STR2DATABR(reg->apuIcms.dataFim))
	
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
	dfwd->setClipboardValueByStr("grid", "escrit", STR2DATABR(regListHead->mestre.dataIni) + " a " + STR2DATABR(regListHead->mestre.dataFim))
	dfwd->setClipboardValueByStrW("grid", "apur", STR2DATABR(reg->apuIcmsST.dataIni) + " a " + STR2DATABR(reg->apuIcmsST.dataFim) + " - INSCRIÇÃO ESTADUAL:")
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

	dfwd->load(baseTemplatesDir + nomeRelatorio + ".dfw")

	dfwd->setClipboardValueByStrW("_header", "nome", regListHead->mestre.nome)
	dfwd->setClipboardValueByStr("_header", "cnpj", STR2CNPJ(regListHead->mestre.cnpj))
	dfwd->setClipboardValueByStr("_header", "ie", STR2IE(regListHead->mestre.ie))
	dfwd->setClipboardValueByStr("_header", "uf", MUNICIPIO2SIGLA(regListHead->mestre.municip))
	
	select case relatorio
	case REL_LRE, REL_LRS
		dfwd->setClipboardValueByStrW("_header", "municipio", codMunicipio2Nome(regListHead->mestre.municip))
		dfwd->setClipboardValueByStr("_header", "apu", STR2DATABR(regListHead->mestre.dataIni) + " a " + STR2DATABR(regListHead->mestre.dataFim))
	
		relSomaLRList.init(10, len(RelSomatorioLR))
		relSomaLRHash.init(10)
		
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
	
	var soma = cast(RelSomatorioLR ptr, relSomaLRHash.lookUp(chave))
	if soma = null then
		soma = relSomaLRList.addOrdAsc(strptr(chave), @cmpFunc)
		soma->chave = chave
		soma->situacao = sit
		soma->cst = anal->cst
		soma->cfop = anal->cfop
		soma->aliq = anal->aliq
		relSomaLRHash.add(soma->chave, soma)
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
sub Efd.adicionarDocRelatorioSaidas(doc as TDocDFe ptr, part as TParticipante ptr)

	if len(doc->dataEmi) > 0 then
		dfwd->setClipboardValueByStr("linha", "demi", STR2DATABR(doc->dataEmi))
	end if
	if len(doc->dataEntSaida) > 0 then
		dfwd->setClipboardValueByStr("linha", "dsaida", STR2DATABR(doc->dataEntSaida))
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
	
end sub

''''''''
sub Efd.adicionarDocRelatorioEntradas(doc as TDocDFe ptr, part as TParticipante ptr)

	dfwd->setClipboardValueByStr("linha", "demi", STR2DATABR(doc->dataEmi))
	dfwd->setClipboardValueByStr("linha", "dent", STR2DATABR(doc->dataEntSaida))
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
		
		relSomaLRHash.end_()
		relSomaLRList.end_()
	case else
		dfwd->paste("ass")
	end select
	
	dfwd->save(STR2YYYYMM(regListHead->mestre.dataIni) + "_" + ultimoRelatorioSufixo + ".docx")
	
	dfwd->close()
	
	ultimoRelatorio = -1

end sub
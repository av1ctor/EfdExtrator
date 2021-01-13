#include once "EfdExt.bi"
#include once "EfdSintegraImport.bi"
#include once "libs/trycatch.bi"

''''''''
constructor EfdSintegraImport(opcoes as OpcoesExtracao ptr)
	base(opcoes)
	sintegraDict = new TDict(2^10)
end constructor

''''''''
destructor EfdSintegraImport()
	delete sintegraDict
end destructor

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
private function lerRegSintegraDocumento(bf as bfile) as TDocumentoSintegra ptr

	var reg = new TDocumentoSintegra

	reg->cnpj 		= bf.nchar(14)
	reg->ie 		= bf.nchar(14)
	reg->dataEmi 	= bf.char8
	reg->uf 		= UF_SIGLA2COD(bf.char2)
	reg->modelo 	= bf.int2
	reg->serie 		= bf.nchar(3)
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("¨") then
		bf.char1
		reg->numero = bf.int9
	else
		reg->numero = bf.int6
	end if
	reg->cfop 		= bf.int4
	reg->operacao 	= iif( bf.char1 = asc("T"), ENTRADA, SAIDA )
	reg->valorTotal = bf.dbl13_2
	reg->bcICMS 	= bf.dbl13_2
	reg->ICMS 		= bf.dbl13_2
	reg->valorIsento= bf.dbl13_2
	reg->valorOutras= bf.dbl13_2
	reg->aliqICMS 	= bf.dbl4_2
	reg->situacao 	= situacaoSintegra2SituacaoEfd( bf.char1 )

	'' ler chave NF-e no final da linha, se for um sintegra convertido pelo SAFI
	if bf.peek1 <> 13 then
		reg->chave 	= bf.nchar(44)
	end if

	'pular \r\n
	bf.char1
	bf.char1

	return reg
end function

''''''''
private function lerRegSintegraDocumentoST(bf as bfile) as TDocumentoSintegra ptr

	var reg = new TDocumentoSintegra

	reg->cnpj 		= bf.nchar(14)
	reg->ie 		= bf.nchar(14)
	reg->dataEmi	= bf.char8
	reg->uf 		= UF_SIGLA2COD(bf.char2)
	reg->modelo 	= bf.int2
	reg->serie 		= bf.nchar(3)
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("¨") then
		bf.char1
		reg->numero = bf.int9
	else
		reg->numero = bf.int6
	end if
	reg->cfop 		= bf.int4
	reg->operacao 	= iif( bf.char1 = asc("T"), ENTRADA, SAIDA )
	reg->bcICMSST 	= bf.dbl13_2
	reg->ICMSST 	= bf.dbl13_2
	reg->despesasAcess = bf.dbl13_2
	reg->situacao 	= situacaoSintegra2SituacaoEfd( bf.char1 )
	bf.nchar(30)

	'pular \r\n
	bf.char1
	bf.char1

	return reg
end function

''''''''
private function lerRegSintegraDocumentoIPI(bf as bfile) as TDocumentoSintegra ptr

	var reg = new TDocumentoSintegra

	reg->cnpj 		= bf.nchar(14)
	reg->ie 		= bf.nchar(14)
	reg->dataEmi 	= bf.char8
	reg->uf 		= UF_SIGLA2COD(bf.char2)
	reg->serie 		= bf.nchar(3)
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("¨") then
		bf.char1
		reg->numero = bf.int9
	else
		reg->numero = bf.int6
	end if
	reg->cfop 		= bf.int4
	reg->valorTotal = bf.dbl13_2
	reg->valorIPI 	= bf.dbl13_2
	reg->valorIsentoIPI = bf.dbl13_2
	reg->valorOutrasIPI = bf.dbl13_2
	bf.nchar(1+20)

	'pular \r\n
	bf.char1
	bf.char1

	return reg
end function

''''''''
private function lerRegSintegraMercadoria(bf as bfile) as TItemId ptr

	var reg = new TItemId

	bf.nchar(8+8)
	reg->id			  	= bf.nchar(14)
	reg->ncm			= vallng(bf.nchar(8))
	reg->descricao	  	= bf.nchar(53)
	reg->unidInventario = bf.nchar(6)
	reg->aliqIPI		= bf.dbl5_2
	reg->aliqICMSInt	= bf.dbl4_2
	reg->redBcICMS	  	= bf.dbl5_2
	reg->bcICMSST	  	= bf.dbl13_2

	'pular \r\n
	bf.char1
	bf.char1

	return reg
end function

''''''''
private function lerRegSintegraDocumentoItem(bf as bfile) as TDocumentoItemSintegra ptr

	var reg = new TDocumentoItemSintegra
	
	reg->cnpj 		= bf.nchar(14)
	bf.nchar(2)
	reg->serie 		= bf.nchar(3)
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("¨") then
		bf.char1
		reg->numero = bf.int9
	else
		reg->numero = bf.int6
	end if
	reg->cfop 		= bf.int4
	reg->CST 		= bf.nchar(3)
	reg->nroItem	= valint(bf.nchar(3))	
	reg->codMercadoria = bf.nchar(14)
	reg->qtd		= bf.dbl11_3
	reg->valor		= bf.dbl12_2
	reg->desconto	= bf.dbl12_2
	reg->bcICMS		= bf.dbl12_2
	reg->bcICMSST	= bf.dbl12_2
	reg->valorIPI	= bf.dbl12_2
	reg->aliqICMS	= bf.dbl4_2
	
	'pular \r\n
	bf.char1
	bf.char1

	return reg
end function

#define GENSINTEGRAKEY(r) ((r)->cnpj + (r)->serie + str((r)->numero) + str((r)->cfop))
  
''''''''
function EfdSintegraImport.lerRegistroSintegra(bf as bfile) as TRegistro ptr

	var reg = cast(TRegistro ptr, null)

	var tipo = bf.int2
	select case as const tipo
	case SINTEGRA_DOCUMENTO
		var node = lerRegSintegraDocumento(bf)
		if node = null then
			return null
		end if
		reg = node

		'adicionar ao dicionário
		node->chaveDict = GENSINTEGRAKEY(node)
		var antReg = cast(TDocumentoSintegra ptr, sintegraDict->lookup(node->chaveDict))
		if antReg = null then
			sintegraDict->add(node->chaveDict, node)
			node->tipo = SINTEGRA_DOCUMENTO
		else
			'' para cada alíquota diferente há um novo registro 50, mas nós só queremos os valores totais
			''antReg->valorTotal	+= node->valorTotal
			''antReg->bcICMS		+= node->bcICMS
			''antReg->ICMS		+= node->ICMS
			''antReg->valorIsento += node->valorIsento
			''antReg->valorOutras += node->valorOutras
			node->tipo = DESCONHECIDO 
		end if

	case SINTEGRA_DOCUMENTO_ST
		var node = lerRegSintegraDocumentoST(bf)
		if node = null then
			return null
		end if
		reg = node

		node->chaveDict = GENSINTEGRAKEY(node)
		var antReg = cast(TDocumentoSintegra ptr, sintegraDict->lookup(node->chaveDict))
		'' NOTA: pode existir registro 53 sem o correspondente 50, para quando só há ICMS ST, sem destaque ICMS próprio
		if antReg = null then
			sintegraDict->add(node->chaveDict, node)
			node->tipo = SINTEGRA_DOCUMENTO
		else
			''antReg->bcICMSST		+= node->bcICMSST
			''antReg->ICMSST		+= node->ICMSST
			''antReg->despesasAcess	+= reg->despesasAcess
			node->tipo = DESCONHECIDO
		end if
	  
	case SINTEGRA_DOCUMENTO_IPI
		var node = lerRegSintegraDocumentoIPI(bf)
		if node = null then
			return null
		end if

		node->chaveDict = GENSINTEGRAKEY(node)
		var antReg = cast(TDocumentoSintegra ptr, sintegraDict->lookup(node->chaveDict))
		if antReg = null then
			onError("Sintegra 53 sem 50: " & node->chaveDict)
		else
			antReg->valorIPI		= node->valorIPI
			antReg->valorIsentoIPI	= node->valorIsentoIPI
			antReg->valorOutrasIPI	= node->valorOutrasIPI
		end if

		return null
		
	case SINTEGRA_DOCUMENTO_ITEM
		var node = lerRegSintegraDocumentoItem(bf)
		if node = null then
			return null
		end if
		reg = node

		var chaveDict = GENSINTEGRAKEY(node)
		var doc = cast(TDocumentoSintegra ptr, sintegraDict->lookup(chaveDict))
		if doc = null then
			onError("Sintegra 54 sem 50: " & chaveDict)
		end if
		
		node->doc = doc
		node->tipo = SINTEGRA_DOCUMENTO_ITEM
		
	case SINTEGRA_MERCADORIA
		var node = lerRegSintegraMercadoria(bf)
		if node = null then
			return null
		end if
		reg = node
		reg->tipo = ITEM_ID

		'adicionar ao dicionário
		if itemIdDict->lookup(node->id) = null then
			itemIdDict->add(node->id, node)
		end if
		
	case else
		pularLinha(bf)
		return null
	end select

	return reg

end function

''''''''
function EfdSintegraImport.carregar(nomeArquivo as string) as boolean
	
	dim bf as bfile
   
	if not bf.abrir( nomeArquivo ) then
		return false
	end if

	tipoArquivo = TIPO_ARQUIVO_SINTEGRA
	regListHead = null
	nroRegs = 0
	
	var fsize = bf.tamanho
	
	dim as TRegistro ptr tail = null
	nroLinha = 0

	try
		do while bf.temProximo()		 
			nroLinha += 1
			var reg = lerRegistroSintegra(bf)
			if reg <> null then 
				if not onProgress(null, bf.posicao / fsize) then
					exit do
				end if
				
				if reg->tipo <> DESCONHECIDO then
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
		loop
	catch
		onError(!"\r\nErro ao carregar o registro da linha (" & nroLinha & !") do arquivo\r\n")
	endtry
	   
	onProgress(null, 1)

	function = true
  
	bf.fechar()

end function

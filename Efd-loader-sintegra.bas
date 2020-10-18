#include once "efd.bi"
#include once "bfile.bi"
#include once "trycatch.bi"

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
	reg->docSint.uf 		= UF_SIGLA2COD(bf.char2)
	reg->docSint.modelo 	= bf.int2
	reg->docSint.serie 		= bf.nchar(3)
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
	reg->docSint.situacao 	= situacaoSintegra2SituacaoEfd( bf.char1 )

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
	reg->docSint.uf 		= UF_SIGLA2COD(bf.char2)
	reg->docSint.modelo 	= bf.int2
	reg->docSint.serie 		= bf.nchar(3)
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
	reg->docSint.uf 		= UF_SIGLA2COD(bf.char2)
	reg->docSint.serie 		= bf.nchar(3)
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

''''''''
private function lerRegSintegraMercadoria(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.nchar(8+8)
	reg->itemId.id			  	= bf.nchar(14)
	reg->itemId.ncm			  	= vallng(bf.nchar(8))
	reg->itemId.descricao	  	= bf.nchar(53)
	reg->itemId.unidInventario 	= bf.nchar(6)
	reg->itemId.aliqIPI		  	= bf.dbl5_2
	reg->itemId.aliqICMSInt	  	= bf.dbl4_2
	reg->itemId.redBcICMS	  	= bf.dbl5_2
	reg->itemId.bcICMSST	  	= bf.dbl13_2

	'pular \r\n
	bf.char1
	bf.char1

	function = true
end function

''''''''
private function lerRegSintegraDocumentoItem(bf as bfile, reg as TRegistro ptr) as Boolean
	
	reg->docItemSint.cnpj 		= bf.nchar(14)
	bf.nchar(2)
	reg->docItemSint.serie 		= bf.nchar(3)
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("¨") then
		bf.char1
		reg->docItemSint.numero = bf.int9
	else
		reg->docItemSint.numero = bf.int6
	end if
	reg->docItemSint.cfop 		= bf.int4
	reg->docItemSint.CST 		= bf.nchar(3)
	reg->docItemSint.nroItem	= valint(bf.nchar(3))	
	reg->docItemSint.codMercadoria = bf.nchar(14)
	reg->docItemSint.qtd		= bf.dbl11_3
	reg->docItemSint.valor		= bf.dbl12_2
	reg->docItemSint.desconto	= bf.dbl12_2
	reg->docItemSint.bcICMS		= bf.dbl12_2
	reg->docItemSint.bcICMSST	= bf.dbl12_2
	reg->docItemSint.valorIPI	= bf.dbl12_2
	reg->docItemSint.aliqICMS	= bf.dbl4_2
	
	'pular \r\n
	bf.char1
	bf.char1

	function = true
end function

#define GENSINTEGRAKEY(r) ((r)->cnpj + (r)->serie + str((r)->numero) + str((r)->cfop))
  
''''''''
function Efd.lerRegistroSintegra(bf as bfile, reg as TRegistro ptr) as Boolean

	var tipo = bf.int2

	select case as const tipo
	case SINTEGRA_DOCUMENTO
		reg->tipo = SINTEGRA_DOCUMENTO
		if not lerRegSintegraDocumento(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		reg->docSint.chaveDict = GENSINTEGRAKEY(@reg->docSint)
		var antReg = cast(TRegistro ptr, sintegraDict->lookup(reg->docSint.chaveDict))
		if antReg = null then
			sintegraDict->add(reg->docSint.chaveDict, reg)
		else
			'' para cada alíquota diferente há um novo registro 50, mas nós só queremos os valores totais
			''antReg->docSint.valorTotal	+= reg->docSint.valorTotal
			''antReg->docSint.bcICMS		+= reg->docSint.bcICMS
			''antReg->docSint.ICMS		+= reg->docSint.ICMS
			''antReg->docSint.valorIsento += reg->docSint.valorIsento
			''antReg->docSint.valorOutras += reg->docSint.valorOutras

			reg->tipo = DESCONHECIDO 
		end if

	case SINTEGRA_DOCUMENTO_ST
		reg->tipo = SINTEGRA_DOCUMENTO
		if not lerRegSintegraDocumentoST(bf, reg) then
			return false
		end if

		reg->docSint.chaveDict = GENSINTEGRAKEY(@reg->docSint)
		var antReg = cast(TRegistro ptr, sintegraDict->lookup(reg->docSint.chaveDict))
		'' NOTA: pode existir registro 53 sem o correspondente 50, para quando só há ICMS ST, sem destaque ICMS próprio
		if antReg = null then
			sintegraDict->add(reg->docSint.chaveDict, reg)
		else
			''antReg->docSint.bcICMSST		+= reg->docSint.bcICMSST
			''antReg->docSint.ICMSST			+= reg->docSint.ICMSST
			''antReg->docSint.despesasAcess	+= reg->docSint.despesasAcess
			reg->tipo = DESCONHECIDO
		end if
	  
	case SINTEGRA_DOCUMENTO_IPI
		reg->tipo = SINTEGRA_DOCUMENTO
		if not lerRegSintegraDocumentoIPI(bf, reg) then
			return false
		end if

		reg->docSint.chaveDict = GENSINTEGRAKEY(@reg->docSint)
		var antReg = cast(TRegistro ptr, sintegraDict->lookup(reg->docSint.chaveDict))
		if antReg = null then
			onError("ERRO: Sintegra 53 sem 50: " & reg->docSint.chaveDict)
		else
			antReg->docSint.valorIPI		= reg->docSint.valorIPI
			antReg->docSint.valorIsentoIPI	= reg->docSint.valorIsentoIPI
			antReg->docSint.valorOutrasIPI	= reg->docSint.valorOutrasIPI
		end if

		reg->tipo = DESCONHECIDO 
		
	case SINTEGRA_DOCUMENTO_ITEM
		reg->tipo = SINTEGRA_DOCUMENTO_ITEM
		if not lerRegSintegraDocumentoItem(bf, reg) then
			return false
		end if

		var chaveDict = GENSINTEGRAKEY(@reg->docItemSint)
		var doc = cast(TRegistro ptr, sintegraDict->lookup(chaveDict))
		if doc = null then
			onError("ERRO: Sintegra 54 sem 50: " & chaveDict)
		end if
		
		reg->docItemSint.doc = @(doc->docSint)
		
	case SINTEGRA_MERCADORIA
		reg->tipo = ITEM_ID
		if not lerRegSintegraMercadoria(bf, reg) then
			return false
		end if

		'adicionar ao dicionário
		if itemIdDict->lookup(reg->itemId.id) = null then
			itemIdDict->add(reg->itemId.id, @reg->itemId)
		end if
		
	case else
		pularLinha(bf)
		reg->tipo = DESCONHECIDO
	end select

	function = true

end function

''''''''
function Efd.carregarSintegra(bf as bfile) as Boolean
	
	var fsize = bf.tamanho
	
	dim as TRegistro ptr tail = null
	var nroLinha = 0

	try
		do while bf.temProximo()		 
			var reg = new TRegistro
			
			nroLinha += 1

			if lerRegistroSintegra( bf, reg ) then 
				if not onProgress(null, bf.posicao / fsize) then
					exit do
				end if
				
				if reg->tipo <> DESCONHECIDO then
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
			 
			else
				exit do
			end if
		loop
	catch
		onError(!"\r\nErro ao carregar o registro da linha (" & nroLinha & !") do arquivo\r\n")
	endtry
	   
	function = true

end function

#include once "EfdExt.bi"
#include once "EfdBoCsvLoader.bi"
#include once "libs/BFile.bi"
#include once "libs/trycatch.bi"

const BO_CSV_SEP = asc(!"\t")
const BO_CSV_DIG = asc(".")

''''''''
constructor EfdBoCsvLoader(ctx as EfdBoLoaderContext ptr, opcoes as OpcoesExtracao ptr)
	base(ctx, opcoes)
end constructor

''''''''
function EfdBoCsvLoader.carregarCsvNFeDestSAFI(bf as bfile, emModoOutrasUFs as boolean) as TDFe_NFe ptr
	
	var dfe = new TDFe_Nfe
	
	dfe->operacao			= ENTRADA
	
	if not emModoOutrasUFs then
		dfe->chave				= bf.charCsv
		dfe->dataEmi			= csvDate2YYYYMMDD(bf.charCsv)
		dfe->cnpjEmit			= bf.charCsv
		dfe->nomeEmit			= bf.charCsv
		dfe->ieEmit				= trim(bf.charCsv)
		dfe->cnpjDest			= bf.charCsv
		dfe->ufDest				= UF_SIGLA2COD(bf.charCsv)
		dfe->nomeDest			= bf.charCsv
		dfe->bcICMSTotal		= bf.dblCsv
		dfe->ICMSTotal			= bf.dblCsv
		dfe->bcICMSSTTotal		= bf.dblCsv
		dfe->ICMSSTTotal		= bf.dblCsv
		dfe->valorOperacao		= bf.dblCsv
		dfe->ufEmit				= UF_SIGLA2COD(bf.charCsv)
		dfe->numero				= bf.intCsv
		dfe->serie				= bf.intCsv
		dfe->modelo				= bf.intCsv
	else
		dfe->chave				= bf.charCsv
		dfe->cnpjDest			= bf.charCsv
		dfe->nomeDest			= bf.charCsv
		dfe->dataEmi			= csvDate2YYYYMMDD(bf.charCsv)
		dfe->ufDest				= 35
		dfe->cnpjEmit			= bf.charCsv
		dfe->nomeEmit			= bf.charCsv
		dfe->ufEmit				= UF_SIGLA2COD(bf.charCsv)
		dfe->bcICMSTotal		= bf.dblCsv
		dfe->ICMSTotal			= bf.dblCsv
		dfe->bcICMSSTTotal		= bf.dblCsv
		dfe->ICMSSTTotal		= bf.dblCsv
		dfe->valorOperacao		= bf.dblCsv
		dfe->modelo				= bf.intCsv
		dfe->serie				= bf.intCsv
		dfe->numero				= bf.intCsv
	end if

	'' pular \r\n
	bf.char1
	bf.char1
	
	function = dfe
	
end function

''''''''
function EfdBoCsvLoader.carregarCsvNFeEmitSAFI(bf as bfile) as TDFe_NFe ptr
	
	var chave = bf.charCsv
	var dfe = cast(TDFe_NFe ptr, ctx->chaveDFeDict->lookup(chave))	
	if dfe = null then
		dfe = new TDFe_NFe
	end if
	
	dfe->chave				= chave
	dfe->dataEmi			= csvDate2YYYYMMDD(bf.charCsv)
	dfe->cnpjEmit			= bf.charCsv
	dfe->nomeEmit			= bf.charCsv
	dfe->ieEmit				= trim(bf.charCsv)
	dfe->ufEmit				= 35
	dfe->cnpjDest			= bf.charCsv
	dfe->ufDest				= UF_SIGLA2COD(bf.charCsv)
	dfe->nomeDest			= bf.charCsv
	dfe->bcICMSTotal		= bf.dblCsv
	dfe->ICMSTotal			= bf.dblCsv
	dfe->bcICMSSTTotal		= bf.dblCsv
	dfe->ICMSSTTotal		= bf.dblCsv
	dfe->valorOperacao		= bf.dblCsv
	var op = bf.charCsv
	dfe->operacao			= iif(op[0] = asc("S"), SAIDA, ENTRADA)
	dfe->numero				= bf.intCsv
	dfe->serie				= bf.intCsv
	dfe->modelo				= bf.intCsv
	
	'' devolução? inverter emit <-> dest
	if dfe->operacao = ENTRADA then
		swap dfe->cnpjEmit, dfe->cnpjDest
		swap dfe->ufEmit, dfe->ufDest
	end if
	
	'' pular \r\n
	bf.char1
	bf.char1
	
	function = dfe
	
end function

''''''''
function EfdBoCsvLoader.carregarCsvNFeEmitItensSAFI(bf as bfile, chave as string) as TDFe_NFeItem ptr
	
	var item = new TDFe_NFeItem
	
	bf.charCsv				'' pular versão
	bf.charCsv				'' pular cnpj emitente
	bf.charCsv				'' pular ie emitente
	bf.charCsv				'' pular cnpj dest
	item->modelo 			= bf.intCsv
	item->serie				= bf.intCsv
	item->numero			= bf.intCsv
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
function EfdBoCsvLoader.carregarCsvCTeSAFI(bf as bfile, emModoOutrasUFs as boolean) as TDFe_CTe ptr
	var dfe = new TDFe_CTe
	
	'' NOTA: só será possível saber se é operação de entrada ou saída quando pegarmos 
	''       o CNPJ base do contribuinte, que só vem no final do arquivo.......
	dfe->operacao		= DESCONHECIDA			
	
	bf.charCsv			'' pular chave quebrada
	dfe->serie			= bf.intCsv
	dfe->numero			= bf.intCsv
	dfe->cnpjEmit		= bf.charCsv
	dfe->dataEmi		= csvDate2YYYYMMDD(bf.charCsv)
	dfe->nomeEmit		= bf.charCsv
	dfe->ufEmit			= UF_SIGLA2COD(bf.charCsv)
	dfe->cnpjToma		= bf.charCsv
	dfe->nomeToma		= bf.charCsv
	dfe->ufToma			= bf.charCsv
	dfe->cnpjRem		= bf.charCsv
	dfe->nomeRem		= bf.charCsv
	dfe->ufRem			= bf.charCsv
	dfe->cnpjDest		= bf.charCsv
	dfe->nomeDest		= bf.charCsv
	dfe->ufDest			= UF_SIGLA2COD(bf.charCsv)
	dfe->cnpjExp		= bf.charCsv
	dfe->ufExp			= bf.charCsv
	dfe->cnpjReceb		= bf.charCsv
	dfe->ufReceb		= bf.charCsv
	dfe->tipo			= valint(left(bf.charCsv,1))
	dfe->chave			= bf.charCsv
	dfe->valorOperacao	= bf.dblCsv
	dfe->valorReceber	= bf.dblCsv
	dfe->qtdCCe			= bf.dblCsv
	dfe->cfop			= bf.intCsv
	dfe->nomeMunicIni	= bf.charCsv
	dfe->ufIni			= bf.charCsv
	dfe->nomeMunicFim	= bf.charCsv
	dfe->ufFim			= bf.charCsv
	dfe->modelo			= 57
	
	'' pular \r\n
	bf.char1
	bf.char1
	
	'' back patching
	if ctx->cteListHead = null then
		ctx->cteListHead = dfe
	else
		ctx->cteListTail->next_ = dfe
	end if
	
	ctx->cteListTail = dfe
	dfe->next_ = null
	dfe->parent = dfe
	
	function = dfe
	
end function

''''''''
function EfdBoCsvLoader.carregarCsvNFeEmitItens(bf as bfile, chave as string, extra as TDFe ptr) as TDFe_NFeItem ptr
	
	var item = new TDFe_NFeItem
	
	'' chave_nfe	num_doc_fiscal	cod_serie_doc_fiscal	cod_modelo	ind_tipo_documento_fiscal	ind_situacao_doc_fiscal	data_emissao	
	'' nome_rsocial_emit	num_cnpj_emit	num_ie_emit	cod_drt_emit	cod_est_emit	nome_rsocial_dest	num_cnpj_dest	num_cpf_dest	
	'' num_ie_dest	cod_drt_dest	cod_est_dest	num_item	descr_prod	cod_prod_servico	cod_gtin	cod_ncm	cod_cfop	
	'' cod_tributacao_icms	cod_csosn	perc_aliquota_icms	perc_aliquota_base_calc	perc_aliquota_icms_st	perc_reduc_icms_st	
	'' quant_comercial	unid_comercial	valor_produto_servico	valor_base_calc_icms	valor_icms	valor_base_calc_icms_st	valor_icms_st	
	'' valor_bc_icms_st_retido	valor_icms_st_retido	valor_ipi	valor_desconto	valor_frete	ind_modalidade_frete	valor_seguro	
	'' valor_outras_desp	valor_pis	valor_cofins	num_docto_importacao	num_fci	data_desembaraco	cod_est_desembaraco	
	'' descr_inf_adic_produto	ind_origem_mercadoria	cod_cnae

	chave 					= bf.varchar(BO_CSV_SEP)

	item->numero			= bf.varint(BO_CSV_SEP) ''vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->serie				= bf.varint(BO_CSV_SEP)
	item->modelo 			= bf.varint(BO_CSV_SEP)
	bf.varchar(BO_CSV_SEP) '' tipo
	bf.varchar(BO_CSV_SEP)	'' situação
	extra->dataEmi			= yyyyMmDd2YyyyMmDd(bf.varchar(BO_CSV_SEP))
	bf.varchar(BO_CSV_SEP) '' razão social emi
	bf.varchar(BO_CSV_SEP) '' cnpj emi
	bf.varchar(BO_CSV_SEP) '' ie emi
	bf.varchar(BO_CSV_SEP) '' drt emi
	bf.varchar(BO_CSV_SEP)	'' uf emi
	extra->nomeDest 		= bf.varchar(BO_CSV_SEP)
	extra->cnpjDest			= bf.varchar(BO_CSV_SEP)
	bf.varchar(BO_CSV_SEP) '' cpf dest
	bf.varchar(BO_CSV_SEP) '' ie dest
	bf.varchar(BO_CSV_SEP) '' drt dest
	extra->ufDest			= UF_SIGLA2COD(bf.varchar(BO_CSV_SEP))
	item->nroItem			= bf.varint(BO_CSV_SEP)
	item->descricao			= bf.varchar(BO_CSV_SEP)
	item->codProduto		= bf.varchar(BO_CSV_SEP)
	bf.varchar(BO_CSV_SEP)	'' GTIN
	item->ncm				= bf.varint(BO_CSV_SEP)
	item->cfop				= bf.varint(BO_CSV_SEP)
	item->cst				= bf.varint(BO_CSV_SEP)
	bf.varchar(BO_CSV_SEP) '' CSOSN
	item->aliqICMS			= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	bf.varchar(BO_CSV_SEP) '' redução bc
	item->aliqIcmsST		= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	bf.varchar(BO_CSV_SEP) '' redução bc ST
	item->qtd				= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->unidade			= bf.varchar(BO_CSV_SEP)
	item->valorProduto		= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->bcICMS			= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->ICMS				= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->bcICMSST			= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->IcmsST			= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	bf.varchar(BO_CSV_SEP) '' bc ICMS ST anterior
	bf.varchar(BO_CSV_SEP) '' ICMS ST anterior
	item->IPI				= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->desconto			= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	bf.varchar(BO_CSV_SEP) '' frete
	bf.varchar(BO_CSV_SEP) '' indicador frete
	bf.varchar(BO_CSV_SEP) '' seguro
	item->despesasAcess		= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	bf.varchar(BO_CSV_SEP) '' pis
	bf.varchar(BO_CSV_SEP) '' cofins
	bf.varchar(BO_CSV_SEP) '' num doc importacao
	bf.varchar(BO_CSV_SEP) '' num fci
	bf.varchar(BO_CSV_SEP) '' data desembaraco
	bf.varchar(BO_CSV_SEP) '' uf desembaraco
	bf.varchar(BO_CSV_SEP) '' info adicional
	bf.varchar(BO_CSV_SEP) '' origem mercadoria
	bf.varchar(BO_CSV_SEP) '' cnae
	item->next_ = null
	
	'' pular \r\n
	bf.char1
	bf.char1
	
	function = item
end function

''''''''
function EfdBoCsvLoader.carregar(nomeArquivo as String) as Boolean

	dim bf as bfile
   
	if not bf.abrir( nomeArquivo ) then
		return false
	end if
	
	dim as integer tipoArquivo
	dim as boolean isSafi = true
	if instr( nomeArquivo, "SAFI_NFe_Destinatario" ) > 0 then
		tipoArquivo = BO_NFe_Dest
		ctx->nfeDestSafiFornecido = true
	
	elseif instr( nomeArquivo, "SAFI_NFe_Emitente_Itens" ) > 0 then
		tipoArquivo = BO_NFe_Emit_Itens
		ctx->itemNFeSafiFornecido = true
	
	elseif instr( nomeArquivo, "SAFI_NFe_Emitente" ) > 0 then
		tipoArquivo = BO_NFe_Emit
		ctx->nfeEmitSafiFornecido = true
	
	elseif instr( nomeArquivo, "SAFI_CTe_CNPJ" ) > 0 then
		tipoArquivo = BO_CTe
		ctx->cteListHead = null
		ctx->cteListTail = null
		ctx->cteSafiFornecido = true
		
	elseif instr( nomeArquivo, "NFE_Emitente_Itens_SP_OSF" ) > 0 then
		tipoArquivo = BO_NFe_Emit_Itens
		isSafi = false
		ctx->itemNFeSafiFornecido = true
	
	else
		onError("impossível resolver tipo de arquivo pelo nome")
		return false
	end if

	var nroLinha = 1
		
	try
		var fsize = bf.tamanho

		'' pular header
		pularLinha(bf)
		nroLinha += 1
		
		var emModoOutrasUFs = false
		var extra = new TDFe
		
		do while bf.temProximo()		 
			if not onProgress(null, bf.posicao / fsize) then
				exit do
			end if
			
			if isSafi then
				'' outro header?
				if bf.peek1 <> asc("""") then
					'' final de arquivo?
					
					var linha = lcase(lerLinha(bf))
					if left(linha, 22) = "cnpj base contribuinte" or left(linha, 26) = "cnpj/cpf base contribuinte" then
						onProgress(null, 1)
						nroLinha += 1
						
						'' se for CT-e, temos que ler o CNPJ base do contribuinte para fazer um 
						'' patch em todos os tipos de operação (saída ou entrada)
						if tipoArquivo = BO_CTe then
							var cnpjBase = bf.charCsv
							var cte = ctx->cteListHead
							do while cte <> null 
								if left(cte->parent->cnpjEmit,8) = cnpjBase then
									cte->parent->operacao = SAIDA
								elseif left(cte->cnpjToma,8) = cnpjBase then
									cte->parent->operacao = ENTRADA
								end if
								adicionarDFe(cte->parent)
								cte = cte->next_
							loop
						end if
						exit do
					else
						emModoOutrasUFs = true
					end if
				end if
			end if
		
			select case as const tipoArquivo  
			case BO_NFe_Dest
				var dfe = carregarCsvNFeDestSAFI( bf, emModoOutrasUFs )
				if dfe <> null then
					adicionarDFe(dfe)
				end if
			
			case BO_NFe_Emit
				var dfe = carregarCsvNFeEmitSAFI( bf )
				if dfe <> null then
					adicionarDFe(dfe)
				end if
				
			case BO_NFe_Emit_Itens
				var chave = ""
				var nfeItem = iif(isSafi, _
					carregarCsvNFeEmitItensSAFI( bf, chave ), _
					carregarCsvNFeEmitItens( bf, chave, extra ))
				if nfeItem <> null then
					adicionarItemDFe(chave, nfeItem)

					var dfe = cast(TDFe_NFe ptr, ctx->chaveDFeDict->lookup(chave))
					'' nf-e não encontrada? pode acontecer se processarmos o csv de itens antes do csv de nf-e
					if dfe = null then
						'' só adicionar ao dicionário e à lista de DFe
						dfe = new TDFe_NFe
						dfe->chave = chave
						dfe->modelo = NFE
						if not isSafi then
							dfe->operacao = SAIDA
							dfe->dataEmi = extra->dataEmi
							dfe->numero = nfeItem->numero
							dfe->serie = nfeItem->serie
							dfe->cnpjDest = extra->cnpjDest
							dfe->nomeDest = extra->nomeDest
							dfe->ufDest = extra->ufDest
						end if
						adicionarDFe(dfe, false)
					end if
					
					if dfe->itemListHead = null then
						dfe->itemListHead = nfeItem
					else
						dfe->itemListTail->next_ = nfeItem
					end if
					
					dfe->itemListTail = nfeItem
				end if
			
			case BO_CTe
				var dfe = carregarCsvCTeSAFI( bf, emModoOutrasUFs )
			end select
			
			nroLinha += 1
		loop
		
		delete extra
		
		if not isSafi then
			'' se for informado só o itens NF-e, gravar a tabela NF-e com os dados disponíveis
			if opcoes->manterDb andalso ctx->itemNFeSafiFornecido andalso not ctx->nfeEmitSafiFornecido then
				var dfe = ctx->dfeListHead
				do while dfe <> null
					if dfe->modelo = NFe then
						adicionarDFe(cast(TDFe_NFe ptr, dfe))
					end if
					dfe = dfe->prox
				loop
			end if
			onProgress(null, 1)
		end if
		
		function = true
	
	catch
		onError(!"\r\n\tErro ao carregar linha " & nroLinha & !"\r\n")
		function = false
	endtry
	   
	bf.fechar()
	
end function

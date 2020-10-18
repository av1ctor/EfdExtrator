#include once "efd.bi"
#include once "bfile.bi"
#include once "trycatch.bi"

const BO_CSV_SEP = asc(!"\t")
const BO_CSV_DIG = asc(".")

''''''''
function Efd.carregarCsvNFeEmitItens(bf as bfile, chave as string, extra as TDFe ptr) as TDFe_NFeItem ptr
	
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
function Efd.carregarCsv(nomeArquivo as String) as Boolean

	dim bf as bfile
   
	if not bf.abrir( nomeArquivo ) then
		return false
	end if
	
	dim as integer tipoArquivo
	dim as boolean isSafi = true
	if instr( nomeArquivo, "SAFI_NFe_Destinatario" ) > 0 then
		tipoArquivo = BO_NFe_Dest
		nfeDestSafiFornecido = true
	
	elseif instr( nomeArquivo, "SAFI_NFe_Emitente_Itens" ) > 0 then
		tipoArquivo = BO_NFe_Emit_Itens
		itemNFeSafiFornecido = true
	
	elseif instr( nomeArquivo, "SAFI_NFe_Emitente" ) > 0 then
		tipoArquivo = BO_NFe_Emit
		nfeEmitSafiFornecido = true
	
	elseif instr( nomeArquivo, "SAFI_CTe_CNPJ" ) > 0 then
		tipoArquivo = BO_CTe
		cteListHead = null
		cteListTail = null
		cteSafiFornecido = true
		
	elseif instr( nomeArquivo, "NFE_Emitente_Itens_SP_OSF" ) > 0 then
		tipoArquivo = BO_NFe_Emit_Itens
		isSafi = false
		itemNFeSafiFornecido = true
	
	else
		onError("Erro: impossível resolver tipo de arquivo pelo nome")
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
							var cte = cteListHead
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

					var dfe = cast(TDFe ptr, chaveDFeDict->lookup(chave))
					'' nf-e não encontrada? pode acontecer se processarmos o csv de itens antes do csv de nf-e
					if dfe = null then
						'' só adicionar ao dicionário e à lista de DFe
						dfe = new TDFe
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
					
					if dfe->nfe.itemListHead = null then
						dfe->nfe.itemListHead = nfeItem
					else
						dfe->nfe.itemListTail->next_ = nfeItem
					end if
					
					dfe->nfe.itemListTail = nfeItem
				end if
			
			case BO_CTe
				var dfe = carregarCsvCTeSAFI( bf, emModoOutrasUFs )
			end select
			
			nroLinha += 1
		loop
		
		delete extra
		
		if not isSafi then
			'' se for informado só o itens NF-e, gravar a tabela NF-e com os dados disponíveis
			if opcoes.manterDb andalso itemNFeSafiFornecido andalso not nfeEmitSafiFornecido then
				var dfe = dfeListHead
				do while dfe <> null
					adicionarDFe(dfe)
					dfe = dfe->next_
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

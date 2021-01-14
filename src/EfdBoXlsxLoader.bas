#include once "EfdExt.bi"
#include once "EfdBoXlsxLoader.bi"
#include once "libs/trycatch.bi"

private function dbl2Cnpj(valor as double) as string
	return iif(valor <> 0, right("00000000000000" + str(valor), 14), "")
end function 

private function limparCNPJ(valor as string) as string
	return iif(len(valor) > 0, right("00000000000000" + strreplace(strreplace(strreplace(valor, ".", ""), "/", ""), "-", ""), 14), "")
end function

#define limparIE(valor) strreplace(valor, ".", "")

''''''''
constructor EfdBoXlsxLoader(ctx as EfdBoLoaderContext ptr, opcoes as OpcoesExtracao ptr)
	base(ctx, opcoes)
end constructor

''''''''
function EfdBoXlsxLoader.carregarXlsxNFeDest(rd as ExcelReader ptr) as TDFe_NFe ptr
	
	'' Chave Acesso NFe,	Número,	Série,	Modelo,	Data Emissão,	Razão Social Emitente
	'' CNPJ Emitente,	Número CPF Emitente,	Inscrição Estadual Emitente,	CRT,	DRT Emitente
	'' UF Emit,	Razão Social Destinatário,	CNPJ Destinatário,	Inscrição Estadual Destinatário,	DRT Destinatário
	'' UF Dest,	Tipo Doc Fiscal,	Descrição Natureza Operação,	Peso Liquido,	Peso Bruto	
	'' Informações Interesse Fisco,	Informações Complementares Interesse Contribuinte,	Indicador Modalidade Frete,	Situação Documento,	Dt. Cancelamento	
	'' Mercadoria - Valor,	Razão Social Transportador	CNPJ do Transportador,	Inscrição Estadual Transportador,	Placa Veículo Transportador	
	'' UF Veículo Transportador,	Total BC  ICMS,	Total ICMS,	Total BC ICMS-ST,	Total ICMS-ST	
	'' Total NFe,	Valor Total Frete,	Valor Total Seguro,	Quantidade Cartas de Correção Eletrônicas,	Quantidade Manifestações Destinatário

	var chave				= rd->read
	if len(chave) <> 44 then
		return null
	end if
		
	if ctx->chaveDFeDict->lookup(chave) <> null then
		return null
	end if

	var dfe = new TDFe_NFe
	
	dfe->loader				= LOADER_NFE_DEST
	dfe->operacao			= ENTRADA
	dfe->chave				= chave
	dfe->numero				= rd->readDbl
	dfe->serie				= rd->readInt
	dfe->modelo				= rd->readInt
	dfe->dataEmi			= rd->readDate
	dfe->nomeEmit			= rd->read(true)
	dfe->cnpjEmit			= dbl2Cnpj(rd->readDbl)
	rd->skip '' cpf emit
	dfe->ieEmit				= trim(limparIE(rd->read))
	rd->skip '' crt emit
	rd->skip '' drt emit
	dfe->ufEmit				= UF_SIGLA2COD(rd->read)
	dfe->nomeDest			= rd->read(true)
	dfe->cnpjDest			= dbl2Cnpj(rd->readDbl)
	dfe->ieDest				= trim(limparIE(rd->read))
	rd->skip '' drt dest
	dfe->ufDest				= UF_SIGLA2COD(rd->read)
	rd->skip '' tipo doc
	rd->skip '' descrição op
	rd->skip '' peso liq
	rd->skip '' peso bruto
	rd->skip '' info fisco
	rd->skip '' info contrib
	rd->skip '' frete
	rd->skip '' situação doc
	rd->skip '' data canc
	rd->skip '' merc valor
	rd->skip '' transportador
	rd->skip '' cnpj transportador
	rd->skip '' ie transportador
	rd->skip '' placa transportador
	rd->skip '' uf transportador
	dfe->bcICMSTotal		= rd->readDbl
	dfe->ICMSTotal			= rd->readDbl
	dfe->bcICMSSTTotal		= rd->readDbl
	dfe->ICMSSTTotal		= rd->readDbl
	dfe->valorOperacao		= rd->readDbl
	
	function = dfe

end function

''''''''
function EfdBoXlsxLoader.carregarXlsxNFeDestItens(rd as ExcelReader ptr) as TDFe_NFe ptr
	
	'' Chave de Acesso NFe, Número Documento Fiscal, Série Documento Fiscal, Modelo Documento Fiscal, Tipo Documento Fiscal, Situação Documento Fiscal, 
	'' Data Emissão, Razão Social Emitente, CNPJ Emitente, CPF Emitente, Inscrição Estadual Emitente, DRT Emitente,	UF Emitente, Razão Social Destinatário,
	'' CNPJ Destinatário, CPF Destinatário,	Inscrição Estadual Destinatário, DRT Destinatário, UF Destinatário, Item, Descrição Produto, Código Produto, 
	'' GTIN, NCM, CFOP, CST, O/CSOSN, Alíquota ICMS, Percentual Redução Base de Cálculo ICMS, Alíquota ICMS-ST, Percentual Redução Base de Cálculo ICMS-ST, 
	'' Quantidade Comercial, Unidade Comercial, Valor Produto ou Serviço, Valor Base de Cálculo ICMS, Valor ICMS, Valor Base Cálculo ICMS-ST, Valor ICMS-ST
	'' Valor Base Cálculo ICMS-ST Retido Operação Anterior, Valor ICMS-ST Retido Operação Anterior, Valor IPI, Valor Desconto,
	'' Valor Frete, Indicador Modalidade Frete, Valor Seguro, Valor Outras Despesas Acessórias, Valor PIS, Valor COFINS, 
	'' Percentual Alíquota Crédito Simples Nacional, Valor Crédito Simples Nacional,
	'' Número DI, Número FCI, Data Desembaraço, Código UF Desembaraço, Descrição Informações Adicionais Produto
	
	var chave				= rd->read
	if len(chave) <> 44 then
		return null
	end if
	
	var dfe = cast(TDFe_NFe ptr, ctx->chaveDFeDict->lookup(chave))	
	if dfe = null then
		dfe = new TDFe_NFe
	else
		if dfe->loader <> LOADER_NFE_DEST_ITENS then
			return null
		end if
	end if
	
	dfe->loader				= LOADER_NFE_DEST_ITENS
	dfe->operacao			= ENTRADA
	dfe->chave				= chave
	dfe->numero				= rd->readDbl
	dfe->serie				= rd->readInt
	dfe->modelo				= rd->readInt
	rd->skip '' tipo
	rd->skip '' situação
	dfe->dataEmi			= rd->readDate
	dfe->nomeEmit			= rd->read(true)
	dfe->cnpjEmit			= dbl2Cnpj(rd->readDbl)
	rd->skip '' cpf emit
	dfe->ieEmit				= trim(limparIE(rd->read))
	rd->skip '' drt emit
	dfe->ufEmit				= UF_SIGLA2COD(rd->read)
	dfe->nomeDest			= rd->read(true)
	dfe->cnpjDest			= dbl2Cnpj(rd->readDbl)
	if dfe->cnpjDest = "" then
		dfe->cnpjDest 		= rd->read
	else
		rd->skip '' cpf dest
	end if
	rd->skip '' ie dest
	rd->skip '' drt dest
	dfe->ufDest				= UF_SIGLA2COD(rd->read)
	rd->skip '' item
	rd->skip '' descrição prod
	rd->skip '' código prod
	rd->skip '' GTIN
	rd->skip '' NCM
	rd->skip '' CFOP
	rd->skip '' CST
	rd->skip '' CSOSN
	rd->skip '' aliq
	rd->skip '' red bc icms
	rd->skip '' aliq ST
	rd->skip '' red bc icms ST
	rd->skip '' qtd
	rd->skip '' unidade
	dfe->valorOperacao		+= rd->readDbl
	dfe->bcICMSTotal		+= rd->readDbl
	dfe->ICMSTotal			+= rd->readDbl
	dfe->bcICMSSTTotal		+= rd->readDbl
	dfe->ICMSSTTotal		+= rd->readDbl

	function = dfe

end function

''''''''
function EfdBoXlsxLoader.carregarXlsxNFeEmit(rd as ExcelReader ptr) as TDFe_NFe ptr
	
	var chave = rd->read
	if len(chave) <> 44 then
		return null
	end if
	
	var dfe = cast(TDFe_NFe ptr, ctx->chaveDFeDict->lookup(chave))	
	if dfe = null then
		dfe = new TDFe_NFe
	end if
	
	'' Chave Acesso NFe,	Número,	Série,	Modelo,	Data Emissão,	Razão Social Emitente
	'' CNPJ Emitente,	Inscrição Estadual Emitente,	CRT, DRT Emit,	UF Emit
	'' Razão Social Destinatário,	CNPJ ou CPF do Destinatário,	Inscrição Estadual Destinatário,	CNAE Destinatário,	Cod Cnae Destinatário (Cadesp)	
	'' DRT Dest,	UF Dest,	Tipo Doc Fiscal,	Descrição Natureza Operação,	Peso Liquido(NFe SP Volume)
	'' Peso Bruto(NFe SP Volume),	Informações Interesse Fisco,	EMPTY, Informações Complementares Interesse Contribuinte,	Indicador Modalidade Frete,	Situação Documento
	'' Dt. Cancelamento,	Mercadoria - Valor,	Razão Social Transportador,	CNPJ do Transportador,	Inscrição Estadual Transportador
	'' Placa Veículo Transportador,	UF Veículo Transportador,	Total BC  ICMS,	Total ICMSv,	Total BC ICMS-ST
	'' Total ICMS-ST,	Total NFe,	Valor Total Frete,	Valor Total Seguro,	Valor ICMS Inter. UF Destino	
	'' Valor ICMS Inter. UF Remetente,	Quantidade Cartas de Correção Eletrônicas,	Quantidade Manifestações Detinatário

	dfe->loader				= LOADER_NFE_EMIT
	dfe->chave				= chave
	dfe->numero				= rd->readDbl
	dfe->serie				= rd->readInt
	dfe->modelo				= rd->readInt
	dfe->dataEmi			= rd->readDate
	dfe->nomeEmit			= rd->read(true)
	dfe->cnpjEmit			= dbl2Cnpj(rd->readDbl)
	dfe->ieEmit				= trim(limparIE(rd->read))
	rd->skip '' crt emit
	rd->skip '' drt emit
	dfe->ufEmit				= UF_SIGLA2COD(rd->read)
	dfe->nomeDest			= rd->read(true)
	dfe->cnpjDest			= limparCNPJ(rd->read)
	dfe->ieDest				= trim(limparIE(rd->read))
	rd->skip '' cnae dest
	rd->skip '' cnae dest cadesp
	rd->skip '' drt dest
	dfe->ufDest				= UF_SIGLA2COD(rd->read)
	var op = rd->read
	dfe->operacao			= iif(op[0] = asc("S"), SAIDA, ENTRADA)
	rd->skip '' descrição op
	rd->skip '' peso liq
	rd->skip '' peso bruto
	rd->skip '' info fisco
	rd->skip '' EMPTY
	rd->skip '' info contrib
	rd->skip '' frete
	rd->skip '' situação doc
	rd->skip '' data canc
	rd->skip '' merc valor
	rd->skip '' transportador
	rd->skip '' cnpj transportador
	rd->skip '' ie transportador
	rd->skip '' placa transportador
	rd->skip '' uf transportador
	dfe->bcICMSTotal		= rd->readDbl
	dfe->ICMSTotal			= rd->readDbl
	dfe->bcICMSSTTotal		= rd->readDbl
	dfe->ICMSSTTotal		= rd->readDbl
	dfe->valorOperacao		= rd->readDbl
	
	'' devolução? inverter emit <-> dest
	if dfe->operacao = ENTRADA then
		swap dfe->cnpjEmit, dfe->cnpjDest
		swap dfe->ufEmit, dfe->ufDest
	end if

	function = dfe
	
end function

''''''''
function EfdBoXlsxLoader.carregarXlsxNFeEmitItens(rd as ExcelReader ptr, chave as string, extra as TDFe ptr) as TDFe_NFeItem ptr
	
	'' Chave de Acesso NFe,	Número Documento Fiscal,	 Série Documento Fiscal,	Modelo Documento Fiscal, Tipo Documento Fiscal,	
	'' Situação Documento Fiscal,	Data Emissão,	Razão Social Emitente,	CNPJ Emitente,	Inscrição Estadual Emitente,	
	'' DRT Emitente,	UF Emitente,	Razão Social Destinatário,	CNPJ Destinatário,	CPF Destinatário,	
	'' Inscrição Estadual Destinatário,	DRT Destinatário,	UF Destinatário,	Item,	Descrição Produto,	
	'' Código Produto,	GTIN,	NCM,	CEST, CFOP,	CST,	
	'' O/CSOSN,	Alíquota ICMS,	Percentual Redução Base de Cálculo ICMS,	Alíquota ICMS-ST,	Percentual Redução Base de Cálculo ICMS-ST,	
	'' Quantidade Comercial,	Unidade Comercial,	 Valor Produto ou Serviço ,	 Valor Base de Cálculo ICMS,	 Valor ICMS, 	
	'' Valor Base Cálculo ICMS-ST,	Valor ICMS-ST,	Valor Base Cálculo ICMS-ST Retido Operação Anterior,	Valor ICMS-ST Retido Operação Anterior,	Valor IPI,	
	'' Valor Desconto,	Valor Frete,	Indicador Modalidade Frete,	Valor Seguro,	Valor Outras Despesas Acessórias, 
	'' Valor PIS,	Valor COFINS,	Número DI,	Número FCI,	Data Desembaraço
	'' Código UF Desembaraço,	Descrição Informações Adicionais Produto
		
	chave = rd->read
	if len(chave) <> 44 then
		return null
	end if
	
	var item = new TDFe_NFeItem
	
	item->numero			= rd->readDbl
	item->serie				= rd->readInt
	item->modelo 			= rd->readInt
	rd->skip '' tipo
	rd->skip	'' situação
	extra->dataEmi 			= rd->readDate
	rd->skip '' razão social emi
	rd->skip '' cnpj emi
	rd->skip '' ie emi
	rd->skip '' drt emi
	rd->skip	'' uf emi
	rd->skip	'' razão social dest
	extra->cnpjDest			= dbl2Cnpj(rd->readDbl)
	rd->skip '' cpf dest
	rd->skip '' ie dest
	rd->skip '' drt dest
	extra->ufDest 			= UF_SIGLA2COD(rd->read)
	item->nroItem			= rd->readInt
	item->descricao			= rd->read(true)
	item->codProduto		= rd->read
	rd->skip	'' GTIN
	item->ncm				= rd->readInt
	rd->skip	'' CEST
	item->cfop				= rd->readInt
	item->cst				= rd->readInt
	rd->skip '' CSOSN
	item->aliqICMS			= rd->readDbl
	rd->skip '' redução bc
	item->aliqIcmsST		= rd->readDbl
	rd->skip '' redução bc ST
	item->qtd				= rd->readDbl
	item->unidade			= rd->read
	item->valorProduto		= rd->readDbl
	item->bcICMS			= rd->readDbl
	item->ICMS				= rd->readDbl
	item->bcICMSST			= rd->readDbl
	item->icmsST			= rd->readDbl
	rd->skip '' bc ICMS ST anterior
	rd->skip '' ICMS ST anterior
	item->IPI				= rd->readDbl
	item->desconto			= rd->readDbl
	rd->skip '' frete
	rd->skip '' indicador frete
	rd->skip '' seguro
	item->despesasAcess		= rd->readDbl
	item->next_ = null
		
	function = item
	
end function

''''''''
function EfdBoXlsxLoader.carregarXlsxCTe(rd as ExcelReader ptr, op as TipoOperacao) as TDFe_CTe ptr
	
	'' ---em branco---,	Chave Acesso CT-e (char),	Série,	Num CTe,	Data Emissão	, Ind. Situação CT-e
	'' CNPJ Emitente,	Num. Inscr. Est. Emitente,	Razão Social Emitente,	UF Emitente,	CNPJ Tomador,	
	'' Num Inscr. Est. Tomador,	Razão Social Tomador,	Indicador Tomador Serviço,	UF Tomador,	CNPJ Remetente,	
	'' Razão Social Remetente,	UF Remetente,	CNPJ Destinatário,	Razão Social Destinatário,	UF Destinatário,	
	'' CNPJ Expedidor,	UF Expedidor,	CNPJ Recebedor,	UF Recebedor,	Tipo CT-e,	
	'' indSN,	Código CFOP,	Descr. Nat. Operação,	Descr. Modal,	Descr. Servico,	
	'' Descr. Cst,	Município Inicial,	UF Inicial,	Município Final,	UF Final,	
	'' Aliqüota Icms,	Perc. Redução Bc,	Valor Bc St Retido,	Valor Icms St Retido,	Valor Icms OutrasUF,	
	'' Valor Crédito Outorgado/Presumido,	Valor Total Prest. Serviço,	Valor Icms,	Valor Bc ICMS,	Quantidade de CCE,	
	'' Quantidade de manifestações do tomador
	
	rd->skip '' ---em branco---
	var chave 				= rd->read
	if len(chave) <> 44 then
		return null
	end if
	
	var dfe = new TDFe_CTe

	dfe->operacao		= op
	dfe->chave			= chave
	dfe->serie			= rd->readInt
	dfe->numero			= rd->readInt
	dfe->dataEmi		= rd->readDate
	rd->skip '' ind situação
	dfe->cnpjEmit		= dbl2Cnpj(rd->readDbl)
	rd->skip '' ie emit
	dfe->nomeEmit		= rd->read(true)
	dfe->ufEmit			= UF_SIGLA2COD(rd->read)
	dfe->cnpjToma		= dbl2Cnpj(rd->readDbl)
	rd->skip '' ie toma
	dfe->nomeToma		= rd->read(true)
	rd->skip '' ind toma
	dfe->ufToma			= rd->read
	dfe->cnpjRem		= dbl2Cnpj(rd->readDbl)
	dfe->nomeRem		= rd->read(true)
	dfe->ufRem			= rd->read
	dfe->cnpjDest		= dbl2Cnpj(rd->readDbl)
	dfe->nomeDest		= rd->read(true)
	dfe->ufDest			= UF_SIGLA2COD(rd->read)
	dfe->cnpjExp		= dbl2Cnpj(rd->readDbl)
	dfe->ufExp			= rd->read
	dfe->cnpjReceb		= dbl2Cnpj(rd->readDbl)
	dfe->ufReceb		= rd->read
	dfe->tipo			= valint(left(rd->read,1))
	rd->skip '' indSN
	dfe->cfop			= rd->readInt
	rd->skip '' Descr. Nat. Operação
	rd->skip '' Descr. Modal
	rd->skip '' Descr. Servico
	rd->skip '' Descr. Cst
	dfe->nomeMunicIni	= rd->read
	dfe->ufIni			= rd->read
	dfe->nomeMunicFim	= rd->read
	dfe->ufFim			= rd->read
	rd->skip '' Aliqüota Icms
	rd->skip '' Perc. Redução Bc
	rd->skip '' Valor Bc St Retido
	rd->skip '' Valor Icms St Retido
	rd->skip '' Valor Icms OutrasUF,	
	rd->skip '' Valor Crédito Outorgado/Presumido
	dfe->valorOperacao		= rd->readDbl
	rd->skip '' Valor Icms
	rd->skip '' Valor Bc ICMS
	dfe->valorReceber	= dfe->valorOperacao
	dfe->qtdCCe			= rd->readInt
	dfe->modelo			= 57
	
	function = dfe
	
end function

''''''''
function EfdBoXlsxLoader.carregarXlsxSATItens(rd as ExcelReader ptr, chave as string) as TDFe_NFeItem ptr
	
	'' ---em branco---, Num Inscr. Estadual Emitente,	Data Emissão,	Identificação CF-e,	Número Cupom CF-e,	Indicador Cupom Cancelado	
	'' Número Série,	Valor ICMS,	Número Item,	Código Produto,	Código EAN,	
	'' Descrição Produto,	Código NCM,	Código CFOP 04 Posições,	Unidade Comercial,	Quantidade Comercial,	
	'' Indicador Regra Cálculo,	Valor Unitário Comercialização,	Valor Produtos,	Valor Desconto,	Valor Outro,	
	'' Valor Item,	Valor Rateio Desconto,	Valor Rateio Acrescimo,	Indicador Origem,	Código CST/CSOSN,	
	'' Alíquota ICMS,	Código CST PIS,	Valor Base Cálculo PIS,	Alíquota PIS,	Valor PIS,	
	'' Quantidade Vendida PIS,	Valor Alíquota PIS,	Valor Base Cálculo PIS-ST,	Alíquota PIS-ST,	Quantidade Vendida PIS-ST,	
	'' Valor Alíquota PIS-ST,	Valor PIS-ST,	Código CST COFINS,	Valor Base Cálculo COFINS,	Alíquota COFINS	Valor COFINS,	
	'' Quantidade Vendida COFINS,	Valor Alíquota COFINS,	Valor Base Cálculo COFINS-ST,	Alíquota COFINS-ST,	Quantidade Vendida COFINS-ST,	
	'' Valor Alíquota COFINS-ST,	Valor COFINS-ST,	Informações Adicicionais,	Descrição Campo,	Descrição Texto Campo
		
	rd->skip '' ---em branco---
	var ie = trim(rd->read)
	if len(ie) = 0 then
		return null
	end if
	if ie[0] < asc("0") or ie[0] > asc("9") then
		return null
	end if
	
	rd->skip '' data emi
	chave = rd->read
	if len(chave) <> 3+44 then
		return null
	end if
	
	chave = right(chave, 44)
	
	var item = new TDFe_NFeItem

	item->modelo 			= SAT
	item->numero			= rd->readInt
	rd->skip '' situação
	item->serie				= rd->readInt
	item->ICMS				= rd->readDbl
	item->nroItem			= rd->readInt
	item->codProduto		= rd->read
	rd->skip '' EAN
	item->descricao			= rd->read(true)
	item->ncm				= rd->readInt
	item->cfop				= rd->readInt
	item->unidade			= rd->read
	item->qtd				= rd->readDbl
	rd->skip '' Indicador Regra Cálculo
	rd->skip '' Valor Unitário Comercialização
	rd->skip '' Valor Produtos
	item->desconto			= rd->readDbl
	item->despesasAcess		= rd->readDbl
	item->valorProduto		= rd->readDbl
	rd->skip '' Valor Rateio Desconto
	rd->skip '' Valor Rateio Acrescimo
	rd->skip '' Indicador Origem
	item->cst				= rd->readInt
	item->aliqICMS			= rd->readDbl
	rd->skip '' Código CST PIS
	rd->skip '' Valor Base Cálculo PIS
	rd->skip '' Alíquota PIS
	item->IPI				= rd->readDbl
	item->bcICMS			= item->valorProduto
	item->bcICMSST			= 0
	item->next_ = null
	
	function = item
	
end function

''''''''
function EfdBoXlsxLoader.carregarXlsxSAT(rd as ExcelReader ptr) as TDFe_NFe ptr
	
	'' ---em branco---, Num Inscr. Estadual Emitente,	Número de Série do SAT,	Data Emissão,	Hora Emissão,	
	'' Indicador Cupom Cancelado,	Identificação CF-e,	Data Recepção Cupom,	Número Cupom CF-e,	Indicador Possui Destinatário,	
	'' Valor Total CF-e,	Valor Total ICMS,	Valor Total Produtos,	Valor Total Desconto,	Valor Total Pis,	Valor Total Cofins,	
	'' Valor Total Pis-ST,	Valor Total Cofins-ST,	Valor Total Outros,	Valor Acrescimo/Desconto Subtotal,	Valor Cfe Lei 12741
	
	rd->skip '' ---em branco---
	var ie = rd->read
	if len(ie) = 0 then
		return null
	end if
	if ie[0] < asc("0") or ie[0] > asc("9") then
		return null
	end if
	
	rd->skip '' Número de Série do SAT
	var dEmi 				= rd->readDate
	rd->skip '' Hora Emissão
	rd->skip '' Indicador Cupom Cancelado
	
	var chave = rd->read
	if len(chave) <> 3+44 then
		return null
	end if
	
	chave = right(chave, 44)
	
	var dfe = cast(TDFe_NFe ptr, ctx->chaveDFeDict->lookup(chave))	
	if dfe = null then
		dfe = new TDFe_NFe
	end if
	
	dfe->chave				= chave
	dfe->dataEmi			= dEmi
	dfe->ieEmit				= str(cdbl(ie))
	rd->skip '' Data Recepção Cupom
	dfe->numero				= rd->readInt
	dfe->serie				= 0
	dfe->modelo				= SAT
	rd->skip '' Indicador Possui Destinatário
	dfe->valorOperacao		= rd->readDbl
	dfe->ICMSTotal			= rd->readDbl
	dfe->bcICMSTotal		= dfe->valorOperacao
	dfe->ufEmit				= 35
	dfe->cnpjDest			= "00000000000000"
	dfe->ufDest				= 35
	dfe->operacao			= SAIDA
	dfe->bcICMSSTTotal		= 0
	dfe->ICMSSTTotal		= 0
	
	function = dfe
	
end function

''''''''
function EfdBoXlsxLoader.carregar(nomeArquivo as String) as Boolean

	if left(nomeArquivo, 1) = "~" then
		return true
	elseif left(nomeArquivo, 7) = "SpedEFD" then
		return true
	elseif nomeArquivo = "__efd__.xlsx" then
		return true
	elseif instr(nomeArquivo, "NFe_Destinatario_Itens_OSF") > 0 then
		onProgress(null, 1)
		return true
	end if

	dim as integer tipoArquivo
	dim as string nomePlanilhas(0 to 1)
	var linhaInicial = 2
	
	if instr( nomeArquivo, "NFe_Destinatario_OSF" ) > 0 then
		tipoArquivo = BO_NFe_Dest
		ctx->nfeDestSafiFornecido = true
		nomePlanilhas(0) = "Planilha NF-e por Destinatário"

	elseif instr( nomeArquivo, "NFe_Emitente_Itens_OSF" ) > 0 then
		tipoArquivo = BO_NFe_Emit_Itens
		ctx->itemNFeSafiFornecido = true
		nomePlanilhas(0) = "Planilha"
	
	elseif instr( nomeArquivo, "NFe_Emitente_OSF" ) > 0 then
		tipoArquivo = BO_NFe_Emit
		ctx->nfeEmitSafiFornecido = true
		nomePlanilhas(0) = "Planilha NF-e por Emitente"
		linhaInicial = 3
	
	elseif instr( nomeArquivo, "CTe_CNPJ_Emitente_Tomador_Remetente_Destinatario_OSF" ) > 0 then
		tipoArquivo = BO_CTe
		nomePlanilhas(0) = "CT-e por Emitente"
		nomePlanilhas(1) = "CT-e por Tomador"
		ctx->cteSafiFornecido = true
		linhaInicial = 3
	
	elseif instr( nomeArquivo, "SAT_-_CuponsEmitidosPorContribuinteCNPJ_OSF" ) > 0 then
		tipoArquivo = BO_SAT
		ctx->nfeEmitSafiFornecido = true
		nomePlanilhas(0) = "Cupons emitidos em dado periodo"
	
	elseif instr( nomeArquivo, "SAT_-_ItensDeCuponsCNPJ_OSF" ) > 0 then
		tipoArquivo = BO_SAT_Itens
		ctx->itemNFeSafiFornecido = true
		nomePlanilhas(0) = "Itens de Cupons"
	
	elseif instr( nomeArquivo, "NFC-e_itens_OSF" ) > 0 then
		tipoArquivo = BO_NFCe_Itens
		ctx->itemNFeSafiFornecido = true
		nomePlanilhas(0) = "Itens"
		onError(!"\n\tErro: relatório não suportado ainda")
		return false
		
	elseif instr( nomeArquivo, "REDF_consulta_Cupons_Fiscais_ECF" ) > 0 then
		tipoArquivo = SAFI_ECF
		ctx->nfeEmitSafiFornecido = true
		nomePlanilhas(0) = "REDF - Cupons Fiscais"
		onError(!"\n\tErro: relatório não suportado ainda")
		return false
	
	elseif instr( nomeArquivo, "REDF_-_Consulta_Cupons_Fiscais_ECF_e_itens_do_CF" ) > 0 then
		tipoArquivo = BO_ECF_Itens
		ctx->itemNFeSafiFornecido = true
		nomePlanilhas(0) = "REDF - Itens dos Cupons Fiscais"
		onError(!"\n\tErro: relatório não suportado ainda")
		return false
	
	else
		onError(!"\n\tErro: impossível resolver tipo de arquivo pelo nome")
		return false
	end if
	
	var reader = new ExcelReader()
	
	if not reader->open(nomeArquivo) then
		onError(!"\n\tErro: arquivo não encontrado ou inválido")
		delete reader
		return false
	end if
	
	var plan = 0
	var extra = new TDFe
	do
		var nomePlanilha = nomePlanilhas(plan)
		if nomePlanilha = "" then
			exit do
		end if
		
		if not reader->setSheet(nomePlanilha) then
			onError(!"\n\tErro: planilha não encontrada (" + nomePlanilha + ")")
			delete reader
			return false
		end if
		
		var nroLinha = 1

		try
			do while (reader->nextRow()) 
				if nroLinha >= linhaInicial then
					select case as const tipoArquivo  
					case BO_NFe_Dest
						var dfe = carregarXlsxNFeDest(reader)
						if dfe <> null then
							adicionarDFe(dfe)
						end if
					
					case BO_NFe_Emit
						var dfe = carregarXlsxNFeEmit( reader )
						if dfe <> null then
							adicionarDFe(dfe)
						end if
						
					case BO_NFe_Emit_Itens
						var chave = ""
						var nfeItem = carregarXlsxNFeEmitItens( reader, chave, extra )
						if nfeItem <> null then
							adicionarItemDFe(chave, nfeItem)

							var dfe = cast(TDFe_Nfe ptr, ctx->chaveDFeDict->lookup(chave))
							'' nf-e não encontrada? pode acontecer se processarmos o csv de itens antes do csv de nf-e
							if dfe = null then
								dfe = new TDFe_NFe
								'' só adicionar ao dicionário e à lista de DFe
								dfe->chave = chave
								dfe->modelo = NFE
								dfe->operacao = SAIDA
								dfe->dataEmi = extra->dataEmi
								dfe->numero = nfeItem->numero
								dfe->serie = nfeItem->serie
								dfe->cnpjDest = extra->cnpjDest
								dfe->nomeDest = extra->nomeDest
								dfe->ufDest = extra->ufDest
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
						var dfe = carregarXlsxCTe( reader, iif(plan = 0, SAIDA, ENTRADA) )
						if dfe <> null then
							adicionarDFe(dfe)
						end if
						
					case BO_SAT
						var dfe = carregarXlsxSAT( reader )
						if dfe <> null then
							adicionarDFe(dfe)
						end if
						
					case BO_SAT_Itens
						var chave = ""
						var satItem = carregarXlsxSATItens( reader, chave )
						if satItem <> null then
							adicionarItemDFe(chave, satItem)

							var dfe = cast(TDFe_NFe ptr, ctx->chaveDFeDict->lookup(chave))
							'' sat não encontrado? pode acontecer se processarmos o csv de itens antes do csv de nf-e
							if dfe = null then
								dfe = new TDFe_NFe
								'' só adicionar ao dicionário e à lista de DFe
								dfe->chave = chave
								dfe->modelo = SAT
								adicionarDFe(dfe, false)
							end if
							
							if dfe->itemListHead = null then
								dfe->itemListHead = satItem
							else
								dfe->itemListTail->next_ = satItem
							end if
							
							dfe->itemListTail = satItem
						end if
						
					case BO_NFCe_Itens
						''var dfe = carregarXlsxNFCeItens( reader )
						''if dfe <> null then
						''end if

					case BO_ECF_Itens
						''var dfe = carregarXlsxECFItens( reader )
						''if dfe <> null then
						''end if
						
					end select
				end if
				
				nroLinha += 1
			loop
			
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
			
			function = true
		
		catch
			onError(!"\r\n\tErro ao carregar linha " & nroLinha & !"\r\n")
			function = false
		endtry
	
		plan += 1
	loop while plan <= ubound(nomePlanilhas)
	
	delete extra
	
	onProgress(null, 1)
	
	delete reader
	
end function

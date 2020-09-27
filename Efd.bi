
#include once "Dict.bi"
#include once "bfile.bi"
#include once "ExcelReader.bi"
#include once "ExcelWriter.bi"
#include once "DB.bi"
#include once "Lua/lualib.bi"
#include once "PDFer.bi"

enum TTipoArquivo
	TIPO_ARQUIVO_EFD
	TIPO_ARQUIVO_SINTEGRA
end enum

type OpcoesExtracao
	gerarRelatorios 				as boolean = false
	pularLreAoGerarRelatorios 		as boolean = false
	pularLrsAoGerarRelatorios 		as boolean = false
	pularLRaicmsAoGerarRelatorios	as boolean = false
	acrescentarDados				as boolean = false
	formatoDeSaida 					as FileType = FT_XLSX
	somenteRessarcimentoST 			as boolean = false
	dbEmDisco 						as boolean = false
	manterDb						as boolean = false
	filtrarCnpj						as boolean = false
	filtrarChaves					as boolean = false
	listaCnpj(any)					as string
	listaChaves(any)				as string
	highlight						as boolean
end type

enum TipoRegistro
	MESTRE
	PARTICIPANTE
	ITEM_ID
	BEM_CIAP
	INFO_COMPL
	DOC_NF										'' NF, NF-e, NFC-e
	DOC_NF_INFO									'' informações complementares de interesse do fisco
	DOC_NF_ITEM    								'' item de NF-e (só informado para entradas)
	DOC_NF_ITEM_RESSARC_ST						'' ressarcimento ST
	DOC_NF_ANAL									'' analítico
	DOC_NF_DIFAL								'' diferencial de alíquota
	DOC_CT     									'' CT, CT-e, CT-e OS, BP-e
	DOC_CT_DIFAL				
	DOC_CT_ANAL					
	EQUIP_ECF					
	ECF_REDUCAO_Z				
	DOC_ECF						
	DOC_ECF_ITEM				
	DOC_ECF_ANAL				
	DOC_NFSCT									'' NF de comunicação e telecomunicação
	DOC_NFSCT_ANAL
	DOC_SAT
	DOC_SAT_ANAL
	DOC_NF_ELETRIC
	DOC_NF_ELETRIC_ANAL
	APURACAO_ICMS_PERIODO		
	APURACAO_ICMS_PROPRIO		
	APURACAO_ICMS_AJUSTE		
	APURACAO_ICMS_PROPRIO_OBRIG	
	APURACAO_ICMS_ST_PERIODO	
	APURACAO_ICMS_ST			
	INVENTARIO_TOTAIS
	INVENTARIO_ITEM
	CIAP_TOTAL
	CIAP_ITEM
	CIAP_ITEM_DOC
	FIM_DO_ARQUIVO								'' NOTA: anterior à assinatura digital que fica no final no arquivo
	DESCONHECIDO   				
	LUA_CUSTOM									'' tratado no script Lua
	SINTEGRA_DOCUMENTO 							= 50
	SINTEGRA_DOCUMENTO_IPI 						= 51
	SINTEGRA_DOCUMENTO_ST						= 53
	SINTEGRA_DOCUMENTO_ITEM						= 54
	SINTEGRA_MERCADORIA							= 75
	__TipoRegistro__LEN__
end enum

enum TipoAtividade
	ATIV_INDUSTRIAL_OU_EQUIPARADO = 0
	ATIV_OUTROS					  = 1
end enum

type TMestre
	versaoLayout		as integer
	original			as boolean
	dataIni				as zstring * 8+1
	dataFim				as zstring * 8+1
	nome				as zstring * 100+1
	cnpj           		as zstring * 14+1
	cpf            		as longint
	uf					as zstring * 2+1
	ie			      	as zstring * 14+1
	municip		   		as integer
	im					as zstring * 32+1
	suframa				as zstring * 9+1
	perfil				as byte
	atividade			as TipoAtividade
end type

type TParticipante
	id			      	as zstring * 60+1
	nome           		as zstring * 100+1
	pais		 	   	as integer
	cnpj           		as zstring * 14+1
	cpf            		as zstring * 11+1
	ie			      	as zstring * 14+1
	municip		   		as integer
	suframa		   		as zstring * 9+1
	ender		      	as zstring * 60+1
	num			   		as zstring * 10+1
	compl		      	as zstring * 60+1
	bairro		   		as zstring * 60+1
end type

enum TipoOperacao
	ENTRADA		  = 0	'' NF
	SAIDA		  = 1	'' NF
	AQUISICAO	  = 0	'' CT
	PRESTACAO	  = 1	'' CT
	DESCONHECIDA  = 2
end enum

enum TipoEmitente
	PROPRIO		  = 0
	TERCEIRO	  = 1
end enum

enum TipoModelo
	NF             = 01
	NF_AVULSA      = &h1b
	NFC            = 02
	CUPOM          = &h2d
	CUPOM_PASSAGEM = &h2e
	NF_PRODUTOR    = 04
	NFC_ELET       = 06
	NF_TRANSP      = 07
	CT_ROD         = 08
	CT_AVULSO      = &h8b
	CT_AQUA        = 09
	CT_AEREO       = 10
	CT_FERROV      = 11
	BILHETE_ROD    = 13
	BILHETE_AQUA   = 14
	BILHETE_BAGAG  = 15
	BILHETE_FERROV = 16
	RESUMO_DIARIO  = 18
	NFS_COMUNIC    = 21
	NFS_TELE       = 22
	CT_MULTIMODAL  = 26
	NF_FERROV_CARG = 27
	NFC_GAS        = 28
	NFC_AGUA       = 29
	NFE			   = 55
	CTE			   = 57
	SAT            = 59
	ECF            = 60
	BPE            = 63
	NFCE           = 65
	CTEOS          = 67
end enum

enum TipoSituacao
	REGULAR		      = 0
	EXTEMPORANEO      = 1
	CANCELADO		  = 2
	CANCELADO_EXT     = 3      'extemporâneo
	DENEGADO		  = 4
	INUTILIZADO	      = 5
	COMPLEMENTAR      = 6
	COMPLEMENTAR_EXT  = 7      'extemporâneo
	REGIME_ESPECIAL   = 8
	SUBSTITUIDO       = 9
	__TipoSituacao__LEN__
end enum

enum TipoPagamento
	A_VISTA			= 0
	A_PRAZO			= 1
	OUTROS			= 2
end enum

enum TipoFrete
	CONTA_EMIT		= 0
	CONTA_DEST		= 1
	CONTA_TERCEIRO	= 2
	SEM_FRETE		= 9
end enum

enum TipoItemId
	TI_Mercadoria_para_Revenda 	= 0
	TI_Materia_Prima 			= 1
	TI_Embalagem 				= 2
	TI_Produto_em_Processo 		= 3
	TI_Produto_Acabado 			= 4
	TI_Subproduto 				= 5
	TI_Produto_Intermediario 	= 6
	TI_Material_de_Uso_e_Consumo = 7
	TI_Ativo_Imobilizado 		= 8
	TI_Servicos 				= 9
	TI_Outros_insumos 			= 10
	TI_Outras 					= 99
end enum

type TRegistro_ as TRegistro ptr

type TItemId
	id             as zstring * 60+1
	descricao      as zstring * 256+1
	codBarra       as zstring * 32+1
	codAnterior    as zstring * 60+1
	unidInventario as zstring * 6+1
	tipoItem       as TipoItemId
	ncm            as LongInt
	exIPI          as zstring * 3+1
	codGenero      as integer
	codServico     as zstring * 5+1
	aliqICMSInt    as Double
	CEST           as integer
	aliqIPI		   as double				'' só presente no SINTEGRA
	redBcIcms	   as double				'' //
	bcICMSST	   as double				'' //
end type

type TInfoCompl
	id             	as zstring * 6+1
	descricao      	as zstring * 256+1
end type

type TBemCiap
	id             	as zstring * 60+1
	tipoMerc		as integer
	descricao      	as zstring * 256+1
	principal	   	as zstring * 60+1
	codAnal			as zstring * 60+1
	parcelas		as integer
end type

enum TipoResponsavelRetencaoRessarcST
	REMETENTE_DIRETO = 1
	REMETENTE_INDIRETO = 2
	PROPRIO_DECLARANTE = 3
end enum

enum TipoMotivoRessarcST
	RES_VENDA_OUTRA_UF = 1
	RES_SAIDA_COM_ISENCAO = 2
	RES_PERDA_OU_DETERIORACAo = 3
	RES_FURTO_OU_ROUBO = 4
	RES_EXPORTACAO = 5
	RES_OUTROS = 9
end enum

enum TipoDocArrecadacao
	ARRECADA_GIA = 1
	ARRECADA_GNRE = 2
end enum

type TDocNFItem_ as TDocNFItem ptr

type TDocNFItemRessarcSt
	documentoPai   			as TDocNFItem_
	modeloUlt				as TipoModelo
	numeroUlt				as longint
	serieUlt				as zstring * 4+1
	dataUlt					as zstring * 8+1		'AAAAMMDD
	idParticipanteUlt		as zstring * 60+1
	qtdUlt					as double
	valorUlt				as double
	valorBcST				as double
	chaveNFeUlt				as zstring * 44+1
	numItemNFeUlt			as short
	bcIcmsUlt				as double
	aliqIcmsUlt				as double
	limiteBcIcmsUlt			as double
	icmsUlt					as double
	aliqIcmsStUlt			as double
	res						as double
	responsavelRet			as TipoResponsavelRetencaoRessarcST
	motivo					as TipoMotivoRessarcST
	chaveNFeRet				as zstring * 44+1
	idParticipanteRet		as zstring * 60+1
	serieRet				as zstring * 4+1
	numeroRet				as longint
	numItemNFeRet			as short
	tipDocArrecadacao		as TipoDocArrecadacao
	numDocArrecadacao		as zstring * 32+1
	next_					as TDocNFItemRessarcSt ptr
end type

type TDocNF_ as TDocNF ptr

type TDocNFItem                       ' nota: só é obrigatório para entradas!!!
	documentoPai   			as TDocNF_
	numItem        			as Integer
	itemId         			as zstring * 60+1
	descricao      			as zstring * 256+1
	qtd            			as double
	unidade        			as zstring * 6+1
	valor          			as Double
	desconto       			as double
	indMovFisica   			as byte
	cstICMS        			as integer
	cfop           			as Integer
	codNatureza    			as zstring * 10+1
	bcICMS         			as Double
	aliqICMS       			as double
	ICMS           			as Double
	bcICMSST       			as Double
	aliqICMSST     			as Double
	ICMSST         			as Double
	indApuracao    			as Byte
	cstIPI         			as Integer
	codEnqIPI      			as zstring * 2+1
	bcIPI          			as double
	aliqIPI        			as Double
	IPI            			as Double
	cstPIS         			as integer
	bcPIS          			as Double
	aliqPISPerc    			as Double
	qtdBcPIS       			as double
	aliqPISMoed    			as Double
	PIS            			as Double
	cstCOFINS      			as Integer
	bcCOFINS       			as Double
	aliqCOFINSPerc 			as Double
	qtdBcCOFINS    			as double
	aliqCOFINSMoed 			as Double
	COFINS         			as Double
	itemRessarcStListHead 	as TDocNFItemRessarcSt ptr
	itemRessarcStListTail 	as TDocNFItemRessarcSt ptr
end type

type TDocECF_ as TDocECF ptr

type TDocECFItem
	documentoPai   as TDocECF_
	numItem        as Integer
	itemId         as zstring * 60+1
	qtd            as double
	qtdCancelada   as double
	unidade        as zstring * 6+1
	valor          as Double
	cstICMS        as integer
	cfop           as Integer
	aliqICMS       as double
	PIS            as Double
	COFINS         as Double
end type

type TDocDifAliq
	fcp				as double
	icmsDest		as double
	icmsOrigem		as double
end type

type TDocItemAnal
	documentoPai   			as TRegistro_
	cst						as integer
	cfop					as integer
	aliq					as double
	valorOp					as double
	bc						as double
	ICMS					as double
	bcST					as double
	ICMSST					as double
	redBC					as double
	IPI						as double
	next_					as TDocItemAnal ptr
end type

type TDocInfoCompl
	idCompl					as zstring * 6+1
	extra					as zstring * 255+1
	next_					as TDocInfoCompl ptr
end type

type TDocDF
	operacao				as TipoOperacao
	situacao				as TipoSituacao
	emitente				as TipoEmitente
	idParticipante			as zstring * 60+1
	modelo					as TipoModelo
	dataEmi					as zstring * 8+1		'AAAAMMDD
	dataEntSaida			as zstring * 8+1
	serie					as zstring * 4+1
	subserie				as zstring * 8+1
	numero					as longint
	chave					as zstring * 44+1
	valorTotal				as double
	bcICMS					as double
	ICMS					as double
	difal					as TDocDifAliq
	infoComplListHead		as TDocInfoCompl ptr
	infoComplListTail		as TDocInfoCompl ptr
	itemAnalListHead 		as TDocItemAnal ptr
	itemAnalListTail 		as TDocItemAnal ptr
end type

type TDocNF extends TDocDF
	pagamento		as TipoPagamento
	valorDesconto	as double
	valorAbatimento as double
	valorMerc		as double
	frete			as TipoFrete
	valorFrete		as double
	valorSeguro		as double
	valorAcessorias as double
	bcICMSST		as double
	ICMSST			as double
	IPI				as double
	PIS				as double
	COFINS			as double
	PISST			as double
	COFINSST		as double
	nroItens		as integer
end type

type TDocCT extends TDocDF
	tipoCTe				as integer
	chaveRef			as zstring * 44+1		'' para CT-e do tipo complementar, substituto ou anulador
	valorDesconto		as double
	frete				as TipoFrete
	valorServico		as double
	valorNaoTributado	as double
	codInfComplementar	as zstring * 6+1
	municipioOrigem		as integer
	municipioDestino	as integer
end type

type TEquipECF_ as TEquipECF ptr

type TDocECF extends TDocDF
	equipECF			as TEquipECF_
	PIS					as double
	COFINS				as double
	cpfCnpjAdquirente	as zstring * 14+1
	nomeAdquirente		as zstring * 60+1
	nroItens			as integer
end type

type TDocSAT extends TDocDF
	PIS					as double
	COFINS				as double
	cpfCnpjAdquirente	as zstring * 14+1
	serieEquip			as zstring * 09+1
	descontos			as double
	valorMerc 			as double
	despesasAcess		as double
	pisST				as double
	cofinsST			as double
	nroItens			as integer
end type

type TDocumentoSintegraBase
	cnpj           	as zstring * 14+1
	serie          	as zstring * 3+1
	numero         	as integer
	cfop           	as short
end type

type TDocumentoSintegra extends TDocumentoSintegraBase
	ie             	as zstring * 14+1
	dataEmi        	as zstring * 8+1
	uf             	as byte
	modelo		  	as TipoModelo
	operacao	   	as TipoOperacao
	valorTotal     	as Double
	bcICMS  		as Double
	ICMS  		  	as Double
	bcICMSST		as Double
	ICMSST  		as Double
	valorIsento	  	as double
	valorOutras	  	as double
	despesasAcess  	as double
	valorIPI		as double
	valorIsentoIPI	as double
	valorOutrasIPI	as double
	aliqICMS	  	as double					'' NOTA: não usar se houver mais de um registro 50 para a mesma NF-e, pois as alíquotas são diferentes
	situacao	    as TipoSituacao
	chave		  	as zstring * 44+1
	chaveDict	  	as zstring * 50+1
end type

type TDocumentoItemSintegra extends TDocumentoSintegraBase
	doc				as TDocumentoSintegra ptr
	CST				as zstring * 3+1
	nroItem			as integer
	codMercadoria	as zstring * 14+1
	qtd				as double
	valor			as double
	desconto		as double
	bcICMS			as double
	bcIcmsST		as double
	valorIPI		as double
	aliqIcms		as double
end type

type TApuracaoIcmsPeriodo
	dataIni					as zstring * 8+1
	dataFim					as zstring * 8+1
	totalDebitos			as double
	ajustesDebitos			as double
	totalAjusteDeb			as double
	estornosCredito			as double
	totalCreditos			as double
	ajustesCreditos			as double
	totalAjusteCred			as double
	estornoDebitos			as double
	saldoCredAnterior		as double
	saldoDevedorApurado		as double
	totalDeducoes			as double
	icmsRecolher			as double
	saldoCredTransportar	as double
	debExtraApuracao		as double
end type

type TApuracaoIcmsSTPeriodo
	dataIni					as zstring * 8+1
	dataFim					as zstring * 8+1
	UF						as zstring * 2+1
	mov						as boolean
	saldoCredAnterior		as double
	devolMercadorias		as double
	totalRessarciment		as double
	totalOutrosCred			as double
	ajusteCred				as double
	totalRetencao			as double
	totalOutrosDeb			as double
	ajusteDeb				as double
	saldoAntesDed			as double
	totalDeducoes			as double
	icmsRecolher			as double
	saldoCredTransportar	as double
	debExtraApuracao		as double
end type

type TEquipECF
	modelo					as TipoModelo
	modeloEquip				as zstring * 20+1
	numSerie				as zstring * 21+1
	numCaixa				as integer
end type

type TECFReducaoZ
	equipECF				as TEquipECF ptr
	dataMov					as zstring * 8+1
	cro						as longint
	crz						as longint
	numOrdem				as longint
	valorFinal				as double
	valorBruto				as double
	numIni					as integer
	numFim					as integer
	itemAnalListHead 		as TDocItemAnal ptr
	itemAnalListTail 		as TDocItemAnal ptr
end type

type TInventarioTotais
	dataInventario			as zstring * 8+1
	valorTotalEstoque		as double
	motivoInventario		as integer
end type

type TInventarioItem
	dataInventario			as zstring * 8+1
	itemId         			as zstring * 60+1
	unidade					as zstring * 6+1
	qtd						as double
	valorUnitario			as double
	valorItem				as double
	indPropriedade			as integer
	idParticipante			as zstring * 60+1
	txtComplementar			as zstring * 99+1
	codConta				as zstring * 32+1
	valorItemIR				as double
end type

type TCiapTotal
	dataIni					as zstring * 8+1
	dataFim					as zstring * 8+1
	saldoInicialICMS		as double
	parcelasSoma			as double
	valorTributExpSoma		as double
	valorTotalSaidas		as double
	indicePercSaidas		as double
	valorIcmsAprop			as double
	valorOutrosCred			as double
end type

type TCiapItem
	pai						as TCiapTotal ptr
	bemId         			as zstring * 60+1
	dataMov					as zstring * 8+1
	tipoMov					as zstring * 2+1
	valorIcms				as double
	valorIcmsSt				as double
	valorIcmsFrete			as double
	valorIcmsDifal			as double
	parcela					as integer
	valorParcela			as double
	docCnt					as integer
end type

type TCiapItemDoc
	pai         			as TCiapItem ptr
	indEmi					as integer
	idParticipante			as zstring * 60+1
	modelo					as integer
	serie					as zstring * 3+1
	numero					as integer
	chaveNfe				as zstring * 44+1
	dataEmi					as zstring * 8+1
end type

type TLuaReg
	tipo					as zstring * 4+1
	table					as integer
end type

type TArquivoInfo
	nome					as zstring * 256+1
end type

type TRegistro
	tipo           			as TipoRegistro
	arquivo					as TArquivoInfo ptr
	linha					as integer
	union
		mestre      		as TMestre
		part        		as TParticipante
		nf         			as TDocNF
		itemNF     			as TDocNFItem
		ct         			as TDocCT
		ecf         		as TDocECF
		itemECF     		as TDocECFItem
		sat         		as TDocSAT
		docInfoCompl		as TDocInfoCompl
		docSint	  			as TDocumentoSintegra
		docItemSint	  		as TDocumentoItemSintegra
		itemId      		as TItemId
		bemCiap				as TBemCiap
		infoCompl			as TInfoCompl
		apuIcms	  			as TApuracaoIcmsPeriodo
		apuIcmsST  			as TApuracaoIcmsSTPeriodo
		itemAnal			as TDocItemAnal
		itemRessarcSt		as TDocNFItemRessarcSt
		equipECF			as TEquipECF
		ecfRedZ				as TECFReducaoZ
		invTotais			as TInventarioTotais
		invItem				as TInventarioItem
		ciapTotal			as TCiapTotal
		ciapItem			as TCiapItem
		ciapItemDoc			as TCiapItemDoc
		lua					as TLuaReg
	end union
	next_          			as TRegistro ptr
end type

enum SAFI_TipoArquivo
	SAFI_NFe_Dest
	SAFI_NFe_Emit
	SAFI_NFe_Emit_Itens
	SAFI_CTe
	SAFI_SAT
	SAFI_SAT_Itens
	SAFI_NFCe_Itens
	SAFI_ECF
	SAFI_ECF_Itens
end enum

enum SAFI_Dfe_Fornecido
	MASK_SAFI_NFE_DEST_FORNECIDO = &b00000001
	MASK_SAFI_NFE_EMIT_FORNECIDO = &b00000010
	MASK_SAFI_ITEM_NFE_FORNECIDO = &b00000100
	MASK_SAFI_CTE_FORNECIDO 	 = &b00001000
end enum

type TDFe_ as TDFe

type TDFe_NFeItem									'' Nota: só existe para NF-e emitidas, já que para as recebidas os itens constam na EFD
	serie			as integer
	numero			as integer
	modelo			as TipoModelo
	nroItem			as integer
	cfop			as short
	ncm				as longint
	cst				as integer
	codProduto		as zstring * 60+1
	descricao		as zstring * 256+1
	qtd				as double
	unidade			as zstring * 6+1
	valorProduto	as double
	desconto		as double
	despesasAcess	as double
	bcICMS			as double
	aliqICMS		as double
	ICMS			as double
	bcICMSST		as double
	IPI				as double
	next_			as TDFe_NFeItem ptr
end type

type TDFe_NFe
	ieEmit			as zstring * 14+1
	ieDest			as zstring * 14+1
	bcICMSTotal		as double
	ICMSTotal		as double
	bcICMSSTTotal	as double
	ICMSSTTotal		as double
	
	itemListHead	as TDFe_NFeItem ptr
	itemListTail	as TDFe_NFeItem ptr
end type

type TDFe_CTe
	cnpjToma		as zstring * 14+1
	nomeToma		as zstring * 100+1
	ufToma			as zstring * 2+1
	cnpjRem			as zstring * 14+1
	nomeRem			as zstring * 100+1
	ufRem			as zstring * 2+1
	cnpjExp			as zstring * 14+1
	ufExp			as zstring * 2+1
	cnpjReceb		as zstring * 14+1
	ufReceb			as zstring * 2+1
	tipo			as byte
	valorReceber	as double
	qtdCCe			as double
	cfop			as integer
	nomeMunicIni	as zstring * 64+1
	ufIni			as zstring * 2+1
	nomeMunicFim	as zstring * 64+1
	ufFim			as zstring * 2+1
	next_			as TDFe_CTe ptr					'' usado para dar patch 
	parent			as TDFe_ ptr
end type

enum TDFE_LOADER
	LOADER_UNKNOWN
	LOADER_NFE_DEST
	LOADER_NFE_DEST_ITENS
	LOADER_NFE_EMIT
	LOADER_NFE_EMIT_ITENS
	LOADER_CTE
	LOADER_NFCE
	LOADER_SAT
	LOADER_ECF
end enum

type TDFe
	modelo			as TipoModelo
	operacao		as TipoOperacao					'' entrada ou saída
	chave			as zstring * 44+1
	dataEmi			as zstring * 8+1
	serie			as integer
	numero			as integer
	cnpjEmit		as zstring * 14+1
	nomeEmit		as zstring * 100+1
	ufEmit			as byte
	cnpjDest		as zstring * 14+1
	nomeDest		as zstring * 100+1
	ufDest			as byte
	valorOperacao	as double
	loader			as TDFE_LOADER
	
	union
		nfe			as TDFe_NFe
		cte			as TDFe_CTe
	end union
	
	next_			as TDFe ptr
end type

type InfoAssinatura
	assinante		as string
	cpf				as string
	hashDoArquivo	as string
end type

enum TipoRelatorio
	REL_LRE				= 1
	REL_LRS				= 2
	REL_RAICMS			= 3
	REL_RAICMSST		= 4
end enum

type RelSomatorioLR
	chave			as zstring * 10+1
	situacao		as TipoSituacao
	cst				as integer
	cfop			as integer
	aliq			as double
	valorOp 		as double
	bc 				as double
	icms 			as double
	bcST 			as double
	icmsST 			as double
	ipi 			as double
end type

enum RelLinhaTipo
	REL_LIN_DF_ENTRADA
	REL_LIN_DF_SAIDA
	REL_LIN_DF_ITEM_ANAL
	REL_LIN_DF_REDZ
	REL_LIN_DF_SAT
end enum

type RelLinhaDF
	doc 			as TDocDF ptr
	part 			as TParticipante ptr
end type

type RelLinhaAnal
	sit 			as TipoSituacao
	item 			as TDocItemAnal ptr
end type

type RelLinhaRedZ
	doc 			as TECFReducaoZ ptr
end type

type RelLinhaSat
	doc 			as TDocSAT ptr
end type

type RelLinha
	tipo			as RelLinhaTipo
	highlight		as boolean
	union
		df			as RelLinhaDF
		anal		as RelLinhaAnal
		redZ		as RelLinhaRedZ
		sat			as RelLinhaSat
	end union
end type

type RelPagina
	page			as PdfTemplatePageNode ptr
	emitir			as boolean
end type

type ProgressoCB as sub(estagio as const wstring ptr, porCompleto as double)

enum TipoInconsistencia
	TI_ESCRIT_FALTA
	TI_ESCRIT_FANTASMA
	TI_ALIQ
	TI_DUP
	TI_DIF
	TI_RESSARC_ST
	TI_CRED
	TI_SEL
	TI_DEB
end enum

enum TipoRegime
	TR_RPA				= 2
	TR_ESTIMATIVA		= 3
	TR_SIMPLIFICADO		= 4
	TR_MICROEMPRESA		= 5
	TR_RPA_DECENDIAL	= asc("A")
	TR_SN				= asc("N")
	TR_SN_MEI			= asc("O") 
	TR_EPP				= asc("M")
	TR_EPP_A			= asc("G")
	TR_EPP_B			= asc("H")
	TR_RURAL_PF			= asc("P")
end enum

type CustomLuaCb
	reader			as zstring ptr
	writer			as zstring ptr
	rel_entradas	as zstring ptr
	rel_saidas		as zstring ptr
	rel_outros		as zstring ptr
end type

type Efd
public:
	declare constructor ()
	declare destructor ()
	declare sub iniciarExtracao(nomeArquivo as String, opcoes as OpcoesExtracao)
	declare sub finalizarExtracao(mostrarProgresso as ProgressoCB)
	declare function carregarTxt(nomeArquivo as String, mostrarProgresso as ProgressoCB) as Boolean
	declare function carregarCsv(nomeArquivo as String, mostrarProgresso as ProgressoCB) as Boolean
	declare function carregarXlsx(nomeArquivo as String, mostrarProgresso as ProgressoCB) as Boolean
	declare function processar(nomeArquivo as string, mostrarProgresso as ProgressoCB) as Boolean
	declare sub analisar(mostrarProgresso as ProgressoCB)
	declare function getPlanilha(nome as const zstring ptr) as ExcelWorksheet ptr
   
private:
	declare sub configurarDB()
	declare sub configurarScripting()
	
	declare function lerRegistro(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegistroSintegra(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerTipo(bf as bfile, tipo as zstring ptr) as TipoRegistro
	declare function lerRegDocNFItem(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNF ptr) as Boolean
	declare function lerRegDocNFElet(bf as bfile, reg as TRegistro ptr) as Boolean
	declare sub lerAssinatura(bf as bfile)
	declare function carregarSintegra(bf as bfile, mostrarProgresso as ProgressoCB) as Boolean
	declare function carregarCsvNFeDestSAFI(bf as bfile, emModoOutrasUFs as boolean) as TDFe ptr
	declare function carregarCsvNFeEmitSAFI(bf as bfile) as TDFe ptr
	declare function carregarCsvNFeEmitItensSAFI(bf as bfile, chave as string) as TDFe_NFeItem ptr
	declare function carregarCsvCTeSAFI(bf as bfile, emModoOutrasUFs as boolean) as TDFe ptr
	declare function carregarCsvNFeEmitItens(bf as bfile, chave as string) as TDFe_NFeItem ptr
	declare function carregarXlsxNFeDest(reader as ExcelReader ptr) as TDFe ptr
	declare function carregarXlsxNFeDestItens(reader as ExcelReader ptr) as TDFe ptr
	declare function carregarXlsxNFeEmit(rd as ExcelReader ptr) as TDFe ptr
	declare function carregarXlsxNFeEmitItens(rd as ExcelReader ptr, chave as string) as TDFe_NFeItem ptr
	declare function carregarXlsxCTe(rd as ExcelReader ptr, op as TipoOperacao) as TDFe ptr
	declare function carregarXlsxSAT(rd as ExcelReader ptr) as TDFe ptr
	declare function carregarXlsxSATItens(rd as ExcelReader ptr, chave as string) as TDFe_NFeItem ptr
	
	declare sub adicionarDFe(dfe as TDFe ptr)
	declare sub adicionarItemDFe(chave as const zstring ptr, item as TDFe_NFeItem ptr)
	declare sub adicionarEfdDfe(chave as zstring ptr, operacao as TipoOperacao, dataEmi as zstring ptr, valorOperacao as double)
	declare sub adicionarDocEscriturado(doc as TDocDF ptr)
	declare sub adicionarDocEscriturado(doc as TDocECF ptr)
	declare sub adicionarDocEscriturado(doc as TDocSAT ptr)
	declare sub adicionarItemNFEscriturado(item as TDocNFItem ptr)
	declare sub adicionarRessarcStEscriturado(doc as TDocNFItemRessarcSt ptr)
	declare sub adicionarItemEscriturado(item as TItemId ptr)
	declare function filtrarPorCnpj(idParticipante as const zstring ptr) as boolean
	declare function filtrarPorChave(chave as const zstring ptr) as boolean
	declare function getInfoCompl(info as TDocInfoCompl ptr) as string
	
	declare sub addRegistroAoDB(reg as TRegistro ptr)
	
	declare sub criarPlanilhas()
	declare sub gerarPlanilhas(nomeArquivo as string, mostrarProgresso as ProgressoCB)
	
	declare sub gerarRelatorios(nomeArquivo as string, mostrarProgresso as ProgressoCB)
	declare sub gerarRelatorioApuracaoICMS(nomeArquivo as string, reg as TRegistro ptr)
	declare sub gerarRelatorioApuracaoICMSST(nomeArquivo as string, reg as TRegistro ptr)
	declare sub iniciarRelatorio(relatorio as TipoRelatorio, nomeRelatorio as string, sufixo as string)
	declare sub adicionarDocRelatorioEntradas(doc as TDocDF ptr, part as TParticipante ptr, highlight as boolean)
	declare sub adicionarDocRelatorioSaidas(doc as TDocDF ptr, part as TParticipante ptr, highlight as boolean)
	declare sub adicionarDocRelatorioSaidas(doc as TECFReducaoZ ptr, highlight as boolean)
	declare sub adicionarDocRelatorioSaidas(doc as TDocSAT ptr, highlight as boolean)
	declare sub adicionarDocRelatorioItemAnal(sit as TipoSituacao, anal as TDocItemAnal ptr)
	declare sub finalizarRelatorio()
	declare sub relatorioSomarLR(sit as TipoSituacao, anal as TDocItemAnal ptr)
	declare function codMunicipio2Nome(cod as integer) as string
	declare sub gerarPaginaRelatorio(lastPage as boolean = false)
	declare sub gerarResumoRelatorio()
	declare sub gerarResumoRelatorioHeader()
	declare sub setNodeText(page as PdfTemplatePageNode ptr, id as string, value as string)
	declare sub setNodeText(page as PdfTemplatePageNode ptr, id as string, value as wstring ptr)
	declare sub setChildText(row as PdfTemplateNode ptr, id as string, value as string)
	declare sub setChildText(row as PdfTemplateNode ptr, id as string, value as wstring ptr)
	declare function gerarLinhaDFe(lg as boolean, highlight as boolean) as PdfTemplateNode ptr
	declare function gerarLinhaAnal() as PdfTemplateNode ptr
	declare function criarPaginaRelatorio(emitir as boolean) as RelPagina ptr
	
	declare sub analisarInconsistenciasLRE(mostrarProgresso as ProgressoCB)
	declare sub analisarInconsistenciasLRS(mostrarProgresso as ProgressoCB)
	
	declare sub exportAPI(L as lua_State ptr)
	declare static function luacb_efd_rel_addItemAnalitico cdecl(L as lua_State ptr) as long
	declare static function luacb_efd_participante_get cdecl(L as lua_State ptr) as long

	arquivos				as TList 		'' de TArquivoInfo
	tipoArquivo				as TTipoArquivo
	
	'' registros das EFD's e do Sintegra (reiniciados a cada novo .txt carregado)
	regListHead         	as TRegistro ptr = null
	nroRegs             	as integer = 0
	participanteDict    	as TDict
	itemIdDict          	as TDict
	bemCiapDict          	as TDict
	infoComplDict			as TDict
	sintegraDict			as TDict
	ultimoReg   			as TRegistro ptr
	ultimoDocNFItem  		as TDocNFItem ptr
	ultimoEquipECF			as TEquipECF ptr
	ultimoECFRedZ			as TRegistro ptr
	ultimoInventario		as TInventarioTotais ptr
	ultimoCiap				as TCiapTotal ptr
	ultimoCiapItem			as TCiapItem ptr
	nroLinha				as integer
	regMestre				as TRegistro ptr

	'' planilhas que serão geradas (mantidos do início ao fim da extração)
	ew                  	as ExcelWriter ptr
	entradas            	as ExcelWorksheet ptr
	saidas              	as ExcelWorksheet ptr
	apuracaoIcms			as ExcelWorksheet ptr
	apuracaoIcmsST			as ExcelWorksheet ptr
	inventario				as ExcelWorksheet ptr
	ciap					as ExcelWorksheet ptr
	ressarcST				as ExcelWorksheet ptr
	inconsistenciasLRE		as ExcelWorksheet ptr
	inconsistenciasLRS		as ExcelWorksheet ptr
	nomeArquivoSaida		as string
	opcoes					as OpcoesExtracao

	'' registros das NF-e's e CT-e's retirados dos relatórios do Infoview (mantidos do início ao fim da extração)
	chaveDFeDict			as TDict
	dfeListHead				as TDFe ptr = null
	dfeListTail				as TDFe ptr = null
	nroDfe					as integer = 0
	cteListHead				as TDFe_CTe ptr = null	'' usado para fazer patch no tipo de operação
	cteListTail				as TDFe_CTe ptr = null
	nfeDestSafiFornecido 	as boolean
	nfeEmitSafiFornecido 	as boolean
	itemNFeSafiFornecido 	as boolean
	cteSafiFornecido		as boolean
	
	'' base de dados de configuração
	dbConfig				as TDb ptr
	
	'' base de dados temporária usadada para análises e cruzamentos
	db						as TDb ptr
	db_dfeEntradaInsertStmt	as TDbStmt ptr
	db_dfeSaidaInsertStmt	as TDbStmt ptr
	db_itensDfeSaidaInsertStmt as TDbStmt ptr
	db_LREInsertStmt		as TDbStmt ptr
	db_itensNfLRInsertStmt	as TDbStmt ptr
	db_LRSInsertStmt		as TDbStmt ptr
	db_ressarcStItensNfLRSInsertStmt as TDbStmt ptr
	db_itensIdInsertStmt as TDbStmt ptr
	
	'' geração de relatórios em formato PDF com o layout do programa EFD-ICMS-IPI da RFB
	baseTemplatesDir		as string
	ultimoRelatorio			as TipoRelatorio
	ultimoRelatorioSufixo	as string
	relSomaLRDict			as TDict
	relSomaLRList			as TList			'' de RelSomatorioLR
	nroRegistrosRel			as integer
	municipDict				as TDict
	relLinhasList			as TList			'' de RelLinha
	relNroLinhas			as double
	relYPos					as double
	relNroPaginas			as integer
	relTemplate				as PdfTemplate ptr
	relPage					as PdfTemplatePageNode ptr
	relPaginasList			as TList			'' de RelPagina
	
	''
	assinaturaP7K_DER(any)	as byte
	infAssinatura			as InfoAssinatura ptr
	
	'' scripting
	lua						as lua_State ptr
	customLuaCbDict			as TDict			'' de CustomLuaCb
end type


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#define DdMmYyyy2Yyyy_Mm(s) (mid(s,1,4) + "-" + mid(s,5,2))

#define STR2CNPJ(s) (left(s,2) + "." + mid(s,3,3) + "." + mid(s,3+3,3) + "/" + mid(s,3+3+3,4) + "-" + right(s,2))

#define STR2CPF(s) (left(s,3) + "." + mid(s,4,3) + "." + mid(s,4+3,3) + "-" + right(s,2))

#define DBL2MONEYBR(d) (format(d,"#,#,0.00"))

#define MUNICIPIO2SIGLA(m) (iif(m >= 1100000 and m <= 5399999, ufCod2Sigla(m \ 100000), "EX"))

declare function csvDate2YYYYMMDD(s as zstring ptr) as string 
declare function ddMmYyyy2YyyyMmDd(s as const zstring ptr) as string
declare function yyyyMmDd2Datetime(s as const zstring ptr) as string 
declare function YyyyMmDd2DatetimeBR(s as const zstring ptr) as string 
declare function STR2IE(ie as string) as string
declare function tipoItem2Str(tipo as TipoItemId) as string
declare function dupstr(s as const zstring ptr) as zstring ptr
declare sub splitstr(Text As String, Delim As String = ",", Ret() As String)
declare function strreplace(byref text as string, byref a as string, byref b as string) as string
declare function UF_SIGLA2COD(s as zstring ptr) as integer
declare function loadstrings(fromFile as string, toArray() as string) as boolean

extern as string ufCod2Sigla(11 to 53)
extern as TDict ufSigla2CodDict
extern as string codSituacao2Str(0 to __TipoSituacao__LEN__-1)

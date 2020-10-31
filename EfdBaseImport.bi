#include once "Dict.bi"
#include once "bfile.bi"
#include once "Db.bi"

enum TTipoArquivo
	TIPO_ARQUIVO_EFD
	TIPO_ARQUIVO_SINTEGRA
end enum

enum TipoRegistro
	MESTRE
	PARTICIPANTE
	ITEM_ID
	BEM_CIAP
	BEM_CIAP_INFO
	CONTA_CONTAB
	CENTRO_CUSTO
	INFO_COMPL
	OBS_LANCAMENTO
	DOC_NF										'' NF, NF-e, NFC-e
	DOC_NF_INFO									'' informações complementares de interesse do fisco
	DOC_NF_ITEM    								'' item de NF-e (só informado para entradas)
	DOC_NF_ITEM_RESSARC_ST						'' ressarcimento ST
	DOC_NF_ANAL									'' analítico
	DOC_NF_OBS
	DOC_NF_OBS_AJUSTE
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
	APURACAO_ICMS_PROPRIO_OBRIG	
	APURACAO_ICMS_ST_PERIODO	
	APURACAO_ICMS_ST
	APURACAO_ICMS_AJUSTE
	APURACAO_ICMS_AJUSTE_EXTRA
	INVENTARIO_TOTAIS
	INVENTARIO_ITEM
	CIAP_TOTAL
	CIAP_ITEM
	CIAP_ITEM_DOC
	CIAP_ITEM_DOC_ITEM
	ESTOQUE_PERIODO
	ESTOQUE_ITEM
	ESTOQUE_ORDEM_PROD
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

type TRegistro
	tipo           			as TipoRegistro
	linha					as integer
	prox          			as TRegistro ptr
end type

type TMestre extends TRegistro
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

type TParticipante extends TRegistro
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
	ENTRADA		  = 0
	SAIDA		  = 1
	DESCONHECIDA  = 2
end enum

enum TipoEmitente
	PROPRIO		  = 0
	TERCEIRO	  = 1
end enum

enum TipoModelo
	INVALIDO	   = 0
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

type TItemId extends TRegistro
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

type TInfoCompl extends TRegistro
	id             	as zstring * 6+1
	descricao      	as zstring * 256+1
end type

type TObsLancamento extends TRegistro
	id             	as zstring * 6+1
	descricao      	as zstring * 256+1
end type

type TBemCiap extends TRegistro
	id             	as zstring * 60+1
	tipoMerc		as integer
	descricao      	as zstring * 256+1
	principal	   	as zstring * 60+1
	codAnal			as zstring * 60+1
	parcelas		as integer
	codCusto		as zstring * 60+1
	funcao			as zstring * 256+1
	vidaUtil		as integer
end type

type TContaContab extends TRegistro
	id             	as zstring * 60+1
	descricao      	as zstring * 256+1
	dataInc			as zstring * 8+1
	codNat			as zstring * 2+1
	ind				as zstring * 1+1
	nivel			as integer
end type

type TCentroCusto extends TRegistro
	id             	as zstring * 60+1
	descricao      	as zstring * 256+1
	dataInc			as zstring * 8+1
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

type TDocNFItemRessarcSt extends TRegistro
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

type TDocNFItem extends TRegistro
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

type TDocECF_ as TDocECF

type TDocECFItem extends TRegistro
	documentoPai   as TDocECF_ ptr
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

type TDocItemAnal extends TRegistro
	documentoPai   			as TRegistro ptr
	num						as integer
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

type TDocInfoCompl extends TRegistro
	idCompl					as zstring * 6+1
	extra					as zstring * 255+1
	next_					as TDocInfoCompl ptr
end type

type TDocObsAjuste extends TRegistro
	idAjuste				as zstring * 10+1
	extra					as zstring * 255+1
	idItem					as zstring * 60+1
	bcICMS					as double
	aliqICMS				as double
	icms					as double
	outros					as double
	next_					as TDocObsAjuste ptr
end type

type TDocObs extends TRegistro
	idLanc					as zstring * 6+1
	extra					as zstring * 255+1
	ajusteListHead 			as TDocObsAjuste ptr
	ajusteListTail 			as TDocObsAjuste ptr
	next_					as TDocObs ptr
end type

type TDocDF extends TRegistro
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
	obsListHead 			as TDocObs ptr
	obsListTail 			as TDocObs ptr
	itemAnalCnt				as integer
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

type TDocumentoSintegraBase extends TRegistro
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

const MAX_AJUSTES = 20

type TApuracaoIcmsAjuste extends TRegistro
	codigo					as zstring * 8+1
	descricao				as zstring * 255+1
	valor					as double
	next_					as TApuracaoIcmsAjuste ptr
end type

type TApuracaoIcmsPeriodo extends TRegistro
	dataIni					as zstring * 8+1
	dataFim					as zstring * 8+1
	saldoCredAnterior		as double
	ajustesDebitos			as double
	ajustesCreditos			as double
	totalDeducoes			as double
	icmsRecolher			as double
	saldoCredTransportar	as double
	debExtraApuracao		as double
	ajustesListHead 		as TApuracaoIcmsAjuste ptr
	ajustesListTail 		as TApuracaoIcmsAjuste ptr
end type

type TApuracaoIcmsPropPeriodo extends TApuracaoIcmsPeriodo
	totalDebitos			as double
	totalAjusteDeb			as double
	estornosCredito			as double
	totalCreditos			as double
	totalAjusteCred			as double
	estornoDebitos			as double
	saldoDevedorApurado		as double
end type

type TApuracaoIcmsSTPeriodo extends TApuracaoIcmsPeriodo
	UF						as zstring * 2+1
	mov						as boolean
	devolMercadorias		as double
	totalRessarciment		as double
	totalOutrosCred			as double
	totalRetencao			as double
	totalOutrosDeb			as double
	saldoAntesDed			as double
end type

type TEquipECF extends TRegistro
	modelo					as TipoModelo
	modeloEquip				as zstring * 20+1
	numSerie				as zstring * 21+1
	numCaixa				as integer
end type

type TECFReducaoZ extends TDocDF
	equipECF				as TEquipECF ptr
	dataMov					as zstring * 8+1
	cro						as longint
	crz						as longint
	numOrdem				as longint
	valorFinal				as double
	valorBruto				as double
	numIni					as integer
	numFim					as integer
end type

type TInventarioTotais extends TRegistro
	dataInventario			as zstring * 8+1
	valorTotalEstoque		as double
	motivoInventario		as integer
end type

type TInventarioItem extends TRegistro
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

type TCiapItem_ as TCiapItem

type TCiapTotal extends TRegistro
	dataIni					as zstring * 8+1
	dataFim					as zstring * 8+1
	saldoInicialICMS		as double
	parcelasSoma			as double
	valorTributExpSoma		as double
	valorTotalSaidas		as double
	indicePercSaidas		as double
	valorIcmsAprop			as double
	valorOutrosCred			as double
	itemListHead 			as TCiapItem_ ptr
	itemListTail 			as TCiapItem_ ptr
end type

type TCiapItemDoc_ as TCiapItemDoc

type TCiapItem extends TRegistro
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
	next_					as TCiapItem ptr
	docListHead 			as TCiapItemDoc_ ptr
	docListTail 			as TCiapItemDoc_ ptr
end type

type TCiapItemDocItem_ as TCiapItemDocItem

type TCiapItemDoc extends TRegistro
	pai         			as TCiapItem ptr
	indEmi					as integer
	idParticipante			as zstring * 60+1
	modelo					as integer
	serie					as zstring * 3+1
	numero					as integer
	chaveNfe				as zstring * 44+1
	dataEmi					as zstring * 8+1
	next_					as TCiapItemDoc ptr
	itemListHead 			as TCiapItemDocItem_ ptr
	itemListTail 			as TCiapItemDocItem_ ptr
end type

type TCiapItemDocItem extends TRegistro
	pai         			as TCiapItemDoc ptr
	num						as integer
	itemId         			as zstring * 60+1
	next_					as TCiapItemDocItem ptr
end type

enum TipoItemEstoque
	PROPRIO_PROPRIO
	PROPRIO_TERCEIRO
	TERCEIRO_PROPRIO
end enum

type TEstoquePeriodo extends TRegistro
	dataIni					as zstring * 8+1
	dataFim					as zstring * 8+1
end type

type TEstoqueItem extends TRegistro
	pai         			as TEstoquePeriodo ptr
	itemId					as zstring * 60+1
	qtd						as double
	tipoEst					as TipoItemEstoque
	idParticipante			as zstring * 60+1
end type

type TEstoqueOrdemProd extends TRegistro
	pai         			as TEstoquePeriodo ptr
	dataIni					as zstring * 8+1
	dataFim					as zstring * 8+1
	idOrdem					as zstring * 30+1
	itemId					as zstring * 60+1
	qtd						as double
end type

type TLuaReg extends TRegistro
	tipo					as zstring * 4+1
	table					as integer
end type

type InfoAssinatura
	assinante		as string
	cpf				as string
	hashDoArquivo	as string
end type

type EfdBaseImport extends object
public:
	declare constructor()
	declare constructor(opcoes as OpcoesExtracao ptr)
	declare destructor()
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdBaseImport ptr
	declare function withLua(lua as lua_State ptr, customLuaCbDict as TDict ptr) as EfdBaseImport ptr
	declare function withDBs(db as TDb ptr) as EfdBaseImport ptr
	declare abstract function carregar(nomeArquivo as string) as boolean
	declare function getFirstReg() as TRegistro ptr
	declare function getMestreReg() as TMestre ptr
	declare function getNroRegs() as integer
	declare function getParticipanteDict() as TDict ptr
	declare function getItemIdDict as TDict ptr
	declare function getInfoComplDict as TDict ptr
	declare function getObsLancamentoDict as TDict ptr
	declare function getBemCiapDict as TDict ptr
	declare function getContaContabDict as TDict ptr
	declare function getCentroCustoDict as TDict ptr
	declare function getTipoArquivo() as TTipoArquivo

protected:
	opcoes					as OpcoesExtracao ptr
	db						as TDb ptr

	onProgress 				as OnProgressCB
	onError 				as OnErrorCB
	
	lua						as lua_State ptr
	customLuaCbDict			as TDict ptr		'' de CustomLuaCb
	
	tipoArquivo				as TTipoArquivo
	regListHead         	as TRegistro ptr = null
	regMestre				as TMestre ptr
	nroRegs             	as integer = 0
	nroLinha				as integer
	
	participanteDict		as TDict ptr
	itemIdDict          	as TDict ptr
	infoComplDict			as TDict ptr
	obsLancamentoDict		as TDict ptr
	bemCiapDict          	as TDict ptr
	contaContabDict			as TDict ptr
	centroCustoDict			as TDict ptr
end type


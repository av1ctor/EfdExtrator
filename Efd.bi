
#include once "hash.bi"
#include once "bfile.bi"
#include once "ExcelWriter.bi"
#include once "DocxFactoryDyn.bi"

enum TTipoArquivo
	TIPO_ARQUIVO_EFD
	TIPO_ARQUIVO_SINTEGRA
end enum

enum TipoRegistro
	MESTRE         				= &h0000
	PARTICIPANTE   				= &h0150
	ITEM_ID        				= &h0200
	DOC_NFE      				= &hC100		'' NF, NF-e, NFC-e
	DOC_NFE_ITEM    			= &hC170		'' item de NF-e (só informado para entradas)
	DOC_NFE_ANAL				= &hC190
	DOC_NFE_DIFAL				= &hC101
	DOC_CTE     				= &hD100		'' CT, CT-e, CT-e OS, BP-e
	DOC_CTE_DIFAL				= &hD101
	DOC_CTE_ITEM				= &hD190		'' item de CT-e  (só informado para entradas)
	APURACAO_ICMS_PERIODO		= &hE100
	APURACAO_ICMS_PROPRIO		= &hE110
	APURACAO_ICMS_AJUSTE		= &hE111
	APURACAO_ICMS_PROPRIO_OBRIG	= &hE116
	APURACAO_ICMS_ST_PERIODO	= &hE200
	APURACAO_ICMS_ST			= &hE210
	EOF_   						= &h9999		'' NOTA: anterior à assinatura digital que fica no final no arquivo
	DESCONHECIDO   				= &h8888
	SINTEGRA_DOCUMENTO 			= 50			'' NOTA: códigos do Sintegra não conflitam com outros tipos de registros na EFD
	SINTEGRA_DOCUMENTO_IPI 		= 51
	SINTEGRA_DOCUMENTO_ST		= 53
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
	cpf            		as longint
	ie			      	as zstring * 14+1
	municip		   		as integer
	suframa		   		as zstring * 9+1
	ender		      	as zstring * 60+1
	num			   		as zstring * 10+1
	compl		      	as zstring * 60+1
	bairro		   		as zstring * 60+1
end type

enum TipoOperacao
	ENTRADA		  = 0	'' NF-e
	SAIDA		  = 1	'' NF-e
	AQUISICAO	  = 0	'' CT-e
	PRESTACAO	  = 1	'' CT-e
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
end type

type TDocNFe_ as TDocNFe ptr

type TDocNFeItem                       ' nota: só é obrigatório para entradas!!!
	documentoPai   as TDocNFe_
	numItem        as Integer
	itemId         as zstring * 60+1
	descricao      as zstring * 256+1
	qtd            as double
	unidade        as zstring * 6+1
	valor          as Double
	desconto       as double
	indMovFisica   as byte
	cstICMS        as integer
	cfop           as Integer
	codNatureza    as zstring * 10+1
	bcICMS         as Double
	aliqICMS       as double
	ICMS           as Double
	bcICMSST       as Double
	aliqICMSST     as Double
	ICMSST         as Double
	indApuracao    as Byte
	cstIPI         as Integer
	codEnqIPI      as zstring * 2+1
	bcIPI          as double
	aliqIPI        as Double
	IPI            as Double
	cstPIS         as integer
	bcPIS          as Double
	aliqPISPerc    as Double
	qtdBcPIS       as double
	aliqPISMoed    as Double
	PIS            as Double
	cstCOFINS      as Integer
	bcCOFINS       as Double
	aliqCOFINSPerc as Double
	qtdBcCOFINS    as double
	aliqCOFINSMoed as Double
	COFINS         as Double
end type

type TDocDifAliq
	fcp				as double
	icmsDest		as double
	icmsOrigem		as double
end type

type TRegistro_ as TRegistro ptr

type TDocItemAnal
	documentoPai   	as TRegistro_
	cst				as integer
	cfop			as integer
	aliq			as double
	valorOp			as double
	bc				as double
	ICMS			as double
	bcST			as double
	ICMSST			as double
	redBC			as double
	IPI				as double
	next_			as TDocItemAnal ptr
end type

type TDocNFe
	operacao		as TipoOperacao
	emitente		as TipoEmitente
	idParticipante	as zstring * 60+1
	modelo			as TipoModelo
	situacao		as TipoSituacao
	serie			as integer
	numero			as longint
	chave			as zstring * 44+1
	dataEmi			as zstring * 8+1		'DDMMAAAA
	dataEntSaida	as zstring * 8+1		'DDMMAAAA
	valorTotal		as double
	pagamento		as TipoPagamento
	valorDesconto	as double
	valorAbatimento as double
	valorMerc		as double
	frete			as TipoFrete
	valorFrete		as double
	valorSeguro		as double
	valorAcessorias as double
	bcICMS			as double
	ICMS			as double
	bcICMSST		as double
	ICMSST			as double
	IPI				as double
	PIS				as double
	COFINS			as double
	PISST			as double
	COFINSST		as double
	difal			as TDocDifAliq
	nroItens		as integer
	
	itemAnalListHead as TDocItemAnal ptr
	itemAnalListTail as TDocItemAnal ptr
end type

type TDocCTe_ as TDocCTe ptr

type TDocCTeItem
	documentoPai   	as TDocCTe_
	cstICMS        	as integer
	cfop           	as Integer
	aliqICMS       	as double
	valorOperacao	as double
	bcICMS         	as double
	ICMS           	as double
	reducaoBcICMS	as double
	codObs       	as zstring * 6+1
	next_			as TDocCTeItem ptr
end type

type TDocCTe
	operacao			as TipoOperacao
	emitente			as TipoEmitente
	idParticipante		as zstring * 60+1
	modelo				as TipoModelo
	situacao			as TipoSituacao
	serie				as integer
	numero				as longint
	chave				as zstring * 44+1
	dataEmi				as zstring * 8+1		'DDMMAAAA
	dataAquPrest		as zstring * 8+1		'DDMMAAAA
	tipoCTe				as integer
	chaveRef			as zstring * 44+1		'' para CT-e do tipo complementar, substituto ou anulador
	valorTotal			as double
	valorDesconto		as double
	frete				as TipoFrete
	valorServico		as double
	bcICMS				as double
	ICMS				as double
	valorNaoTributado	as double
	codInfComplementar	as zstring * 6+1
	municipioOrigem		as integer
	municipioDestino	as integer
	difal				as TDocDifAliq
	nroItens			as integer
	itemListHead		as TDocCTeItem ptr
	itemListTail		as TDocCTeItem ptr
end type

type TDocumentoSintegra
	cnpj           	as zstring * 14+1
	ie             	as zstring * 14+1
	dataEmi        	as zstring * 8+1
	uf             	as zstring * 2+1
	modelo		  	as TipoModelo
	serie          	as short
	numero         	as integer
	cfop           	as short
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
	chaveHash	  	as zstring * 50+1
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

type TRegistro
	tipo           	as TipoRegistro
	union
		mestre      as TMestre
		part        as TParticipante
		nfe         as TDocNFe
		itemNFe     as TDocNFeItem
		cte         as TDocCTe
		itemCTe     as TDocCTeItem
		docSint	  	as TDocumentoSintegra
		itemId      as TItemId
		apuIcms	  	as TApuracaoIcmsPeriodo
		apuIcmsST  	as TApuracaoIcmsSTPeriodo
		itemAnal	as TDocItemAnal
	end union
	next_          	as TRegistro ptr
end type

enum SAFI_TipoArquivo
	SAFI_NFe_Dest
	SAFI_NFe_Emit
	SAFI_NFe_Emit_Itens
	SAFI_CTe
end enum

type TDFe_NFeItem									'' Nota: só existe para NF-e emitidas, já que para as recebidas os itens constam na EFD
	cfop			as short
	nroItem			as integer
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
	serie			as integer
	numero			as integer
	cnpjEmit		as zstring * 14+1
	nomeEmit		as zstring * 100+1
	ieEmit			as zstring * 14+1
	ufEmit			as zstring * 2+1
	cnpjDest		as zstring * 14+1
	nomeDest		as zstring * 100+1
	ufDest			as zstring * 2+1
	bcICMSTotal		as double
	ICMSTotal		as double
	bcICMSSTTotal	as double
	ICMSSTTotal		as double
	valorNotaTotal	as double
	
	itemListHead	as TDFe_NFeItem ptr
	itemListTail	as TDFe_NFeItem ptr
end type

type TDFe_ as TDFe ptr

type TDFe_CTe
	serie			as integer
	numero			as integer
	cnpjEmit		as zstring * 14+1
	nomeEmit		as zstring * 100+1
	ufEmit			as zstring * 2+1
	cnpjDest		as zstring * 14+1
	nomeDest		as zstring * 100+1
	ufDest			as zstring * 2+1
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
	valorPrestacao	as double
	valorReceber	as double
	qtdCTe			as double
	cfop			as integer
	nomeMunicIni	as zstring * 64+1
	ufIni			as zstring * 2+1
	nomeMunicFim	as zstring * 64+1
	ufFim			as zstring * 2+1
	next_			as TDFe_CTe ptr					'' usado para dar patch 
	parent_			as TDFe_
end type

type TDFe
	modelo			as TipoModelo
	operacao		as TipoOperacao					'' entrada ou saída
	chave			as zstring * 44+1
	dataEmi			as zstring * 10+1
	valorOperacao	as double
	
	union
		nfe			as TDFe_NFe
		cte			as TDFe_CTe
	end union
	
	next_			as TDFe ptr
end type

type TEfd_DFe
	chave			as zstring * 44+1
	dataEmi			as zstring * 8+1
	operacao		as TipoOperacao
	valorOperacao	as double
	next_			as TEfd_DFe ptr
end type

type InfoAssinatura
	assinante		as string
	cpf				as string
	hashDoArquivo	as string
end type

type Efd
public:
	declare constructor ()
	declare destructor ()
	declare sub iniciarExtracao(nomeArquivo as String)
	declare sub finalizarExtracao(mostrarProgresso as sub(porCompleto as double))
	declare function carregarTxt(nomeArquivo as String, mostrarProgresso as sub(porCompleto as double)) as Boolean
	declare function carregarCsv(nomeArquivo as String, mostrarProgresso as sub(porCompleto as double)) as Boolean
	declare function processar(nomeArquivo as string, mostrarProgresso as sub(porCompleto as double), gerarRelatorios as boolean) as Boolean
	declare sub analisar(mostrarProgresso as sub(porCompleto as double))
   
private:
	tipoArquivo				as TTipoArquivo
	
	'' registros das EFD's e do Sintegra (reiniciados a cada novo .txt carregado)
	regListHead         	as TRegistro ptr = null
	regListTail         	as TRegistro ptr = null
	nroRegs             	as integer = 0
	participanteDict    	as THASH
	itemIdDict          	as THASH
	sintegraDict			as THASH
	ultimoReg   			as TRegistro ptr

	'' registros para cruzamento das EFD's com as NF-e/CT-e (mantidos do início ao fim da extração)
	efdDFeDict				as THASH
	efdDFeListHead			as TEfd_DFe ptr
	efdDFeListTail			as TEfd_DFe ptr

	'' planilhas que serão geradas (mantidos do início ao fim da extração)
	ew                  	as ExcelWriter ptr
	entradas            	as ExcelWorksheet ptr
	saidas              	as ExcelWorksheet ptr
	naoEscrituradas			as ExcelWorksheet ptr
	apuracaoIcms			as ExcelWorksheet ptr
	apuracaoIcmsST			as ExcelWorksheet ptr

	'' registros das NF-e's e CT-e's retirados dos relatórios do Infoview (mantidos do início ao fim da extração)
	chaveDFeDict			as THASH
	dfeListHead				as TDFe ptr = null
	dfeListTail				as TDFe ptr = null
	nroDfe					as integer = 0
	cteListHead				as TDFe_CTe ptr = null	'' usado para fazer patch no tipo de operação
	cteListTail				as TDFe_CTe ptr = null
	nfeDestSafiFornecido 	as boolean
	nfeEmitSafiFornecido 	as boolean
	itemNFeSafiFornecido 	as boolean
	cteSafiFornecido		as boolean
	
	'' geração de relatórios em formato PDF com o layout do programa EFD-ICMS-IPI da RFB
	baseTemplatesDir		as string
	dfwd					as DocxFactoryDyn ptr
	
	''
	assinaturaP7K_DER(any)	as byte
	infAssinatura			as InfoAssinatura ptr

	declare function lerRegistro(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegistroSintegra(bf as bfile, reg as TRegistro ptr) as Boolean
	declare sub lerAssinatura(bf as bfile)
	declare function carregarSintegra(bf as bfile, mostrarProgresso as sub(porCompleto as double)) as Boolean
	declare function carregarCsvNFeDest(bf as bfile, emModoOutrasUFs as boolean) as TDFe ptr
	declare function carregarCsvNFeEmit(bf as bfile) as TDFe ptr
	declare function carregarCsvNFeEmitItens(bf as bfile, chave as string) as TDFe_NFeItem ptr
	declare function carregarCsvCTe(bf as bfile, emModoOutrasUFs as boolean) as TDFe ptr
	declare sub adicionarDFe(dfe as TDFe ptr)
	declare sub adicionarEfdDfe(chave as zstring ptr, operacao as TipoOperacao, dataEmi as zstring ptr, valorOperacao as double)
	declare sub criarPlanilhas()
	declare sub gerarRelatorioApuracaoICMS(nomeArquivo as string, reg as TRegistro ptr)
	declare sub gerarRelatorioApuracaoICMSST(nomeArquivo as string, reg as TRegistro ptr)
	declare sub iniciarRelatorioSaidas(nomeArquivo as string)
	declare sub adicionarDocRelatorioSaidas(doc as TDocNFe ptr, part as TParticipante ptr)
	declare sub finalizarRelatorioSaidas(nomeArquivo as string)
end type


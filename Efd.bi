
#include once "Dict.bi"
#include once "bfile.bi"
#include once "ExcelReader.bi"
#include once "ExcelWriter.bi"
#include once "DB.bi"
#include once "Lua/lualib.bi"
#include once "PDFer.bi"

type OnProgressCB as function(estagio as const zstring ptr, porCompleto as double) as boolean
type OnErrorCB as sub(msg as const zstring ptr)
type OnFilterByStrCB as function(key as const zstring ptr, arr() as string) as boolean

type OpcoesExtracao
	gerarRelatorios 				as boolean = false
	pularLre 						as boolean = false
	pularLrs 						as boolean = false
	pularLraicms					as boolean = false
	pularCiap						as boolean = false
	pularAnalises					as boolean = false
	pularResumos					as boolean = false
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

#include once "EfdBaseImport.bi"
type EfdBaseImport_ as EfdBaseImport

enum BO_TipoArquivo
	BO_NFe_Dest
	BO_NFe_Emit
	BO_NFe_Emit_Itens
	BO_CTe
	BO_SAT
	BO_SAT_Itens
	BO_NFCe_Itens
	SAFI_ECF
	BO_ECF_Itens
end enum

enum BO_Dfe_Fornecido
	MASK_BO_NFe_DEST_FORNECIDO = &b00000001
	MASK_BO_NFe_EMIT_FORNECIDO = &b00000010
	MASK_BO_ITEM_NFE_FORNECIDO = &b00000100
	MASK_BO_CTe_FORNECIDO 	 = &b00001000
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
	aliqIcmsST		as double
	icmsST			as double
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

enum TipoLivro
	TL_ENTRADAS
	TL_SAIDAS
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

#include once "EfdTabelaExport.bi"
type EfdTabelaExport_ as EfdTabelaExport

type Efd
public:
	declare constructor (onProgress as OnProgressCB, onError as OnErrorCB)
	declare destructor ()
	declare sub iniciar(nomeArquivo as String, opcoes as OpcoesExtracao)
	declare sub finalizar()
	declare function carregarTxt(nomeArquivo as String) as EfdBaseImport_ ptr
	declare function carregarCsv(nomeArquivo as String) as Boolean
	declare function carregarXlsx(nomeArquivo as String) as Boolean
	declare function processar(imp_ as EfdBaseImport_ ptr, nomeArquivo as string) as Boolean
	declare sub analisar()
	declare sub resumir()
	declare sub descarregarDFe()

	exp						as EfdTabelaExport_ ptr
	onProgress 				as OnProgressCB
	onError 				as OnErrorCB
   
private:
	declare sub configurarDB()
	declare sub fecharDb()
	declare sub configurarScripting()
	
	declare function carregarCsvNFeDestSAFI(bf as bfile, emModoOutrasUFs as boolean) as TDFe ptr
	declare function carregarCsvNFeEmitSAFI(bf as bfile) as TDFe ptr
	declare function carregarCsvNFeEmitItensSAFI(bf as bfile, chave as string) as TDFe_NFeItem ptr
	declare function carregarCsvCTeSAFI(bf as bfile, emModoOutrasUFs as boolean) as TDFe ptr
	declare function carregarCsvNFeEmitItens(bf as bfile, chave as string, extra as TDFe ptr) as TDFe_NFeItem ptr
	
	declare function carregarXlsxNFeDest(reader as ExcelReader ptr) as TDFe ptr
	declare function carregarXlsxNFeDestItens(reader as ExcelReader ptr) as TDFe ptr
	declare function carregarXlsxNFeEmit(rd as ExcelReader ptr) as TDFe ptr
	declare function carregarXlsxNFeEmitItens(rd as ExcelReader ptr, chave as string, extra as TDFe ptr) as TDFe_NFeItem ptr
	declare function carregarXlsxCTe(rd as ExcelReader ptr, op as TipoOperacao) as TDFe ptr
	declare function carregarXlsxSAT(rd as ExcelReader ptr) as TDFe ptr
	declare function carregarXlsxSATItens(rd as ExcelReader ptr, chave as string) as TDFe_NFeItem ptr
	
	declare function adicionarDFe(dfe as TDFe ptr, fazerInsert as boolean = true) as long
	declare function adicionarItemDFe(chave as const zstring ptr, item as TDFe_NFeItem ptr) as long
	declare function adicionarEfdDfe(chave as zstring ptr, operacao as TipoOperacao, dataEmi as zstring ptr, valorOperacao as double) as long
	
	declare sub exportAPI(L as lua_State ptr)
	declare static function luacb_efd_participante_get cdecl(L as lua_State ptr) as long
	
	declare function getDfeMask() as long

	'' dicionários
	municipDict				as TDict ptr
	chaveDFeDict			as TDict ptr
	
	''
	nomeArquivoSaida		as string
	opcoes					as OpcoesExtracao
	baseTemplatesDir		as string

	'' registros das NF-e's e CT-e's retirados dos relatórios do Infoview (mantidos do início ao fim da extração)
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
	configDb				as TDb ptr
	
	'' base de dados temporária usadada para análises e cruzamentos
	db						as TDb ptr
	db_dfeEntradaInsertStmt	as TDbStmt ptr
	db_dfeSaidaInsertStmt	as TDbStmt ptr
	db_itensDfeSaidaInsertStmt as TDbStmt ptr
	db_LREInsertStmt		as TDbStmt ptr
	db_itensNfLRInsertStmt	as TDbStmt ptr
	db_LRSInsertStmt		as TDbStmt ptr
	db_analInsertStmt		as TDbStmt ptr
	db_ressarcStItensNfLRSInsertStmt as TDbStmt ptr
	db_itensIdInsertStmt 	as TDbStmt ptr
	db_mestreInsertStmt 	as TDbStmt ptr
	
	'' scripting
	lua						as lua_State ptr
	customLuaCbDict			as TDict ptr		'' de CustomLuaCb
end type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#define DdMmYyyy2Yyyy_Mm(s) (mid(s,1,4) + "-" + mid(s,5,2))

#define STR2CNPJ(s) iif(len(s) > 0, left(s,2) + "." + mid(s,3,3) + "." + mid(s,3+3,3) + "/" + mid(s,3+3+3,4) + "-" + right(s,2), "")

#define STR2CPF(s) (left(s,3) + "." + mid(s,4,3) + "." + mid(s,4+3,3) + "-" + right(s,2))

#define DBL2MONEYBR(d) (format(d,"#,#,0.00"))

#define MUNICIPIO2SIGLA(m) (iif(m >= 1100000 and m <= 5399999, ufCod2Sigla(m \ 100000), "EX"))

declare function ISREGULAR(sit as TipoSituacao) as boolean
declare function csvDate2YYYYMMDD(s as zstring ptr) as string 
declare function ddMmYyyy2YyyyMmDd(s as const zstring ptr) as string
declare function yyyyMmDd2YyyyMmDd(s as const zstring ptr) as string
declare function yyyyMmDd2Datetime(s as const zstring ptr) as string 
declare function YyyyMmDd2DatetimeBR(s as const zstring ptr) as string 
declare function STR2IE(ie as string) as string
declare function tipoItem2Str(tipo as TipoItemId) as string
declare function UF_SIGLA2COD(s as zstring ptr) as integer
declare function codMunicipio2Nome(cod as integer, municipDict as TDict ptr, configDb as TDb ptr) as string
declare sub pularLinha(bf as bfile)
declare function lerLinha(bf as bfile) as string
declare sub lua_setarGlobal overload (lua as lua_State ptr, varName as const zstring ptr, value as integer)
declare sub lua_setarGlobal overload (lua as lua_State ptr, varName as const zstring ptr, value as any ptr)
declare function filtrarPorCnpj(idParticipante as const zstring ptr, listaCnpj() as string) as boolean
declare function filtrarPorChave(chave as const zstring ptr, listaChaves() as string) as boolean

extern as string ufCod2Sigla(11 to 53)
extern as TDict ptr ufSigla2CodDict
extern as string codSituacao2Str(0 to __TipoSituacao__LEN__-1)

#include once "strings.bi"


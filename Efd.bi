
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

#include once "EfdBoBaseLoader.bi"

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
	loaderCtx				as EfdBoLoaderContext ptr
	onProgress 				as OnProgressCB
	onError 				as OnErrorCB
   
private:
	declare sub configurarDB()
	declare sub fecharDb()
	declare sub configurarScripting()
	
	declare sub exportAPI(L as lua_State ptr)
	declare static function luacb_efd_participante_get cdecl(L as lua_State ptr) as long
	
	declare function getDfeMask() as long

	'' dicionários
	municipDict				as TDict ptr
	
	''
	nomeArquivoSaida		as string
	opcoes					as OpcoesExtracao
	baseTemplatesDir		as string

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


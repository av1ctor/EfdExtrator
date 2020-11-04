
#include once "Dict.bi"
#include once "bfile.bi"
#include once "ExcelReader.bi"
#include once "TableWriter.bi"
#include once "SQLite.bi"
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

type EfdExt
public:
	declare constructor (onProgress as OnProgressCB, onError as OnErrorCB)
	declare destructor ()
	declare function iniciar(nomeArquivo as String, opcoes as OpcoesExtracao) as boolean
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
	configDb				as SQLite ptr
	
	'' base de dados temporária usadada para análises e cruzamentos
	dbName					as string
	db						as SQLite ptr
	db_dfeEntradaInsertStmt	as SQLiteStmt ptr
	db_dfeSaidaInsertStmt	as SQLiteStmt ptr
	db_itensDfeSaidaInsertStmt as SQLiteStmt ptr
	db_LREInsertStmt		as SQLiteStmt ptr
	db_itensNfLRInsertStmt	as SQLiteStmt ptr
	db_LRSInsertStmt		as SQLiteStmt ptr
	db_analInsertStmt		as SQLiteStmt ptr
	db_ressarcStItensNfLRSInsertStmt as SQLiteStmt ptr
	db_itensIdInsertStmt 	as SQLiteStmt ptr
	db_mestreInsertStmt 	as SQLiteStmt ptr
	
	'' scripting
	lua						as lua_State ptr
	customLuaCbDict			as TDict ptr		'' de CustomLuaCb
end type

#include once "misc.bi"
#include once "strings.bi"


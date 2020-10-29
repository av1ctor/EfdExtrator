#include once "Efd.bi"
#include once "ExcelWriter.bi"

type EfdTabelaExport
public:
	declare constructor(nomeArquivo as String, opcoes as OpcoesExtracao ptr)
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdTabelaExport ptr
	declare function withLua(lua as lua_State ptr, customLuaCbDict as TDict ptr) as EfdTabelaExport ptr
	declare function withState(itemNFeSafiFornecido as boolean) as EfdTabelaExport ptr
	declare function withDicionarios(participanteDict as TDict ptr, itemIdDict as TDict ptr, chaveDFeDict as TDict ptr, infoComplDict as TDict ptr, obsLancamentoDict as TDict ptr, bemCiapDict as TDict ptr) as EfdTabelaExport ptr
	declare function withFiltros(filtrarPorCnpj as OnFilterByStrCB, filtrarPorChave as OnFilterByStrCB) as EfdTabelaExport ptr
	declare destructor()
	declare function getPlanilha(nome as const zstring ptr) as ExcelWorksheet ptr
	declare sub gerar(regListHead as TRegistro ptr, regMestre as TRegistro ptr, nroRegs as integer)
	declare sub finalizar()

private:
	nomeArquivo				as string
	opcoes					as OpcoesExtracao ptr
	itemNFeSafiFornecido	as boolean
	
	participanteDict		as TDict ptr
	itemIdDict          	as TDict ptr
	infoComplDict			as TDict ptr
	obsLancamentoDict		as TDict ptr
	bemCiapDict          	as TDict ptr
	chaveDFeDict          	as TDict ptr

	onProgress 				as OnProgressCB
	onError 				as OnErrorCB
	filtrarPorCnpj 			as OnFilterByStrCB
	filtrarPorChave			as OnFilterByStrCB
	
	lua						as lua_State ptr
	customLuaCbDict			as TDict ptr		'' de CustomLuaCb
	
	ew                  	as ExcelWriter ptr
	entradas            	as ExcelWorksheet ptr
	saidas              	as ExcelWorksheet ptr
	apuracaoIcms			as ExcelWorksheet ptr
	apuracaoIcmsST			as ExcelWorksheet ptr
	inventario				as ExcelWorksheet ptr
	ciap					as ExcelWorksheet ptr
	estoque					as ExcelWorksheet ptr
	producao				as ExcelWorksheet ptr
	ressarcST				as ExcelWorksheet ptr
	inconsistenciasLRE		as ExcelWorksheet ptr
	inconsistenciasLRS		as ExcelWorksheet ptr
	resumosLRE				as ExcelWorksheet ptr
	resumosLRS				as ExcelWorksheet ptr

	declare sub criarPlanilhas()
	declare function getInfoCompl(info as TDocInfoCompl ptr) as string
	declare function getObsLanc(obs as TDocObs ptr) as string
end type
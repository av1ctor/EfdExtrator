#include once "Efd.bi"
#include once "TableWriter.bi"

type EfdTabelaExport
public:
	declare constructor(nomeArquivo as String, opcoes as OpcoesExtracao ptr)
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdTabelaExport ptr
	declare function withLua(lua as lua_State ptr, customLuaCbDict as TDict ptr) as EfdTabelaExport ptr
	declare function withState(itemNFeSafiFornecido as boolean) as EfdTabelaExport ptr
	declare function withDicionarios(participanteDict as TDict ptr, itemIdDict as TDict ptr, chaveDFeDict as TDict ptr, infoComplDict as TDict ptr, obsLancamentoDict as TDict ptr, bemCiapDict as TDict ptr) as EfdTabelaExport ptr
	declare function withFiltros(filtrarPorCnpj as OnFilterByStrCB, filtrarPorChave as OnFilterByStrCB) as EfdTabelaExport ptr
	declare destructor()
	declare function criar() as boolean
	declare function getPlanilha(nome as const zstring ptr) as TableTable ptr
	declare sub gerar(regListHead as TRegistro ptr, regMestre as TMestre ptr, nroRegs as integer)
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
	
	ew                  	as TableWriter ptr
	entradas            	as TableTable ptr
	saidas              	as TableTable ptr
	apuracaoIcms			as TableTable ptr
	apuracaoIcmsST			as TableTable ptr
	inventario				as TableTable ptr
	ciap					as TableTable ptr
	estoque					as TableTable ptr
	producao				as TableTable ptr
	ressarcST				as TableTable ptr
	inconsistenciasLRE		as TableTable ptr
	inconsistenciasLRS		as TableTable ptr
	resumosLRE				as TableTable ptr
	resumosLRS				as TableTable ptr

	declare sub criarPlanilhas()
	declare function getInfoCompl(info as TDocInfoCompl ptr) as string
	declare function getObsLanc(obs as TDocObs ptr) as string
end type
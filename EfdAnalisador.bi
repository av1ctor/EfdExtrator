#include once "Efd.bi"

type EfdAnalisador
public:
	declare constructor(tableExp as EfdTabelaExportador ptr)
	declare function withDBs(db as TDb ptr) as EfdAnalisador ptr
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdAnalisador ptr
	declare function withLua(lua as lua_State ptr) as EfdAnalisador ptr
	declare sub executar(safiFornecidoMask as long) 
	
private:
	db						as TDb ptr
	tableExp				as EfdTabelaExportador ptr
	onProgress 				as OnProgressCB
	onError 				as OnErrorCB
	
	lua						as lua_State ptr
	
	declare sub analisarInconsistenciasLRE()
	declare sub analisarInconsistenciasLRS()
end type
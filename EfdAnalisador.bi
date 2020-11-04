
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

type EfdAnalisador
public:
	declare constructor(tableExp as EfdTabelaExport ptr)
	declare function withDBs(db as SQLite ptr) as EfdAnalisador ptr
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdAnalisador ptr
	declare function withLua(lua as lua_State ptr) as EfdAnalisador ptr
	declare sub executar(safiFornecidoMask as long) 
	
private:
	db						as SQLite ptr
	tableExp				as EfdTabelaExport ptr
	onProgress 				as OnProgressCB
	onError 				as OnErrorCB
	
	lua						as lua_State ptr
	
	declare sub analisarInconsistenciasLRE()
	declare sub analisarInconsistenciasLRS()
end type
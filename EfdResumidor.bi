
enum TipoResumo
	TR_CFOP
	TR_CST
	TR_CST_CFOP
end enum

type EfdResumidor
public:
	declare constructor(tableExp as EfdTabelaExport ptr)
	declare function withDBs(db as TDb ptr) as EfdResumidor ptr
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdResumidor ptr
	declare function withLua(lua as lua_State ptr) as EfdResumidor ptr
	declare sub executar(safiFornecidoMask as long) 
	
private:
	db						as TDb ptr
	tableExp				as EfdTabelaExport ptr
	onProgress 				as OnProgressCB
	onError 				as OnErrorCB
	
	lua						as lua_State ptr

	declare sub criarResumosLRE()
	declare sub criarResumosLRS()
end type
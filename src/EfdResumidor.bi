
enum TipoResumo
	TR_CFOP
	TR_CST
	TR_CST_CFOP
end enum

type EfdResumidor
public:
	declare constructor(opcoes as OpcoesExtracao ptr, tableExp as EfdTabelaExport ptr)
	declare function withDBs(db as SQLite ptr) as EfdResumidor ptr
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdResumidor ptr
	declare function withLua(lua as lua_State ptr) as EfdResumidor ptr
	declare sub executar(safiFornecidoMask as long) 
	
private:
	opcoes as OpcoesExtracao ptr
	db as SQLite ptr
	tableExp as EfdTabelaExport ptr
	onProgress as OnProgressCB
	onError as OnErrorCB
	
	lua as lua_State ptr

	declare sub criarResumosLRE()
	declare sub criarResumosLRS()
	declare sub resumoAddHeaderCfopLRE(ws as TableTable ptr)
	declare sub resumoAddHeaderCstLRE(ws as TableTable ptr)
	declare sub resumoAddHeaderCstCfopLRE(ws as TableTable ptr)
	declare sub resumoAddHeaderCfopLRS(ws as TableTable ptr)
	declare sub resumoAddHeaderCstLRS(ws as TableTable ptr)
	declare sub resumoAddHeaderCstCfopLRS(ws as TableTable ptr)
end type
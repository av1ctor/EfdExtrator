#include once "EfdBoBaseLoader.bi"

type EfdBoXlsxLoader extends EfdBoBaseLoader
public:
	declare constructor(ctx as EfdBoLoaderContext ptr, opcoes as OpcoesExtracao ptr)
	declare function carregar(nomeArquivo as string) as boolean

private:
	declare function carregarXlsxNFeDest(reader as ExcelReader ptr) as TDFe_NFe ptr
	declare function carregarXlsxNFeDestItens(reader as ExcelReader ptr) as TDFe_NFe ptr
	declare function carregarXlsxNFeEmit(rd as ExcelReader ptr) as TDFe_NFe ptr
	declare function carregarXlsxNFeEmitItens(rd as ExcelReader ptr, chave as string, extra as TDFe ptr) as TDFe_NFeItem ptr
	declare function carregarXlsxCTe(rd as ExcelReader ptr, op as TipoOperacao) as TDFe_CTe ptr
	declare function carregarXlsxSAT(rd as ExcelReader ptr) as TDFe_NFe ptr
	declare function carregarXlsxSATItens(rd as ExcelReader ptr, chave as string) as TDFe_NFeItem ptr
end type
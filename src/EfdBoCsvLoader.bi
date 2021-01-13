#include once "EfdBoBaseLoader.bi"

type EfdBoCsvLoader extends EfdBoBaseLoader
public:
	declare constructor(ctx as EfdBoLoaderContext ptr, opcoes as OpcoesExtracao ptr)
	declare function carregar(nomeArquivo as string) as boolean

private:
	declare function carregarCsvNFeDestSAFI(bf as bfile, emModoOutrasUFs as boolean) as TDFe_Nfe ptr
	declare function carregarCsvNFeEmitSAFI(bf as bfile) as TDFe_Nfe ptr
	declare function carregarCsvNFeEmitItensSAFI(bf as bfile, chave as string) as TDFe_NFeItem ptr
	declare function carregarCsvCTeSAFI(bf as bfile, emModoOutrasUFs as boolean) as TDFe_CTe ptr
	declare function carregarCsvNFeEmitItens(bf as bfile, chave as string, extra as TDFe ptr) as TDFe_NFeItem ptr
end type
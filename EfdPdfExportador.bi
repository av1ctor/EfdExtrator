#include once "Efd.bi"

type EfdPdfExportador
public:
	declare constructor(baseTemplatesDir as string, infAssinatura as InfoAssinatura ptr, opcoes as OpcoesExtracao ptr)
	declare function withDBs(configDb as TDb ptr) as EfdPdfExportador ptr
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdPdfExportador ptr
	declare function withLua(lua as lua_State ptr, customLuaCbDict as TDict ptr) as EfdPdfExportador ptr
	declare function withDicionarios(participanteDict as TDict ptr, itemIdDict as TDict ptr, chaveDFeDict as TDict ptr, infoComplDict as TDict ptr, _
		obsLancamentoDict as TDict ptr, bemCiapDict as TDict ptr, contaContabDict as TDict ptr, centroCustoDict as TDict ptr, municipDict as TDict ptr) as EfdPdfExportador ptr
	declare function withFiltros(filtrarPorCnpj as OnFilterByStrCB, filtrarPorChave as OnFilterByStrCB) as EfdPdfExportador ptr
	declare destructor()
	declare sub gerar(regListHead as TRegistro ptr, regMestre as TRegistro ptr, nroRegs as integer)

private:
	infAssinatura 			as InfoAssinatura ptr
	opcoes					as OpcoesExtracao ptr
	baseTemplatesDir		as string
	
	configDb				as TDb ptr

	regMestre 				as TRegistro ptr	
	
	participanteDict		as TDict ptr
	itemIdDict          	as TDict ptr
	infoComplDict			as TDict ptr
	obsLancamentoDict		as TDict ptr
	bemCiapDict          	as TDict ptr
	chaveDFeDict          	as TDict ptr
	municipDict				as TDict ptr
	contaContabDict			as TDict ptr
	centroCustoDict			as TDict ptr

	onProgress 				as OnProgressCB
	onError 				as OnErrorCB
	filtrarPorCnpj 			as OnFilterByStrCB
	filtrarPorChave			as OnFilterByStrCB
	
	lua						as lua_State ptr
	customLuaCbDict			as TDict ptr		'' de CustomLuaCb
	
	ultimoRelatorio			as TipoRelatorio
	ultimoRelatorioSufixo	as string
	relSomaAnalDict			as TDict ptr
	relSomaAnalList			as TList ptr		'' de RelSomatorioAnal
	relSomaAjustesDict		as TDict ptr
	relSomaAjustesList		as TList ptr		'' de RelSomatorioAjuste
	nroRegistrosRel			as integer
	relLinhasList			as TList ptr		'' de RelLinha
	relNroLinhas			as double
	relYPos					as double
	relNroPaginas			as integer
	relNroTotalPaginas		as integer
	relTemplate				as PdfTemplate ptr
	relPage					as PdfPageElement ptr
	relOutFile 				as PdfDoc ptr

	declare sub gerarRelatorioCiap(reg as TRegistro ptr, isPre as boolean)
	declare sub gerarRelatorioApuracaoICMS(reg as TRegistro ptr, isPre as boolean)
	declare sub gerarAjusteTotalRelatorioApuracaoICMS(tipo as integer, total as double, isPre as boolean)
	declare sub gerarAjusteSubTotalRelatorioApuracaoICMS(tipo as integer, codigo as string, subtotal as double, isPre as boolean)
	declare sub gerarRelatorioApuracaoICMSST(reg as TRegistro ptr, isPre as boolean)
	declare sub iniciarRelatorio(relatorio as TipoRelatorio, nomeRelatorio as string, sufixo as string, isPre as boolean)
	declare sub adicionarDocRelatorioEntradas(doc as TDocDF ptr, part as TParticipante ptr, highlight as boolean, lg as boolean)
	declare sub adicionarDocRelatorioSaidas(doc as TDocDF ptr, part as TParticipante ptr, highlight as boolean, lg as boolean)
	declare sub adicionarDocRelatorioSaidas(doc as TECFReducaoZ ptr, highlight as boolean)
	declare sub adicionarDocRelatorioSaidas(doc as TDocSAT ptr, highlight as boolean)
	declare sub adicionarDocRelatorioItemAnal(sit as TipoSituacao, anal as TDocItemAnal ptr)
	declare sub adicionarDocRelatorioObs(sit as TipoSituacao, obs as TDocObs ptr, isFirst as boolean)
	declare sub adicionarDocRelatorioObsAjuste(sit as TipoSituacao, ajuste as TDocObsAjuste ptr, isFirst as boolean)
	declare sub finalizarRelatorio(isPre as boolean)
	declare sub relatorioSomarAnal(sit as TipoSituacao, anal as TDocItemAnal ptr, isPre as boolean)
	declare sub relatorioSomarAjuste(sit as TipoSituacao, ajuste as TDocObsAjuste ptr)
	declare function gerarPaginaRelatorio(lastPage as boolean, isPre as boolean) as boolean
	declare sub gerarResumoRelatorio(emitir as boolean, isPre as boolean)
	declare sub gerarResumoRelatorioHeader(emitir as boolean, isPre as boolean)
	declare sub gerarResumoAjustesRelatorioHeader(emitir as boolean, isPre as boolean)
	declare sub setNodeText(page as PdfPageElement ptr, id as zstring ptr, value as string, convert as boolean = false)
	declare sub setNodeText(page as PdfPageElement ptr, id as zstring ptr, value as wstring ptr)
	declare sub setChildText(row as PdfElement ptr, id as zstring ptr, value as string, convert as boolean = false)
	declare sub setChildText(row as PdfElement ptr, id as zstring ptr, value as wstring ptr)
	declare function gerarLinhaDFe(lg as boolean, highlight as boolean) as PdfElement ptr
	declare function gerarLinhaAnal() as PdfElement ptr
	declare function gerarLinhaObs(isFirst as boolean, parts as integer) as PdfElement ptr
	declare function gerarLinhaObsAjuste(isFirst as boolean) as PdfElement ptr
	declare sub criarPaginaRelatorio(emitir as boolean, isPre as boolean)
	declare sub emitirPaginaRelatorio(emitir as boolean, isPre as boolean)
	declare function calcObsHeight(sit as TipoSituacao, obs as TDocObs ptr, isFirst as boolean) as double
	
	declare static function luacb_efd_rel_addItemAnalitico cdecl(L as lua_State ptr) as long
end type

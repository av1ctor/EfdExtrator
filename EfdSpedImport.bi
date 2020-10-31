#include once "EfdBaseImport.bi"

type EfdSpedImport extends EfdBaseImport
public:
	declare constructor(opcoes as OpcoesExtracao ptr)
	declare function withStmts(lreInsertStmt as TDbStmt ptr, itensNfLRInsertStmt as TDbStmt ptr, lrsInsertStmt as TDbStmt ptr, _
		analInsertStmt as TDbStmt ptr, ressarcStItensNfLRSInsertStmt as TDbStmt ptr, itensIdInsertStmt as TDbStmt ptr, mestreInsertStmt as TDbStmt ptr) as EfdSpedImport ptr
	declare destructor()
	declare function carregar(nomeArquivo as string) as boolean
	declare function lerInfoAssinatura(nomeArquivo as string) as InfoAssinatura ptr

private:
	ultimoReg   			as TRegistro ptr
	ultimoDocNFItem  		as TDocNFItem ptr
	ultimoEquipECF			as TEquipECF ptr
	ultimoECFRedZ			as TECFReducaoZ ptr
	ultimoDocObs			as TDocObs ptr
	ultimoInventario		as TInventarioTotais ptr
	ultimoBemCiap			as TBemCiap ptr
	ultimoCiap				as TCiapTotal ptr
	ultimoCiapItem			as TCiapItem ptr
	ultimoCiapItemDoc		as TCiapItemDoc ptr
	ultimoEstoque			as TEstoquePeriodo ptr
	assinaturaP7K_DER(any) 	as byte

	db_LREInsertStmt		as TDbStmt ptr
	db_itensNfLRInsertStmt	as TDbStmt ptr
	db_LRSInsertStmt		as TDbStmt ptr
	db_analInsertStmt		as TDbStmt ptr
	db_ressarcStItensNfLRSInsertStmt as TDbStmt ptr
	db_itensIdInsertStmt 	as TDbStmt ptr
	db_mestreInsertStmt 	as TDbStmt ptr

	declare function lerRegistro(bf as bfile) as TRegistro ptr
	declare function lerTipo(bf as bfile, tipo as zstring ptr) as TipoRegistro
	declare function lerRegMestre(bf as bfile) as TMestre ptr
	declare function lerRegParticipante(bf as bfile) as TParticipante ptr
	declare function lerRegDocNF(bf as bfile) as TDocNF ptr
	declare function lerRegDocNFInfo(bf as bfile, pai as TDocNF ptr) as TDocInfoCompl ptr
	declare function lerRegDocNFItem(bf as bfile, documentoPai as TDocNF ptr) as TDocNFItem ptr
	declare function lerRegDocNFItemAnal(bf as bfile, documentoPai as TDocNF ptr) as TDocItemAnal ptr
	declare function lerRegDocNFItemRessarcSt(bf as bfile, documentoPai as TDocNFItem ptr) as TDocNFItemRessarcSt ptr
	declare function lerRegDocNFDifal(bf as bfile, documentoPai as TDocNF ptr) as TDocNF ptr
	declare function lerRegDocCT(bf as bfile) as TDocCT ptr
	declare function lerRegDocCTItemAnal(bf as bfile, docPai as TDocCT ptr) as TDocItemAnal ptr
	declare function lerRegDocCTDifal(bf as bfile, docPai as TDocCT ptr) as TDocCT ptr
	declare function lerRegEquipECF(bf as bfile) as TEquipECF ptr
	declare function lerRegDocECF(bf as bfile, equipECF as TEquipECF ptr) as TDocECF ptr
	declare function lerRegECFReducaoZ(bf as bfile, equipECF as TEquipECF ptr) as TECFReducaoZ ptr
	declare function lerRegDocECFItem(bf as bfile, documentoPai as TDocECF ptr) as TDocECFItem ptr
	declare function lerRegDocECFItemAnal(bf as bfile, documentoPai as TRegistro ptr) as TDocItemAnal ptr
	declare function lerRegDocSAT(bf as bfile) as TDocSAT ptr
	declare function lerRegDocSATItemAnal(bf as bfile, documentoPai as TRegistro ptr) as TDocItemAnal ptr
	declare function lerRegDocNFSCT(bf as bfile) as TDocNF ptr
	declare function lerRegDocNFSCTItemAnal(bf as bfile, documentoPai as TRegistro ptr) as TDocItemAnal ptr
	declare function lerRegDocNFElet(bf as bfile) as TDocNF ptr
	declare function lerRegDocNFEletItemAnal(bf as bfile, documentoPai as TRegistro ptr) as TDocItemAnal ptr
	declare function lerRegDocObs(bf as bfile) as TDocObs ptr
	declare function lerRegDocObsAjuste(bf as bfile) as TDocObsAjuste ptr
	declare function lerRegItemId(bf as bfile) as TItemId ptr
	declare function lerRegBemCiap(bf as bfile) as TBemCiap ptr
	declare function lerRegBemCiapInfo(bf as bfile, reg as TBemCiap ptr) as TBemCiap ptr
	declare function lerRegContaContab(bf as bfile) as TContaContab ptr
	declare function lerRegCentroCusto(bf as bfile) as TCentroCusto ptr
	declare function lerRegInfoCompl(bf as bfile) as TInfoCompl ptr
	declare function lerRegObsLancamento(bf as bfile) as TObsLancamento ptr
	declare function lerRegApuIcmsPeriodo(bf as bfile) as TApuracaoIcmsPropPeriodo ptr
	declare function lerRegApuIcmsProprio(bf as bfile, reg as TApuracaoIcmsPropPeriodo ptr) as TApuracaoIcmsPropPeriodo ptr
	declare function lerRegApuIcmsAjuste(bf as bfile, pai as TApuracaoIcmsPeriodo ptr) as TApuracaoIcmsAjuste ptr
	declare function lerRegApuIcmsSTPeriodo(bf as bfile) as TApuracaoIcmsSTPeriodo ptr
	declare function lerRegApuIcmsST(bf as bfile, reg as TApuracaoIcmsSTPeriodo ptr) as TApuracaoIcmsSTPeriodo ptr
	declare function lerRegInventarioTotais(bf as bfile) as TInventarioTotais ptr
	declare function lerRegInventarioItem(bf as bfile, inventarioPai as TInventarioTotais ptr) as TInventarioItem ptr
	declare function lerRegCiapTotal(bf as bfile) as TCiapTotal ptr
	declare function lerRegCiapItem(bf as bfile, pai as TCiapTotal ptr) as TCiapItem ptr
	declare function lerRegCiapItemDoc(bf as bfile, pai as TCiapItem ptr) as TCiapItemDoc ptr
	declare function lerRegCiapItemDocItem(bf as bfile, pai as TCiapItemDoc ptr) as TCiapItemDocItem ptr
	declare function lerRegEstoquePeriodo(bf as bfile) as TEstoquePeriodo ptr
	declare function lerRegEstoqueItem(bf as bfile, pai as TEstoquePeriodo ptr) as TEstoqueItem ptr
	declare function lerRegEstoqueOrdemProd(bf as bfile, pai as TEstoquePeriodo ptr) as TEstoqueOrdemProd ptr
	declare sub lerAssinatura(bf as bfile)

	declare function adicionarDocEscriturado(doc as TDocDF ptr) as long
	declare function adicionarDocEscriturado(doc as TDocECF ptr) as long
	declare function adicionarDocEscriturado(doc as TDocSAT ptr) as long
	declare function adicionarItemNFEscriturado(item as TDocNFItem ptr) as long
	declare function adicionarAnalEscriturado(item as TDocItemAnal ptr) as long
	declare function adicionarRessarcStEscriturado(doc as TDocNFItemRessarcSt ptr) as long
	declare function adicionarItemEscriturado(item as TItemId ptr) as long
	declare function adicionarMestre(reg as TMestre ptr) as long
	declare function addRegistroAoDB(reg as TRegistro ptr) as long
end type
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
	ultimoECFRedZ			as TRegistro ptr
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

	declare function lerRegistro(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerTipo(bf as bfile, tipo as zstring ptr) as TipoRegistro
	declare function lerRegMestre(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegParticipante(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegDocNF(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegDocNFInfo(bf as bfile, reg as TRegistro ptr, pai as TDocNF ptr) as Boolean
	declare function lerRegDocNFItem(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNF ptr) as Boolean
	declare function lerRegDocNFItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean
	declare function lerRegDocNFItemRessarcSt(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNFItem ptr) as Boolean
	declare function lerRegDocNFDifal(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNF ptr) as Boolean
	declare function lerRegDocCT(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegDocCTItemAnal(bf as bfile, reg as TRegistro ptr, docPai as TRegistro ptr) as Boolean
	declare function lerRegDocCTDifal(bf as bfile, reg as TRegistro ptr, docPai as TDocCT ptr) as Boolean
	declare function lerRegEquipECF(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegDocECF(bf as bfile, reg as TRegistro ptr, equipECF as TEquipECF ptr) as Boolean
	declare function lerRegECFReducaoZ(bf as bfile, reg as TRegistro ptr, equipECF as TEquipECF ptr) as Boolean
	declare function lerRegDocECFItem(bf as bfile, reg as TRegistro ptr, documentoPai as TDocECF ptr) as Boolean
	declare function lerRegDocECFItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean
	declare function lerRegDocSAT(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegDocSATItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean
	declare function lerRegDocNFSCT(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegDocNFSCTItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean
	declare function lerRegDocNFElet(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegDocNFEletItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean
	declare function lerRegDocObs(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegDocObsAjuste(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegItemId(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegBemCiap(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegBemCiapInfo(bf as bfile, reg as TBemCiap ptr) as Boolean
	declare function lerRegContaContab(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegCentroCusto(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegInfoCompl(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegObsLancamento(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegApuIcmsPeriodo(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegApuIcmsProprio(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegApuIcmsAjuste(bf as bfile, reg as TRegistro ptr, pai as TApuracaoIcmsPeriodo ptr) as Boolean
	declare function lerRegApuIcmsSTPeriodo(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegApuIcmsST(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegInventarioTotais(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegInventarioItem(bf as bfile, reg as TRegistro ptr, inventarioPai as TInventarioTotais ptr) as Boolean
	declare function lerRegCiapTotal(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegCiapItem(bf as bfile, reg as TRegistro ptr, pai as TCiapTotal ptr) as Boolean
	declare function lerRegCiapItemDoc(bf as bfile, reg as TRegistro ptr, pai as TCiapItem ptr) as Boolean
	declare function lerRegCiapItemDocItem(bf as bfile, reg as TRegistro ptr, pai as TCiapItemDoc ptr) as Boolean
	declare function lerRegEstoquePeriodo(bf as bfile, reg as TRegistro ptr) as Boolean
	declare function lerRegEstoqueItem(bf as bfile, reg as TRegistro ptr, pai as TEstoquePeriodo ptr) as Boolean
	declare function lerRegEstoqueOrdemProd(bf as bfile, reg as TRegistro ptr, pai as TEstoquePeriodo ptr) as Boolean
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
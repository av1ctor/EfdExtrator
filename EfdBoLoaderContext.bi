type EfdBoLoaderContext
public:
	chaveDFeDict			as TDict ptr
	dfeListHead				as TDFe ptr = null
	dfeListTail				as TDFe ptr = null
	nroDfe					as integer = 0
	cteListHead				as TDFe_CTe ptr = null	'' usado para fazer patch no tipo de operação
	cteListTail				as TDFe_CTe ptr = null
	nfeDestSafiFornecido 	as boolean
	nfeEmitSafiFornecido 	as boolean
	itemNFeSafiFornecido 	as boolean
	cteSafiFornecido		as boolean
	
	declare constructor()
	declare destructor()
	declare sub descarregar()
	declare function getFirstDFe() as TDFe ptr
	declare function getNroDFes() as integer
	declare function getChaveDFeDict() as TDict ptr
end type


enum BO_TipoArquivo
	BO_NFe_Dest
	BO_NFe_Emit
	BO_NFe_Emit_Itens
	BO_CTe
	BO_SAT
	BO_SAT_Itens
	BO_NFCe_Itens
	SAFI_ECF
	BO_ECF_Itens
end enum

enum BO_Dfe_Fornecido
	MASK_BO_NFe_DEST_FORNECIDO = &b00000001
	MASK_BO_NFe_EMIT_FORNECIDO = &b00000010
	MASK_BO_ITEM_NFE_FORNECIDO = &b00000100
	MASK_BO_CTe_FORNECIDO 	 = &b00001000
end enum

enum TDFE_LOADER
	LOADER_UNKNOWN
	LOADER_NFE_DEST
	LOADER_NFE_DEST_ITENS
	LOADER_NFE_EMIT
	LOADER_NFE_EMIT_ITENS
	LOADER_CTE
	LOADER_NFCE
	LOADER_SAT
	LOADER_ECF
end enum

type TDFe
	modelo			as TipoModelo
	operacao		as TipoOperacao					'' entrada ou saída
	chave			as zstring * 44+1
	dataEmi			as zstring * 8+1
	serie			as integer
	numero			as integer
	cnpjEmit		as zstring * 14+1
	nomeEmit		as zstring * 100+1
	ufEmit			as byte
	cnpjDest		as zstring * 14+1
	nomeDest		as zstring * 100+1
	ufDest			as byte
	valorOperacao	as double
	loader			as TDFE_LOADER
	prox			as TDFe ptr
end type

type TDFe_NFeItem									'' Nota: só existe para NF-e emitidas, já que para as recebidas os itens constam na EFD
	serie			as integer
	numero			as integer
	modelo			as TipoModelo
	nroItem			as integer
	cfop			as short
	ncm				as longint
	cst				as integer
	codProduto		as zstring * 60+1
	descricao		as zstring * 256+1
	qtd				as double
	unidade			as zstring * 6+1
	valorProduto	as double
	desconto		as double
	despesasAcess	as double
	bcICMS			as double
	aliqICMS		as double
	ICMS			as double
	bcICMSST		as double
	aliqIcmsST		as double
	icmsST			as double
	IPI				as double
	next_			as TDFe_NFeItem ptr
end type

type TDFe_NFe extends TDFe
	ieEmit			as zstring * 14+1
	ieDest			as zstring * 14+1
	bcICMSTotal		as double
	ICMSTotal		as double
	bcICMSSTTotal	as double
	ICMSSTTotal		as double
	
	itemListHead	as TDFe_NFeItem ptr
	itemListTail	as TDFe_NFeItem ptr
end type

type TDFe_CTe extends TDFe
	cnpjToma		as zstring * 14+1
	nomeToma		as zstring * 100+1
	ufToma			as zstring * 2+1
	cnpjRem			as zstring * 14+1
	nomeRem			as zstring * 100+1
	ufRem			as zstring * 2+1
	cnpjExp			as zstring * 14+1
	ufExp			as zstring * 2+1
	cnpjReceb		as zstring * 14+1
	ufReceb			as zstring * 2+1
	tipo			as byte
	valorReceber	as double
	qtdCCe			as double
	cfop			as integer
	nomeMunicIni	as zstring * 64+1
	ufIni			as zstring * 2+1
	nomeMunicFim	as zstring * 64+1
	ufFim			as zstring * 2+1
	next_			as TDFe_CTe ptr					'' usado para dar patch 
	parent			as TDFe_CTe ptr
end type

#include once "EfdBoLoaderContext.bi"

type EfdBoBaseLoader extends object
public:
	declare constructor()
	declare constructor(ctx as EfdBoLoaderContext ptr, opcoes as OpcoesExtracao ptr)
	declare destructor()
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdBoBaseLoader ptr
	declare function withDBs(db as TDb ptr) as EfdBoBaseLoader ptr
	declare function withStmts(dfeEntradaInsertStmt as TDbStmt ptr, dfeSaidaInsertStmt as TDbStmt ptr, itensDfeSaidaInsertStmt as TDbStmt ptr) as EfdBoBaseLoader ptr
	declare abstract function carregar(nomeArquivo as string) as boolean

protected:
	ctx 					as EfdBoLoaderContext ptr

	opcoes					as OpcoesExtracao ptr
	db						as TDb ptr

	onProgress 				as OnProgressCB
	onError 				as OnErrorCB
	
	db_dfeEntradaInsertStmt	as TDbStmt ptr
	db_dfeSaidaInsertStmt	as TDbStmt ptr
	db_itensDfeSaidaInsertStmt as TDbStmt ptr
	
	declare function adicionarDFe(dfe as TDFe_NFe ptr, fazerInsert as boolean = true) as long
	declare function adicionarDFe(dfe as TDFe_CTe ptr, fazerInsert as boolean = true) as long
	declare function adicionarDFe(dfe as TDFe ptr, isNfe as boolean, fazerInsert as boolean = true) as long
	declare function adicionarItemDFe(chave as const zstring ptr, item as TDFe_NFeItem ptr) as long
end type
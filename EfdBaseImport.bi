#include once "Efd.bi"
#include once "Dict.bi"
#include once "bfile.bi"
#include once "Db.bi"

type EfdBaseImport extends object
public:
	declare constructor()
	declare constructor(opcoes as OpcoesExtracao ptr)
	declare destructor()
	declare function withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdBaseImport ptr
	declare function withLua(lua as lua_State ptr, customLuaCbDict as TDict ptr) as EfdBaseImport ptr
	declare function withDBs(db as TDb ptr) as EfdBaseImport ptr
	declare abstract function carregar(nomeArquivo as string) as boolean
	declare function getFirstReg() as TRegistro ptr
	declare function getMestreReg() as TMestre ptr
	declare function getNroRegs() as integer
	declare function getParticipanteDict() as TDict ptr
	declare function getItemIdDict as TDict ptr
	declare function getInfoComplDict as TDict ptr
	declare function getObsLancamentoDict as TDict ptr
	declare function getBemCiapDict as TDict ptr
	declare function getContaContabDict as TDict ptr
	declare function getCentroCustoDict as TDict ptr
	declare function getTipoArquivo() as TTipoArquivo

protected:
	opcoes					as OpcoesExtracao ptr
	db						as TDb ptr

	onProgress 				as OnProgressCB
	onError 				as OnErrorCB
	
	lua						as lua_State ptr
	customLuaCbDict			as TDict ptr		'' de CustomLuaCb
	
	tipoArquivo				as TTipoArquivo
	regListHead         	as TRegistro ptr = null
	regMestre				as TMestre ptr
	nroRegs             	as integer = 0
	nroLinha				as integer
	
	participanteDict		as TDict ptr
	itemIdDict          	as TDict ptr
	infoComplDict			as TDict ptr
	obsLancamentoDict		as TDict ptr
	bemCiapDict          	as TDict ptr
	contaContabDict			as TDict ptr
	centroCustoDict			as TDict ptr
end type


#include once "EfdBaseImport.bi"

''''''''
constructor EfdBaseImport(opcoes as OpcoesExtracao ptr)
	this.opcoes = opcoes
	
	participanteDict = new TDict(2^16)
	itemIdDict = new TDict(2^10)
	bemCiapDict = new TDict(2^10)
	infoComplDict = new TDict(2^10)
	obsLancamentoDict = new TDict(2^10)
	contaContabDict = new TDict(2^10)
	centroCustoDict = new TDict(2^10)
end constructor

''''''''
function EfdBaseImport.withCallbacks(onProgress as OnProgressCB, onError as OnErrorCB) as EfdBaseImport ptr
	this.onProgress = onProgress
	this.onError = onError
	return @this
end function

''''''''
function EfdBaseImport.withLua(lua as lua_State ptr, customLuaCbDict as TDict ptr) as EfdBaseImport ptr
	this.lua = lua
	this.customLuaCbDict = customLuaCbDict
	return @this
end function

''''''''
function EfdBaseImport.withDBs(db as TDb ptr) as EfdBaseImport ptr
	this.db = db
	return @this
end function

''''''''
destructor EfdBaseImport()
	do while regListHead <> null
		var prox = regListHead->prox
		delete regListHead
		regListHead = prox
	loop

	delete infoComplDict
	delete obsLancamentoDict
	delete centroCustoDict
	delete contaContabDict
	delete bemCiapDict
	delete itemIdDict
	delete participanteDict

end destructor

''''''''
function EfdBaseImport.getFirstReg() as TRegistro ptr
	return regListHead
end function

''''''''
function EfdBaseImport.getMestreReg() as TMestre ptr
	return regMestre
end function

''''''''
function EfdBaseImport.getNroRegs() as integer
	return nroRegs
end function

''''''''
function EfdBaseImport.getParticipanteDict() as TDict ptr
	return participanteDict
end function

''''''''
function EfdBaseImport.getItemIdDict as TDict ptr
	return itemIdDict
end function

''''''''
function EfdBaseImport.getInfoComplDict as TDict ptr
	return infoComplDict
end function

''''''''
function EfdBaseImport.getObsLancamentoDict as TDict ptr
	return obsLancamentoDict
end function

''''''''
function EfdBaseImport.getBemCiapDict as TDict ptr
	return bemCiapDict
end function

''''''''
function EfdBaseImport.getContaContabDict as TDict ptr
	return contaContabDict
end function

''''''''
function EfdBaseImport.getCentroCustoDict as TDict ptr
	return centroCustoDict
end function

''''''''
function EfdBaseImport.getTipoArquivo() as TTipoArquivo
	return tipoArquivo
end function
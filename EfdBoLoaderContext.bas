#include once "Efd.bi"
#include once "EfdBoLoaderContext.bi"

''''''''
constructor EfdBoLoaderContext()
	chaveDFeDict = new TDict(2^20)
	nfeDestSafiFornecido = false
	nfeEmitSafiFornecido = false
	itemNFeSafiFornecido = false
	cteSafiFornecido = false
	dfeListHead = null
	dfeListTail = null
end constructor

''''''''
destructor EfdBoLoaderContext()
	if chaveDFeDict <> null then
		delete chaveDFeDict
		chaveDFeDict = null
	end if
	
	descarregar()
end destructor

''''''''
sub EfdBoLoaderContext.descarregar()
	do while dfeListHead <> null
		var prox = dfeListHead->prox
		select case dfeListHead->modelo
		case NFE, SAT
			var head = cast(TDFe_NFe ptr, dfeListHead)->itemListHead
			do while head <> null
				var next_ = head->next_
				delete head
				head = next_
			loop
		end select
		delete dfeListHead
		dfeListHead = prox
	loop
end sub

''''''''
function EfdBoLoaderContext.getFirstDFe() as TDFe ptr
	return dfeListHead
end function

''''''''
function EfdBoLoaderContext.getNroDFes() as integer
	return nroDFe
end function

''''''''
function EfdBoLoaderContext.getChaveDFeDict() as TDict ptr
	return chaveDFeDict
end function


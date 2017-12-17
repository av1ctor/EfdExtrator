
#include once "efd.bi"
#include once "ExcelWriter.bi"
#include once "vbcompat.bi"
#include once "DB.bi"

''''''''
sub Efd.analisar(mostrarProgresso as ProgressoCB) 

	analisarFaltaDeEscrituracao(mostrarProgresso)

end sub

''''''''
private sub faltaDeEscrituracaoAddHeaderCols(ws as ExcelWorksheet ptr)
	ws->AddCellType(CT_STRING, "Chave")
	ws->AddCellType(CT_DATE, "Data")
	ws->AddCellType(CT_INTNUMBER, "Modelo")
	ws->AddCellType(CT_INTNUMBER, "Serie")
	ws->AddCellType(CT_INTNUMBER, "Numero")
	ws->AddCellType(CT_MONEY, "Valor Operacao")
end sub

''''''''
private sub faltaDeEscrituracaoAddCols(xrow as ExcelRow ptr, byref drow as TRSetRow)
	xrow->addCell(drow["chave"])
	xrow->addCell(yyyyMmDd2Datetime(drow["dataEmit"]))
	xrow->addCell(drow["modelo"])
	xrow->addCell(drow["serie"])
	xrow->addCell(drow["numero"])
	xrow->addCell(drow["valorOp"])
end sub

''''''''
sub Efd.analisarFaltaDeEscrituracao(mostrarProgresso as ProgressoCB)
	
	'' entradas
	entradasNaoEscrituradas = ew->AddWorksheet("Entradas nao escrituradas")
	faltaDeEscrituracaoAddHeaderCols(entradasNaoEscrituradas)
	
	mostrarProgresso(wstr(!"\tFalta de escrituração nas entradas"), 0)
	
	if not nfeDestSafiFornecido or not cteSafiFornecido then
		var row = entradasNaoEscrituradas->AddRow()
		row->addCell("Nao foi possivel verificar falta de escrituracao nas entradas porque os relatorios SAFI_NFe_Destinatario ou SAFI_CTe_CNPJ nao foram fornecidos")
	else
		var rs = db->exec( _
			"select " + _
					"d.chave, d.dataEmit, d.modelo, d.serie, d.numero, d.valorOp " + _
				"from dfeEntrada d " + _
				"left join LRE l " + _
					"on l.cnpjEmit = d.cnpjEmit and l.ufEmit = d.ufEmit and l.serie = d.serie and l.numero = d.numero " + _
				"where l.cnpjEmit is null " + _
				"order by d.dataEmit asc" _
		)
		
		do while rs->hasNext()
			faltaDeEscrituracaoAddCols( entradasNaoEscrituradas->AddRow(), *rs->row )
			rs->next_()
		loop
	end if
	
	mostrarProgresso(null, 1)
	
	'' saÃ­das
	saidasNaoEscrituradas = ew->AddWorksheet("Saidas nao escrituradas")
	faltaDeEscrituracaoAddHeaderCols(saidasNaoEscrituradas)
	
	mostrarProgresso(wstr(!"\tFalta de escrituração nas saídas"), 0)
	
	if not nfeEmitSafiFornecido or not cteSafiFornecido then
		var row = saidasNaoEscrituradas->AddRow()
		row->addCell("Nao foi possivel verificar falta de escrituracao nas saidas porque os relatorios SAFI_NFe_Emitente ou SAFI_CTe_CNPJ nao foram fornecidos")
	else
		var rs = db->exec( _
			"select " + _
					"d.chave, d.dataEmit, d.modelo, d.serie, d.numero, d.valorOp " + _
				"from dfeSaida d " + _
				"left join LRS l " + _
					"on l.cnpjDest = d.cnpjDest and l.ufDest = d.ufDest and l.serie = d.serie and l.numero = d.numero " + _
				"where l.cnpjDest is null " + _
				"order by d.dataEmit asc" _
		)
		
		do while rs->hasNext()
			faltaDeEscrituracaoAddCols( saidasNaoEscrituradas->AddRow(), *rs->row )
			rs->next_()
		loop
	end if
	
	mostrarProgresso(null, 1)	
end sub


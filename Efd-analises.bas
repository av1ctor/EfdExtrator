
#include once "efd.bi"
#include once "ExcelWriter.bi"
#include once "vbcompat.bi"
#include once "DB.bi"

''''''''
sub Efd.analisar(mostrarProgresso as ProgressoCB) 

	if not (nfeDestSafiFornecido or nfeEmitSafiFornecido or itemNFeSafiFornecido or cteSafiFornecido) then
		print wstr(!"\tNão será possivel realizar análises e cruzamentos porque os relatórios Infoview BO do SAFI não foram fornecidos")
	else
		analisarInconsistenciasLRE(mostrarProgresso)
		analisarInconsistenciasLRS(mostrarProgresso)
	end if
	
end sub

''''''''
private sub inconsistenciaAddHeader(ws as ExcelWorksheet ptr)
	ws->AddCellType(CT_STRING, "Chave")
	ws->AddCellType(CT_DATE, "Data")
	ws->AddCellType(CT_INTNUMBER, "Modelo")
	ws->AddCellType(CT_INTNUMBER, "Serie")
	ws->AddCellType(CT_INTNUMBER, "Numero")
	ws->AddCellType(CT_MONEY, "Valor Operacao")
	ws->AddCellType(CT_INTNUMBER, "Tipo Inconsistencia")
	ws->AddCellType(CT_STRING, "Descricao Inconsistencia")
end sub

''''''''
private sub inconsistenciaAddRow(xrow as ExcelRow ptr, byref drow as TDataSetRow, tIncons as TipoInconsistencia, descricao as const zstring ptr)
	xrow->addCell(drow["chave"])
	xrow->addCell(yyyyMmDd2Datetime(drow["dataEmit"]))
	xrow->addCell(drow["modelo"])
	xrow->addCell(drow["serie"])
	xrow->addCell(drow["numero"])
	xrow->addCell(drow["valorOp"])
	xrow->addCell(tIncons)
	xrow->addCell(*descricao)
end sub

''''''''
sub Efd.analisarInconsistenciasLRE(mostrarProgresso as ProgressoCB)

	var ws = ew->AddWorksheet("Inconsistencias LRE")
	inconsistenciaAddHeader(ws)
	
	mostrarProgresso(wstr(!"\tInconsistências nas entradas"), 0)

	'' docs escriturados, mas não encontrados no BO
	scope
		var ds = db->exec( _
			"select " + _
					"l.chave, l.dataEmit, l.modelo, l.serie, l.numero, l.valorOp " + _
				"from LRE l " + _
				"left join dfeEntrada d " + _
					"on l.cnpjEmit = d.cnpjEmit and l.ufEmit = d.ufEmit and l.serie = d.serie and l.numero = d.numero and l.modelo = d.modelo " + _
				"where d.cnpjEmit is null and l.modelo >= 55 " + _
				"order by l.dataEmit asc" _
		)
		
		do while ds->hasNext()
			inconsistenciaAddRow( ws->AddRow(), *ds->row, TI_ESCRIT_FANTASMA, "DF-e nao encontrado no BO" )
			ds->next_()
		loop
	end scope

	'' docs escriturados em valores superiores
	scope
		var ds = db->exec( _
			"select " + _
					"l.chave, l.dataEmit, l.modelo, l.serie, l.numero, l.valorOp valorOp, d.valorOp d_valorOp " + _
				"from LRE l " + _
				"inner join dfeEntrada d " + _
					"on l.cnpjEmit = d.cnpjEmit and l.ufEmit = d.ufEmit and l.serie = d.serie and l.numero = d.numero and l.modelo = d.modelo " + _
				"where d.valorOp < l.valorOp " + _
				"order by l.dataEmit asc" _
		)
		
		do while ds->hasNext()
			var row = ds->row
			var dif = val(*(*row)["valorOp"]) - val(*(*row)["d_valorOp"])
			inconsistenciaAddRow( ws->AddRow(), *row, TI_DIF, "Diferença de valores: R$ " & DBL2MONEYBR(dif) )
			ds->next_()
		loop
	end scope

	'' docs escriturados em duplicidade
	scope
		var ds = db->exec( _
			"select " + _
					"l.chave, l.dataEmit, l.modelo, l.serie, l.numero, l.valorOp, l.periodo periodo1, l2.periodo periodo2 " + _
				"from LRE l " + _
				"inner join LRE l2 " + _
					"on l2.cnpjEmit = l.cnpjEmit and l2.ufEmit = l.ufEmit and l2.serie = l.serie and l2.numero = l.numero and l2.modelo = l.modelo " + _
				"where l.periodo != l2.periodo " + _
				"order by l.dataEmit asc" _
		)
		
		do while ds->hasNext()
			var row = ds->row
			var periodo1 = *(*row)["periodo1"]
			var periodo2 = *(*row)["periodo2"]
			inconsistenciaAddRow( ws->AddRow(), *row, TI_DUP, "Duplicado em: " & DdMmYyyy2Yyyy_Mm(periodo2) )
			ds->next_()
		loop
	end scope

	'' docs escriturados de operações interestaduais com alíquota > 12%
	scope
		var ds = db->exec( _
			"select " + _
					"l.chave, l.dataEmit, l.modelo, l.serie, l.numero, l.valorOp, it.aliq " + _
				"from LRE l " + _
				"inner join itensNfLRE it " + _
					"on it.cnpjEmit = l.cnpjEmit and it.ufEmit = l.ufEmit and it.serie = l.serie and it.numero = l.numero and it.periodo = l.periodo and it.modelo = l.modelo " + _
				"where it.aliq > 12 and l.ufEmit != 35 " + _
				"order by l.dataEmit asc" _
		)
		
		do while ds->hasNext()
			var row = ds->row
			var aliq = *(*row)["aliq"]
			inconsistenciaAddRow( ws->AddRow(), *row, TI_ALIQ, "Interestadual com aliquota superior a 12%: " & aliq & "%" )
			ds->next_()
		loop
	end scope

	'' docs escriturados com crédito de fornecedor SN acima do permitido
	scope
		const aliqMaxSN = str(9)
		var ds = db->exec( _
			"select " + _
					"l.chave, l.dataEmit, l.modelo, l.serie, l.numero, l.valorOp, it.aliq " + _
				"from LRE l " + _
				"inner join itensNfLRE it " + _
					"on it.cnpjEmit = l.cnpjEmit and it.ufEmit = l.ufEmit and it.serie = l.serie and it.numero = l.numero and it.periodo = l.periodo and it.modelo = l.modelo " + _
				"inner join cdb.Contribuinte c " + _
					"on c.cnpj = l.cnpjEmit and l.ufEmit = 35 and c.dataIni <= l.dataEmit and c.dataFim > l.dataEmit " + _
				"inner join cdb.Regimes r " + _
					"on r.ie = c.ie and r.tipo = 'N' and r.dataIni <= cast(substr(l.dataEmit,1,6) as integer) and r.dataFim > cast(substr(l.dataEmit,1,6) as integer) " + _
				"where it.aliq > " + aliqMaxSN + " " + _
				"order by l.dataEmit asc" _
		)
		
		do while ds->hasNext()
			var row = ds->row
			var aliq = *(*row)["aliq"]
			inconsistenciaAddRow( ws->AddRow(), *row, TI_ALIQ, "Credito de SN acima do permitido: " & aliq & "%" )
			ds->next_()
		loop
	end scope
	
	'' entradas não escrituradas
	scope
		var ds = db->exec( _
			"select " + _
					"d.chave, d.dataEmit, d.modelo, d.serie, d.numero, d.valorOp " + _
				"from dfeEntrada d " + _
				"left join LRE l " + _
					"on l.cnpjEmit = d.cnpjEmit and l.ufEmit = d.ufEmit and l.serie = d.serie and l.numero = d.numero and l.modelo = d.modelo " + _
				"where l.cnpjEmit is null " + _
				"order by d.dataEmit asc" _
		)
		
		do while ds->hasNext()
			inconsistenciaAddRow( ws->AddRow(), *ds->row, TI_ESCRIT_FALTA, "DF-e nao escriturado" )
			ds->next_()
		loop
	end scope
	
	mostrarProgresso(null, 1)

end sub

''''''''
sub Efd.analisarInconsistenciasLRS(mostrarProgresso as ProgressoCB)
	
	var ws = ew->AddWorksheet("Inconsistencias LRS")
	inconsistenciaAddHeader(ws)
	
	mostrarProgresso(wstr(!"\tInconsistências nas saídas"), 0)

	'' saídas não escrituradas
	scope
		var ds = db->exec( _
			"select " + _
					"d.chave, d.dataEmit, d.modelo, d.serie, d.numero, d.valorOp " + _
				"from dfeSaida d " + _
				"left join LRS l " + _
					"on l.cnpjDest = d.cnpjDest and l.ufDest = d.ufDest and l.serie = d.serie and l.numero = d.numero and l.modelo = d.modelo " + _
				"where l.cnpjDest is null " + _
				"order by d.dataEmit asc" _
		)
		
		do while ds->hasNext()
			inconsistenciaAddRow( ws->AddRow(), *ds->row, TI_ESCRIT_FALTA, "DF-e nao escriturado" )
			ds->next_()
		loop
	end scope
	
	mostrarProgresso(null, 1)
end sub



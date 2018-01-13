-- API do Extrator com as funções que podem ser chamadas dos scripts Lua:

-- DB:
	db = db_new()							-- construtor
	db_del(db)								-- destrutor
	bool = db_open(db, filename)			-- abre um DB SQLite
	db_close(db)							-- fecha o db aberto
	ds = db_exec(db, query)					-- executa a query e retorna um DataSet (nil se erro)
	ds = db_exec(db, stmt)					-- executa o statement e retorna um DataSet (nil se erro); stmt retornado pelo db_prepare
	db_execNonQuery(db, query)				-- executa a query
	db_execNonQuery(db, stmt)				-- executa o statemnt; stmt retornado pelo db_prepare
	stmt = db_prepare(db, query)			-- compila a query e retorna o statement (nil se erro)
	
-- DB DataSet:
	ds_del(ds)								-- destrutor
	bool = ds_hasNext(ds)					-- retorna se ainda há linhas para processar
	ds_next(ds)								-- vai para a próxima linha
	str = ds_row_getColValue(ds, colname)	-- retorna o valor (string) da coluna 'colname' na linha atual do dataset
	row = ds_row(ds)						-- retorna uma array com as colunas da linha atual do dataset
	
	
-- ExcelWriter:
	ew = ew_new()							-- construtor
	ew_del(ew)								-- destrutor
	bool = ew_create(ew, filename)			-- criar um arquivo (.xml para Excel 2003+)
	ew_close(ew)							-- fecha o arquivo
	ws = ew_addWorksheet(ew, name)			-- cria uma planilha (aba)
	
-- ExcelWriter Worksheet:
	er = ws_addRow(ws)						-- adiciona uma linha à planilha
	ws_addCellType(typ, name)				-- adiciona um CellType (header de uma coluna), com tipo (CT_STRING,CT_NUMBER,CT_INTNUMBER,CT_DATE,CT_MONEY) e nome
	
-- ExcelWriter Row:
	ec = er_addCell(er, contents)			-- adiciona uma célula à linha da planilha; 'contents' pode ser string ou número
	
	
-- efd	
	ws = efd_plan_get(nome)					-- retorna uma planilha interna, pesquisando pelo nome (entradas, saidas, inconsistenciasLRE, inconsistenciasLRS)
	efd_plan_inconsistencias_AddRow(ws, ds, tipoInconsistencia, descricao)	-- tipo in (TI_ESCRIT_FALTA,TI_ESCRIT_FANTASMA,TI_ALIQ,TI_DUP,TI_DIF)
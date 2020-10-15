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
	res = db_execScalarInt(db, query)		-- executa a query e retorna um inteiro (nil se erro)
	res = db_execScalarDbl(db, query)		-- executa a query e retorna um double (nil se erro)
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
	bool = ew_create(ew, filename)			-- criar um arquivo (Excel 2003+)
	ew_close(ew)							-- fecha o arquivo
	ws = ew_addWorksheet(ew, name)			-- cria uma planilha (aba)
	
-- ExcelWriter Worksheet:
	er = ws_addRow(ws, num, asIs)			-- adiciona uma linha à planilha; num = número da linha (0 = primeira), asIs = 1 se é para gravar sem converter valores
	ws_addCellType(typ)						-- adiciona um CellType (tipo da coluna), com tipo (CT_STRING,CT_STRING_UTF8,CT_NUMBER,CT_INTNUMBER,CT_DATE,CT_MONEY,CT_PERCENT)
	
-- ExcelWriter Row:
	ec = er_addCell(er, contents)			-- adiciona uma célula à linha da planilha; 'contents' pode ser string ou número

-- efd	
	ws = efd_plan_get(nome)					-- retorna uma planilha interna, pesquisando pelo nome (entradas, saidas, inconsistencias LRE, inconsistencias LRS, resumos LRE, resumos LRS)
	efd_plan_entradas						-- planilha de entradas (variável global)
	efd_plan_saidas							-- planilha de saidas (variável global)
	efd_plan_inconsistencias_AddRow(ws, ds, tipoInconsistencia, descricao)	-- tipo in (TI_ESCRIT_FALTA,TI_ESCRIT_FANTASMA,TI_ALIQ,TI_DUP,TI_DIF)
	efd_plan_resumos_AddRow(ws, ds, livro)	-- livro in (TL_ENTRADAS, TL_SAIDAS)
	part = efd_participante_get(id, formatar) -- retorna o objeto participante { cnpj, ie, nome, uf, municip }; formatar = true formatará os campos cnpj, ie etc
	
-- bfile
	char = bf_char1(bf)						-- ler um char (1 byte)
	int = bf_int1(bf)						-- ler um char e converter para inteiro
	int = bf_varint(bf[, separador])		-- ler um inteiro até encontrar o separador; separador padrão = asc("|")
	dbl = bf_vardbl(bf[, separador])		-- ler um double até encontrar o separador; separador padrão = asc("|")
	str = bf_varchar(bf[, separador])		-- ler uma string até encontrar o separador; separador padrão = asc("|")
	
-- misc	
	onError(msg)
	onProgress(msg, total)

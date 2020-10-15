
----------------------------------------------------------------------
function configurarDB(db, dbPath)
	db_execNonQuery(db, "attach '" .. dbPath .. "config.db' as conf")
	db_execNonQuery(db, "attach '" .. dbPath .. "CadContribuinte.db' as cdb")
	db_execNonQuery(db, "attach '" .. dbPath .. "inidoneos.db' as idb")
	db_execNonQuery(db, "attach '" .. dbPath .. "GIA.db' as gdb")
end

----------------------------------------------------------------------
-- criar tabela de dfe's de entrada (relatórios do SAFI ou do BO)
function criarTabela_DFe_Entradas(db)

	db_execNonQuery( db, [[
		create table DFe_Entradas(
			chave		char(44) not null,
			cnpjEmit	bigint not null,
			ufEmit		short not null,
			serie		short not null,
			numero		integer not null,
			modelo		short not null,
			dataEmit	integer not null,
			valorOp		real not null,
			ieEmit		varchar(20) null,
			PRIMARY KEY (
				chave
			)
		)
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX DFe_Entradas_chaveIdx ON DFe_Entradas (
			cnpjEmit,
			ufEmit,
			serie,
			numero,
			modelo
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX DFe_Entradas_cnpjEmitIdx ON DFe_Entradas (
			cnpjEmit,
			ufEmit
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX DFe_Entradas_ieEmitIdx ON DFe_Entradas (
			ieEmit
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into DFe_Entradas 
			(cnpjEmit, ufEmit, serie, numero, modelo, chave, dataEmit, valorOp, ieEmit)
			values (?,?,?,?,?,?,?,?,?)
	]]

end

----------------------------------------------------------------------
-- criar tabela de dfe's de saída (relatórios do SAFI ou do BO)
function criarTabela_DFe_Saidas(db)

	db_execNonQuery( db, [[
		create table DFe_Saidas( 
			serie		short not null,
			numero		integer not null,
			modelo		short not null,
			chave		char(44) not null,
			dataEmit	integer not null,
			valorOp		real not null,
			cnpjDest	bigint not null,
			ufDest		short not null,
			ieDest		varchar(20) null,
			PRIMARY KEY (
				serie,
				numero,
				modelo
			)
		) 
	]])
	
	db_execNonQuery( db, [[ 
		CREATE INDEX DFe_Saidas_chaveIdx ON DFe_Saidas (
			serie,
			numero,
			modelo,
			cnpjDest,
			ufDest
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX DFe_Saidas_cnpjDestIdx ON DFe_Saidas (
			cnpjDest,
			ufDest
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX DFe_Saidas_ieDestIdx ON DFe_Saidas (
			ieDest
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into DFe_Saidas 
			(cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp, ieDest) 
			values (?,?,?,?,?,?,?,?,?)
	]]

end

----------------------------------------------------------------------
-- criar tabela de itens de docs saída (relatórios do SAFI ou do BO)
function criarTabela_DFe_Saidas_Itens(db)

	db_execNonQuery( db, [[
		create table DFe_Saidas_Itens( 
			serie		short not null,
			numero		integer not null,
			modelo		short not null,
			numItem		short not null,
			chave		char(44) not null,
			cfop		integer not null,
			valorProd	real not null,
			valorDesc	real not null,
			valorAcess	real not null,
			bc			real not null,
			aliq		real not null,
			icms		real not null,
			bcIcmsST	real not null,
			aliqST		real not null,
			icmsST		real not null,
			ncm			bigint null,
			cst			integer null,
			qtd			real null,
			unidade		varchar(8) null,
			codProduto	varchar(64) null,
			descricao	varchar(256) null,
			PRIMARY KEY (
				serie,
				numero,
				modelo,
				numItem
			)
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX DFe_Saidas_ItensChaveIdx ON DFe_Saidas_Itens (
			chave
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX DFe_Saidas_ItensCfopIdx ON DFe_Saidas_Itens (
			cfop
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX DFe_Saidas_ItensAliqIdx ON DFe_Saidas_Itens (
			aliq
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX DFe_Saidas_ItensNcmIdx ON DFe_Saidas_Itens (
			ncm
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX DFe_Saidas_ItensCstIdx ON DFe_Saidas_Itens (
			cst
		) 
	]])

	-- retornar a query que será usada no insert
	return [[
		insert into DFe_Saidas_Itens 
			(serie, numero, modelo, numItem, chave, cfop, valorProd, valorDesc, valorAcess, bc, aliq, icms, bcIcmsST, aliqST, icmsST, ncm, cst, qtd, unidade, codProduto, descricao) 
			values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
	]]

end

----------------------------------------------------------------------
-- criar tabela LRE
function criarTabela_EFD_LRE(db)

	db_execNonQuery( db, [[
		create table EFD_LRE( 
			periodo		integer not null,
			cnpjEmit	bigint not null,
			ufEmit		short not null,
			serie		short not null,
			numero		integer not null,
			modelo		short not null,
			dataEmit	integer not null,
			valorOp		real not null,
			chave		char(44) null,
			ieEmit		varchar(20) null,
			PRIMARY KEY (
				periodo,
				cnpjEmit,
				ufEmit,
				serie,
				numero,
				modelo
			)
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRE_cnpjUfSerieNumeroIdx ON EFD_LRE (
			cnpjEmit,
			ufEmit,
			serie,
			numero,
			modelo
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRE_chaveIdx ON EFD_LRE (
			chave
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRE_cnpjEmitIdx ON EFD_LRE (
			cnpjEmit,
			ufEmit
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRE_ufEmitIdx ON EFD_LRE (
			ufEmit
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRE_ieEmitIdx ON EFD_LRE (
			ieEmit
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into EFD_LRE 
			(periodo, cnpjEmit, ufEmit, serie, numero, modelo, chave, dataEmit, valorOp, ieEmit) 
			values (?,?,?,?,?,?,?,?,?,?)
	]]
	
end

----------------------------------------------------------------------
-- criar tabela itens de DFe da LRE (ou LRS no caso de ressarcimento ST)
function criarTabela_EFD_Itens(db)

	db_execNonQuery( db, [[
		create table EFD_Itens( 
			periodo		integer not null,
			cnpjEmit	bigint not null,
			ufEmit		short not null,
			serie		short not null,
			numero		integer not null,
			modelo		short not null,
			numItem		short not null,
			cst			short not null,
			cst_origem	short not null,
			cst_tribut	short not null,
			cfop		short not null,
			qtd			real not null,
			valorProd	real not null,
			valorDesc	real not null,
			bc			real not null,
			aliq		real not null,
			icms		real not null,
			bcIcmsST	real not null,
			aliqIcmsST	real not null,
			icmsST		real not null,
			itemId		varchar(64) null,
			PRIMARY KEY (
				periodo,
				cnpjEmit,
				ufEmit,
				serie,
				numero,
				modelo,
				numItem
			)
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_Itens_IcmsIdx ON EFD_Itens (
			icms
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_Itens_CfopIdx ON EFD_Itens (
			cfop
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_Itens_AliqIdx ON EFD_Itens (
			aliq
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_Itens_CstIdx ON EFD_Itens (
			cst
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_Itens_Cst2Idx ON EFD_Itens (
			cst_origem,
			cst_tribut
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into EFD_Itens
		(periodo, cnpjEmit, ufEmit, serie, numero, modelo, numItem, cst, cst_origem, cst_tribut, cfop, qtd, valorProd, valorDesc, bc, aliq, icms, bcIcmsST, aliqIcmsST, icmsST, itemId) 
		values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
	]]
	
end

----------------------------------------------------------------------
-- criar tabela LRS
function criarTabela_EFD_LRS(db)

	db_execNonQuery( db, [[
		create table EFD_LRS( 
			periodo		integer not null,
			serie		short not null,
			numero		integer not null,
			modelo		short not null,
			dataEmit	integer not null,
			valorOp		real not null,
			chave		char(44) null,
			cnpjDest	bigint not null,
			ufDest		short not null,
			ieDest		varchar(20) null,
			PRIMARY KEY (
				periodo,
				serie,
				numero,
				modelo
			)
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRS_cnpjUfSerieNumeroIdx ON EFD_LRS (
			serie,
			numero,
			modelo,
			cnpjDest,
			ufDest
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRS_serieNumeroIdx ON EFD_LRS (
			serie,
			numero,
			modelo
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRS_chaveIdx ON EFD_LRS (
			chave
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRS_cnpjDestIdx ON EFD_LRS (
			cnpjDest,
			ufDest
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRS_ufDestIdx ON EFD_LRS (
			ufDest
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_LRS_ieDestIdx ON EFD_LRS (
			ieDest
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into EFD_LRS 
			(periodo, cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp, ieDest) 
			values (?,?,?,?,?,?,?,?,?,?)
	]]
	
end

----------------------------------------------------------------------
-- criar tabela registros analíticos (LRE e LRS)
function criarTabela_EFD_Anal(db)

	db_execNonQuery( db, [[
		create table EFD_Anal( 
			operacao	short not null,
			periodo		integer not null,
			cnpj		bigint not null,
			uf			short not null,
			serie		short not null,
			numero		integer not null,
			modelo		short not null,
			numReg		short not null,
			cst			short not null,
			cst_origem	short not null,
			cst_tribut	short not null,
			cfop		short not null,
			aliq		real not null,
			valorOp		real not null,
			bc			real not null,
			icms		real not null,
			bcIcmsST	real null,
			icmsST		real null,
			redBC		real null,
			ipi			real null,
			PRIMARY KEY (
				periodo,
				cnpj,
				uf,
				serie,
				numero,
				modelo,
				numReg
			)
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_Anal_opIdx ON EFD_Anal (
			operacao
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_Anal_cfopIdx ON EFD_Anal (
			cfop
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_Anal_cstIdx ON EFD_Anal (
			cst
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX EFD_Anal_cst2Idx ON EFD_Anal (
			cst_origem,
			cst_tribut
		) 
	]])

	-- retornar a query que será usada no insert
	return [[
		insert into EFD_Anal 
			(operacao, periodo, cnpj, uf, serie, numero, modelo, numReg, cst, cst_origem, cst_tribut, cfop, aliq, valorOp, bc, icms, bcIcmsST, icmsST, redBC, ipi) 
			values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
	]]
	
end

----------------------------------------------------------------------
-- criar tabela de itens de ressarcimento ST (há n itens para cada ItemNf)
function criarTabela_EFD_Ressarc_Itens(db)

	db_execNonQuery( db, [[
		create table EFD_Ressarc_Itens( 
			periodo		integer not null,
			cnpjEmit	bigint not null,
			ufEmit		short not null,
			serie		short not null,
			numero		integer not null,
			modelo		short not null,
			nroItem		short not null,
			cnpjUlt		bigint not null,
			ufUlt		short not null,
			serieUlt	short not null,
			numeroUlt	integer not null,
			modeloUlt	short not null,
			dataUlt		integer not null,
			valorUlt	real not null,
			bcSTUlt		real not null,
			qtdUlt		real not null,
			chaveUlt	char(44) null,
			nroItemUlt	short null,
			PRIMARY KEY (
				periodo,
				cnpjEmit,
				ufEmit,
				serie,
				numero,
				modelo,
				nroItem
			)
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX EFD_Ressarc_Itens_cnpjUfSerieNumeroIdx ON EFD_Ressarc_Itens (
			periodo,
			cnpjUlt,
			ufUlt,
			serieUlt,
			numeroUlt,
			modeloUlt
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX EFD_Ressarc_Itens_chaveUltIdx ON EFD_Ressarc_Itens (
			chaveUlt
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX EFD_Ressarc_Itens_nroItemUltIdx ON EFD_Ressarc_Itens (
			nroItemUlt
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into EFD_Ressarc_Itens 
			(periodo, cnpjEmit, ufEmit, serie, numero, modelo, nroItem, cnpjUlt, ufUlt, serieUlt, numeroUlt, modeloUlt, chaveUlt, dataUlt, valorUlt, bcSTUlt, qtdUlt, nroItemUlt) 
			values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
	]]
	
end

-- criar tabela itensId
function criarTabela_EFD_ItensId(db)

	db_execNonQuery( db, [[
		create table EFD_ItensId( 
			id			varchar(64) not null,
			descricao	varchar(1024) not null,
			ncm			bigint null,
			cest		integer null,
			aliqInt		real null,
			PRIMARY KEY (
				id
			)
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into EFD_ItensId 
			(id, descricao, ncm, cest, aliqInt) 
			values (?,?,?,?,?)
	]]
	
end

-- criar tabela mestre
function criarTabela_EFD_Mestre(db)

	db_execNonQuery( db, [[
		create table EFD_Mestre( 
			versao		integer not null,
			original	bit not null,
			dataIni		integer not null,
			dataFim		integer not null,
			nome		varchar(100) not null,
			cnpj        bigint not null,
			uf			char(2) not null,
			ie			varchar(14) null,
			PRIMARY KEY (
				cnpj,
				dataIni,
				dataFim
			)
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into EFD_Mestre 
			(versao, original, dataIni, dataFim, nome, cnpj, uf, ie) 
			values (?,?,?,?,?,?,?,?)
	]]
	
end

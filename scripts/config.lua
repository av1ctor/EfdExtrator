
----------------------------------------------------------------------
function configurarDB(db, dbPath)
	db_execNonQuery(db, "attach '" .. dbPath .. "config.db' as conf")
	db_execNonQuery(db, "attach '" .. dbPath .. "CadContribuinte.db' as cdb")
	db_execNonQuery(db, "attach '" .. dbPath .. "inidoneos.db' as idb")
	db_execNonQuery(db, "attach '" .. dbPath .. "GIA.db' as gdb")
end

----------------------------------------------------------------------
-- criar tabela de dfe's de entrada (relatórios do SAFI ou do BO)
function criarTabela_dfeEntrada(db)

	db_execNonQuery( db, [[
		create table dfeEntrada(
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
		CREATE INDEX chaveDfeEntradaIdx ON dfeEntrada (
			cnpjEmit,
			ufEmit,
			serie,
			numero,
			modelo
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX cnpjDfeEntradaEmitIdx ON dfeEntrada (
			cnpjEmit,
			ufEmit
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX ieEmitDfeEntradaIdx ON dfeEntrada (
			ieEmit
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into dfeEntrada 
			(cnpjEmit, ufEmit, serie, numero, modelo, chave, dataEmit, valorOp, ieEmit)
			values (?,?,?,?,?,?,?,?,?)
	]]

end

----------------------------------------------------------------------
-- criar tabela de dfe's de saída (relatórios do SAFI ou do BO)
function criarTabela_dfeSaida(db)

	db_execNonQuery( db, [[
		create table dfeSaida( 
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
		CREATE INDEX chaveDfeSaidaIdx ON dfeSaida (
			serie,
			numero,
			modelo,
			cnpjDest,
			ufDest
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX cnpjDfeSaidaDestIdx ON dfeSaida (
			cnpjDest,
			ufDest
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX ieDestDfeSaidaIdx ON dfeSaida (
			ieDest
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into dfeSaida 
			(cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp, ieDest) 
			values (?,?,?,?,?,?,?,?,?)
	]]

end

----------------------------------------------------------------------
-- criar tabela de itens de docs saída (relatórios do SAFI ou do BO)
function criarTabela_itensDfeSaida(db)

	db_execNonQuery( db, [[
		create table itensDfeSaida( 
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
		CREATE INDEX itensDfeSaidaChaveIdx ON itensDfeSaida (
			chave
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX itensDfeSaidaCfopIdx ON itensDfeSaida (
			cfop
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX itensDfeSaidaAliqIdx ON itensDfeSaida (
			aliq
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX itensDfeSaidaNcmIdx ON itensDfeSaida (
			ncm
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX itensDfeSaidaCstIdx ON itensDfeSaida (
			cst
		) 
	]])

	-- retornar a query que será usada no insert
	return [[
		insert into itensDfeSaida 
			(serie, numero, modelo, numItem, chave, cfop, valorProd, valorDesc, valorAcess, bc, aliq, icms, bcIcmsST, ncm, cst, qtd, unidade, codProduto, descricao) 
			values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
	]]

end

----------------------------------------------------------------------
-- criar tabela LRE
function criarTabela_LRE(db)

	db_execNonQuery( db, [[
		create table LRE( 
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
		CREATE INDEX cnpjUfSerieNumeroLREIdx ON LRE (
			cnpjEmit,
			ufEmit,
			serie,
			numero,
			modelo
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX chaveLREIdx ON LRE (
			chave
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX cnpjEmitIdx ON LRE (
			cnpjEmit,
			ufEmit
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX ufEmitIdx ON LRE (
			ufEmit
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX ieEmitIdx ON LRE (
			ieEmit
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into LRE 
			(periodo, cnpjEmit, ufEmit, serie, numero, modelo, chave, dataEmit, valorOp, ieEmit) 
			values (?,?,?,?,?,?,?,?,?,?)
	]]
	
end

----------------------------------------------------------------------
-- criar tabela itens de NF da LRE (ou LRS no caso de ressarcimento ST)
function criarTabela_itensNfLR(db)

	db_execNonQuery( db, [[
		create table itensNfLR( 
			periodo		integer not null,
			cnpjEmit	bigint not null,
			ufEmit		short not null,
			serie		short not null,
			numero		integer not null,
			modelo		short not null,
			numItem		short not null,
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
		CREATE INDEX itensNfLRIcmsIdx ON itensNfLR (
			icms
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX itensNfLRCfopIdx ON itensNfLR (
			cfop
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX itensNfLRAliqIdx ON itensNfLR (
			aliq
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX itensNfLRCstIdx ON itensNfLR (
			cst_origem,
			cst_tribut
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into itensNfLR 
		(periodo, cnpjEmit, ufEmit, serie, numero, modelo, numItem, cst_origem, cst_tribut, cfop, qtd, valorProd, valorDesc, bc, aliq, icms, bcIcmsST, aliqIcmsST, icmsST, itemId) 
		values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
	]]
	
end

----------------------------------------------------------------------
-- criar tabela LRS
function criarTabela_LRS(db)

	db_execNonQuery( db, [[
		create table LRS( 
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
		CREATE INDEX cnpjUfSerieNumeroLRSIdx ON LRS (
			serie,
			numero,
			modelo,
			cnpjDest,
			ufDest
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX serieNumeroLRSIdx ON LRS (
			serie,
			numero,
			modelo
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX chaveLRSIdx ON LRS (
			chave
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX cnpjDestLRSIdx ON LRS (
			cnpjDest,
			ufDest
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX ufDestLRSIdx ON LRS (
			ufDest
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX ieDestIdx ON LRS (
			ieDest
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into LRS 
			(periodo, cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp, ieDest) 
			values (?,?,?,?,?,?,?,?,?,?)
	]]
	
end

----------------------------------------------------------------------
-- criar tabela de itens de ressarcimento ST (há n itens para cada ItemNf)
function criarTabela_ressarcStItensNfLRS(db)

	db_execNonQuery( db, [[
		create table ressarcStItensNfLRS( 
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
		CREATE INDEX cnpjUfSerieNumeroRessarcStLRSIdx ON ressarcStItensNfLRS (
			periodo,
			cnpjUlt,
			ufUlt,
			serieUlt,
			numeroUlt,
			modeloUlt
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX chaveUltItensNfLRSIdx ON ressarcStItensNfLRS (
			chaveUlt
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX nroItemUltItensNfLRSIdx ON ressarcStItensNfLRS (
			nroItemUlt
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into ressarcStItensNfLRS 
			(periodo, cnpjEmit, ufEmit, serie, numero, modelo, nroItem, cnpjUlt, ufUlt, serieUlt, numeroUlt, modeloUlt, chaveUlt, dataUlt, valorUlt, bcSTUlt, qtdUlt, nroItemUlt) 
			values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
	]]
	
end

-- criar tabela itensId
function criarTabela_itensId(db)

	db_execNonQuery( db, [[
		create table itensId( 
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
		insert into itensId 
			(id, descricao, ncm, cest, aliqInt) 
			values (?,?,?,?,?)
	]]
	
end

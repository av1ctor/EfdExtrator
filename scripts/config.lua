
----------------------------------------------------------------------
function configurarDB(db, dbPath)
	db_execNonQuery(db, "attach '" .. dbPath .. "CadContribuinte.db' as cdb")
	db_execNonQuery(db, "attach '" .. dbPath .. "inidoneos.db' as idb")
	db_execNonQuery(db, "attach '" .. dbPath .. "GIA.db' as gdb")
end

----------------------------------------------------------------------
-- criar tabela de dfe's de entrada (relatórios do SAFI)
function criarTabela_dfeEntrada(db)

	db_execNonQuery( db, [[
		create table dfeEntrada(
			chave		char(44) not null,
			cnpjEmit	bigint not null,
			ufEmit		bigint not null,
			serie		integer not null,
			numero		integer not null,
			modelo		integer not null,
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
-- criar tabela de dfe's de saída (relatórios do SAFI)
function criarTabela_dfeSaida(db)

	db_execNonQuery( db, [[
		create table dfeSaida( 
			chave		char(44) not null,
			serie		integer not null,
			numero		integer not null,
			modelo		integer not null,
			dataEmit	integer not null,
			valorOp		real not null,
			cnpjDest	bigint not null,
			ufDest		bigint not null,
			ieDest		varchar(20) null,
			PRIMARY KEY (
				chave
			)
		) 
	]])
	
	db_execNonQuery( db, [[ 
		CREATE INDEX chaveDfeSaidaIdx ON dfeSaida (
			cnpjDest,
			ufDest,
			serie,
			numero,
			modelo
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
-- criar tabela de itens de docs saída (relatórios do SAFI)
function criarTabela_itensDfeSaida(db)

	db_execNonQuery( db, [[
		create table itensDfeSaida( 
			chave		char(44) not null,
			cfop		integer not null,
			valorProd	real not null,
			valorDesc	real not null,
			valorAcess	real not null,
			bc			real not null,
			aliq		real not null,
			icms		real not null,
			bcIcmsST	real not null,
			PRIMARY KEY (
				chave
			)
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

	-- retornar a query que será usada no insert
	return [[
		insert into itensDfeSaida 
			(chave, cfop, valorProd, valorDesc, valorAcess, bc, aliq, icms, bcIcmsST) 
			values (?,?,?,?,?,?,?,?,?)
	]]

end

----------------------------------------------------------------------
-- criar tabela LRE
function criarTabela_LRE(db)

	db_execNonQuery( db, [[
		create table LRE( 
			periodo		integer not null,
			cnpjEmit	bigint not null,
			ufEmit		bigint not null,
			serie		integer not null,
			numero		integer not null,
			modelo		integer not null,
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
-- criar tabela itens de NF da LRE
function criarTabela_itensNfLRE(db)

	db_execNonQuery( db, [[
		create table itensNfLRE( 
			periodo		integer not null,
			cnpjEmit	bigint not null,
			ufEmit		bigint not null,
			serie		integer not null,
			numero		integer not null,
			modelo		integer not null,
			numItem		integer not null,
			cst_origem	short not null,
			cst_tribut	short not null,
			cfop		integer not null,
			qtd			real not null,
			valorProd	real not null,
			valorDesc	real not null,
			bc			real not null,
			aliq		real not null,
			icms		real not null,
			bcIcmsST	real not null,
			aliqIcmsST	real not null,
			icmsST		real not null,
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
		CREATE INDEX itensNfLREIcmsIdx ON itensNfLRE (
			icms
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX itensNfLRECfopIdx ON itensNfLRE (
			cfop
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX itensNfLREAliqIdx ON itensNfLRE (
			aliq
		) 
	]])

	db_execNonQuery( db, [[
		CREATE INDEX itensNfLRECstIdx ON itensNfLRE (
			cst_origem,
			cst_tribut
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into itensNfLRE 
		(periodo, cnpjEmit, ufEmit, serie, numero, modelo, numItem, cst_origem, cst_tribut, cfop, qtd, valorProd, valorDesc, bc, aliq, icms, bcIcmsST, aliqIcmsST, icmsST) 
		values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
	]]
	
end

----------------------------------------------------------------------
-- criar tabela LRS
function criarTabela_LRS(db)

	db_execNonQuery( db, [[
		create table LRS( 
			periodo		integer not null,
			serie		integer not null,
			numero		integer not null,
			modelo		integer not null,
			dataEmit	integer not null,
			valorOp		real not null,
			chave		char(44) null,
			cnpjDest	bigint not null,
			ufDest		bigint not null,
			ieDest		varchar(20) null,
			PRIMARY KEY (
				periodo,
				cnpjDest,
				ufDest,
				serie,
				numero,
				modelo
			)
		) 
	]])
	
	db_execNonQuery( db, [[
		CREATE INDEX cnpjUfSerieNumeroLRSIdx ON LRS (
			cnpjDest,
			ufDest,
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
			cnpjUlt		bigint not null,
			ufUlt		bigint not null,
			serieUlt	integer not null,
			numeroUlt	integer not null,
			modeloUlt	integer not null,
			dataUlt		integer not null,
			valorUlt	real not null,
			bcSTUlt		real not null,
			qtdUlt		real not null,
			chaveUlt	char(44) null,
			nroItemUlt	integer null
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
		CREATE INDEX chaveItensNfLRSIdx ON ressarcStItensNfLRS (
			chaveUlt
		) 
	]])
	
	-- retornar a query que será usada no insert
	return [[
		insert into ressarcStItensNfLRS 
			(periodo, cnpjUlt, ufUlt, serieUlt, numeroUlt, modeloUlt, chaveUlt, dataUlt, valorUlt, bcSTUlt, qtdUlt, nroItemUlt) 
			values (?,?,?,?,?,?,?,?,?,?,?,?)
	]]
	
end


#include once "efd.bi"
#include once "bfile.bi"
#include once "Dict.bi"
#include once "ExcelWriter.bi"
#include once "vbcompat.bi"
#include once "ssl_helper.bi"
#include once "DB.bi"
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"
#include once "trycatch.bi"

const ASSINATURA_P7K_HEADER = "SBRCAAEPDR"

private function my_lua_Alloc cdecl _
	( _
		byval ud as any ptr, _
		byval p as any ptr, _
		byval osize as uinteger, _
		byval nsize as uinteger _
	) as any ptr

	if( nsize = 0 ) then
		deallocate( p )
		function = NULL
	else
		function = reallocate( p, nsize )
	end if

end function

''''''''
constructor Efd(onProgress as OnProgressCB, onError as OnErrorCB)
	
	'' eventos
	this.onProgress = onProgress
	this.onError = onError
	
	''
	chaveDFeDict = new TDict(2^20)
	nfeDestSafiFornecido = false
	nfeEmitSafiFornecido = false
	itemNFeSafiFornecido = false
	cteSafiFornecido = false
	dfeListHead = null
	dfeListTail = null
	arquivos = new TList(10, len(TArquivoInfo))
	
	''
	baseTemplatesDir = ExePath + "\templates\"
	
	municipDict = new TDict(2^10, true, true, true)
	
	''
	dbConfig = new TDb
	dbConfig->open(ExePath + "\db\config.db")
	
end constructor

destructor Efd()

	''
	dbConfig->close()
	delete dbConfig
	
	''
	delete municipDict
	
	''
	delete chaveDFeDict
	
	do while dfeListHead <> null
		var next_ = dfeListHead->next_
		if dfeListHead->modelo = NFE then
			do while dfeListHead->nfe.itemListHead <> null
				var next_ = dfeListHead->nfe.itemListHead->next_
				delete dfeListHead->nfe.itemListHead
				dfeListHead->nfe.itemListHead = next_
			loop
		end if
		delete dfeListHead
		dfeListHead = next_
	loop
	
	delete arquivos
	
end destructor

''''''''
private sub lua_carregarCustoms(d as TDict ptr, L as lua_State ptr) 

	lua_getglobal(L, "getCustomCallbacks")
	lua_call(L, 0, 1)
	if lua_isnil(L, -1) = 0 then
		lua_pushnil(L)
		do while lua_next(L, -2) <> 0
			var key = lua_tostring(L, -2)
			
			var lcb = new CustomLuaCb
			lua_pushnil(L)
			do while lua_next(L, -2) <> 0
				
				var funct = dupstr(lua_tostring(L, -1)) 
				select case *lua_tostring(L, -2)
				case "reader"
					lcb->reader = funct
				case "writer"
					lcb->writer = funct
				case "rel_entradas"
					lcb->rel_entradas = funct
				case "rel_saidas"
					lcb->rel_saidas = funct
				case "rel_outros"
					lcb->rel_outros = funct
				end select
				
				d->add(key, lcb)
				lua_pop(L, 1)
			loop
				
			lua_pop(L, 1)
		loop
		lua_pop(L, lua_gettop(L))
	end if

end sub

''''''''
sub EFd.configurarScripting()
	try
		lua = lua_newstate(@my_lua_Alloc, NULL)
		luaL_openlibs(lua)
		
		TDb.exportAPI(lua)
		ExcelWriter.exportAPI(lua)
		bfile.exportAPI(lua)
		exportAPI(lua)

		luaL_dofile(lua, ExePath + "\scripts\config.lua")
		luaL_dofile(lua, ExePath + "\scripts\customizacao.lua")	
		
		customLuaCbDict = new TDict(16, true, true, true)
		lua_carregarCustoms(customLuaCbDict, lua)
	catch
		onError("Erro ao carregar script lua. Verifique erros de sintaxe")
	endtry
end sub

''''''''
private function lua_criarTabela(lua as lua_State ptr, db as TDb ptr, tabela as const zstring ptr, onError as OnErrorCB) as TDbStmt ptr

	try
		lua_getglobal(lua, "criarTabela_" + *tabela)
		lua_pushlightuserdata(lua, db)
		lua_call(lua, 1, 1)
		var res = db->prepare(lua_tostring(lua, -1))
		if res = null then
			onError("Erro ao executar script lua de cria��o de tabela: " + "criarTabela_" + *tabela + ": " + *db->getErrorMsg())
		end if
		function = res
		lua_pop(lua, 1)
	catch
		onError("Erro ao executar script lua de cria��o de tabela: " + "criarTabela_" + *tabela + ". Verifique erros de sintaxe")
	endtry

end function

''''''''
sub Efd.configurarDB()

	db = new TDb
	if not opcoes.dbEmDisco then
		db->open()
	else
		kill nomeArquivoSaida + ".db"
		db->open(nomeArquivoSaida + ".db")
		db->execNonQuery("PRAGMA JOURNAL_MODE=OFF")
		db->execNonQuery("PRAGMA SYNCHRONOUS=0")
		db->execNonQuery("PRAGMA LOCKING_MODE=EXCLUSIVE")
	end if

	var dbPath = ExePath + "\db\"
	
	try
		'' chamar configurarDB()
		lua_getglobal(lua, "configurarDB")
		lua_pushlightuserdata(lua, db)
		lua_pushstring(lua, dbPath)
		lua_call(lua, 2, 0)

		'' criar tabelas
		db_dfeEntradaInsertStmt = lua_criarTabela(lua, db, "DFe_Entradas", onError)

		db_dfeSaidaInsertStmt = lua_criarTabela(lua, db, "DFe_Saidas", onError)
		
		db_itensDfeSaidaInsertStmt = lua_criarTabela(lua, db, "DFe_Saidas_Itens", onError)
		
		db_LREInsertStmt = lua_criarTabela(lua, db, "EFD_LRE", onError)

		db_itensNfLRInsertStmt = lua_criarTabela(lua, db, "EFD_Itens", onError)

		db_LRSInsertStmt = lua_criarTabela(lua, db, "EFD_LRS", onError)
		
		db_analInsertStmt = lua_criarTabela(lua, db, "EFD_Anal", onError)

		db_ressarcStItensNfLRSInsertStmt = lua_criarTabela(lua, db, "EFD_Ressarc_Itens", onError)
		
		db_itensIdInsertStmt = lua_criarTabela(lua, db, "EFD_ItensId", onError)
		
		db_mestreInsertStmt = lua_criarTabela(lua, db, "EFD_Mestre", onError)
		
		if db_dfeEntradaInsertStmt = null or _
			db_dfeSaidaInsertStmt = null or _
			 db_itensDfeSaidaInsertStmt = null or _
			  db_LREInsertStmt = null or _
			   db_itensNfLRInsertStmt = null or _
			    db_LRSInsertStmt = null or _
				 db_ressarcStItensNfLRSInsertStmt = null or _
					db_itensIdInsertStmt = null or _
						db_analInsertStmt = null then
			
		end if
	catch
		onError("Erro ao executar script lua de cria��o de DB. Verifique erros de sintaxe")
	endtry

end sub   

''''''''
sub Efd.fecharDb()
	if db <> null then
		if db_dfeEntradaInsertStmt <> null then
			delete db_dfeEntradaInsertStmt
		end if
		if db_dfeSaidaInsertStmt <> null then
			delete db_dfeSaidaInsertStmt
		end if
		if db_itensDfeSaidaInsertStmt <> null then
			delete db_itensDfeSaidaInsertStmt
		end if
		if db_LREInsertStmt <> null then
			delete db_LREInsertStmt
		end if
		if db_itensNfLRInsertStmt <> null then
			delete db_itensNfLRInsertStmt
		end if
		if db_LRSInsertStmt <> null then
			delete db_LRSInsertStmt
		end if
		if db_analInsertStmt <> null then
			delete db_analInsertStmt
		end if
		if db_ressarcStItensNfLRSInsertStmt <> null then
			delete db_ressarcStItensNfLRSInsertStmt
		end if
		if db_itensIdInsertStmt <> null then
			delete db_itensIdInsertStmt
		end if
		if db_mestreInsertStmt <> null then
			delete db_mestreInsertStmt
		end if
		
		db->close()
		delete db
		db = null
	end if
end sub
  
''''''''
sub Efd.iniciarExtracao(nomeArquivo as String, opcoes as OpcoesExtracao)
	
	''
	ew = new ExcelWriter
	ew->create(nomeArquivo, opcoes.formatoDeSaida)

	entradas = null
	saidas = null
	nomeArquivoSaida = nomeArquivo
	this.opcoes = opcoes
	
	''
	configurarScripting()

	''
	configurarDB()
	
end sub

''''''''
sub Efd.finalizarExtracao()

	''
	onProgress("Gravando planilha: " + nomeArquivoSaida, 0)
	ew->Flush(onProgress)
	ew->Close
	delete ew
   
	''
	fecharDb()
	if opcoes.dbEmDisco then
		if not opcoes.manterDb then
			kill nomeArquivoSaida + ".db"
		end if
	end if
	
	''
	lua_close( lua )
	
end sub

''''''''
private sub pularLinha(bf as bfile) 

	'ler at� \r
	do
		var c = bf.char1
		
		if c = 13 or c = 10 then
			exit do
		end if
	loop

	'pular \n
	if bf.peek1 = 10 then
		bf.char1 
	end if
	
end sub

''''''''
private function lerLinha(bf as bfile) as string

	var res = ""
	var c = " "
	
	'ler at� \r
	do
		c[0] = bf.char1
		if c[0] = 13 or c[0] = 10 then
			exit do
		end if
		
		res += c
	loop
	
	'pular \n
	if bf.peek1 = 10 then
		bf.char1 
	end if

	function = res
	
end function

''''''''
function Efd.lerTipo(bf as bfile, tipo as zstring ptr) as TipoRegistro

	if bf.peek1 <> asc("|") then
		onError("Erro: fora de sincronia na linha:" & nroLinha)
	else
		bf.char1 ' pular |
	end if
	
	*tipo = bf.char4
	var subtipo = valint(right(*tipo, 3))

	var tp = DESCONHECIDO
	
	select case as const tipo[0]
	case asc("0")
		select case subtipo
		case 150
			tp = PARTICIPANTE
		case 200
			tp = ITEM_ID
		case 300
			tp = BEM_CIAP
		case 450
			tp = INFO_COMPL
		case 000
			tp = MESTRE
		end select
	case asc("C")
		select case subtipo
		case 100
			tp = DOC_NF
		case 110
			tp = DOC_NF_INFO
		case 170
			tp = DOC_NF_ITEM
		case 176
			tp = DOC_NF_ITEM_RESSARC_ST
		case 190
			tp = DOC_NF_ANAL
		case 101
			tp = DOC_NF_DIFAL
		case 460
			tp = DOC_ECF
		case 470
			tp = DOC_ECF_ITEM
		case 490
			tp = DOC_ECF_ANAL
		case 400
			tp = EQUIP_ECF
		case 405
			tp = ECF_REDUCAO_Z
		case 500
			tp = DOC_NF_ELETRIC
		case 590
			tp = DOC_NF_ELETRIC_ANAL
		case 800
			tp = DOC_SAT
		case 850
			tp = DOC_SAT_ANAL
		end select
	case asc("D")
		select case subtipo
		case 100
			tp = DOC_CT
		case 190
			tp = DOC_CT_ANAL
		case 101
			tp = DOC_CT_DIFAL
		case 500
			tp = DOC_NFSCT
		case 590
			tp = DOC_NFSCT_ANAL
		end select
	case asc("E")	
		select case subtipo
		case 100
			tp = APURACAO_ICMS_PERIODO
		case 110
			tp = APURACAO_ICMS_PROPRIO
		case 111
			tp = APURACAO_ICMS_AJUSTE
		case 200
			tp = APURACAO_ICMS_ST_PERIODO
		case 210
			tp = APURACAO_ICMS_ST
		end select
	case asc("G")
		select case subtipo
		case 110
			tp = CIAP_TOTAL
		case 125
			tp = CIAP_ITEM
		case 130
			tp = CIAP_ITEM_DOC
		end select
	case asc("H")	
		select case subtipo
		case 005
			tp =  INVENTARIO_TOTAIS
		case 010
			tp =  INVENTARIO_ITEM
		end select
	case asc("K")
		select case subtipo
		case 100
			tp = ESTOQUE_PERIODO
		case 200
			tp = ESTOQUE_ITEM
		case 230
			tp = ESTOQUE_ORDEM_PROD
		end select
	case asc("9")
		select case subtipo
		case 999
			tp = FIM_DO_ARQUIVO
		end select
	end select
	
	if tp = DESCONHECIDO then
		if customLuaCbDict->lookup(*tipo) <> null then
			tp = LUA_CUSTOM
		end if
	end if
	
	function = tp

end function

''''''''
function Efd.lerRegMestre(bf as bfile, reg as TRegistro ptr) as Boolean
   
	bf.char1		'pular |

	reg->mestre.versaoLayout= bf.varint
	reg->mestre.original 	= (bf.int1 = 0)
	bf.char1		'pular |
	reg->mestre.dataIni		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->mestre.dataFim		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->mestre.nome	   	= bf.varchar
	reg->mestre.cnpj	   	= bf.varchar
	reg->mestre.cpf	   		= bf.varint
	reg->mestre.uf			= bf.varchar
	reg->mestre.ie			= bf.varchar
	reg->mestre.municip		= bf.varint
	reg->mestre.im  		= bf.varchar
	reg->mestre.suframa  	= bf.varchar
	reg->mestre.perfil  	= bf.char1
	bf.char1		'pular |
	reg->mestre.atividade	= bf.int1
	bf.char1		'pular |

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function Efd.lerRegParticipante(bf as bfile, reg as TRegistro ptr) as Boolean
   
	bf.char1		'pular |

	reg->part.id		= bf.varchar
	reg->part.nome		= bf.varchar
	reg->part.pais	   	= bf.varint
	reg->part.cnpj	   	= bf.varchar
	reg->part.cpf	   	= bf.varchar
	reg->part.ie		= bf.varchar
	reg->part.municip	= bf.varint
	reg->part.suframa  	= bf.varchar
	reg->part.ender	   	= bf.varchar
	reg->part.num		= bf.varchar
	reg->part.compl	   	= bf.varchar
	reg->part.bairro	= bf.varchar
   
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocNF(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->nf.operacao		= bf.int1
	bf.char1		'pular |
	reg->nf.emitente		= bf.int1
	bf.char1		'pular |
	reg->nf.idParticipante	= bf.varchar
	reg->nf.modelo			= bf.int2
	bf.char1		'pular |
	reg->nf.situacao		= bf.int2
	bf.char1		'pular |
	reg->nf.serie			= bf.varchar
	reg->nf.numero			= bf.varint
	reg->nf.chave			= bf.varchar
	reg->nf.dataEmi			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.valorTotal		= bf.vardbl
	reg->nf.pagamento		= bf.varint
	reg->nf.valorDesconto	= bf.vardbl
	reg->nf.valorAbatimento	= bf.vardbl
	reg->nf.valorMerc		= bf.vardbl
	reg->nf.frete			= bf.varint
	reg->nf.valorFrete		= bf.vardbl
	reg->nf.valorSeguro		= bf.vardbl
	reg->nf.valorAcessorias	= bf.vardbl
	reg->nf.bcICMS			= bf.vardbl
	reg->nf.ICMS			= bf.vardbl
	reg->nf.bcICMSST		= bf.vardbl
	reg->nf.ICMSST			= bf.vardbl
	reg->nf.IPI				= bf.vardbl
	reg->nf.PIS				= bf.vardbl
	reg->nf.COFINS			= bf.vardbl
	reg->nf.PISST			= bf.vardbl
	reg->nf.COFINSST		= bf.vardbl
	reg->nf.nroItens		= 0

	reg->nf.itemAnalListHead = null
	reg->nf.itemAnalListTail = null
	reg->nf.infoComplListHead = null
	reg->nf.infoComplListTail = null

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocNFInfo(bf as bfile, reg as TRegistro ptr, pai as TDocNF ptr) as Boolean

	bf.char1		'pular |

	reg->docInfoCompl.idCompl			= bf.varchar
	reg->docInfoCompl.extra				= bf.varchar
	reg->docInfoCompl.next_				= null
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocNFItem(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNF ptr) as Boolean

	bf.char1		'pular |

	reg->itemNF.documentoPai	= documentoPai
   
	reg->itemNF.numItem			= bf.varint
	reg->itemNF.itemId			= bf.varchar
	reg->itemNF.descricao		= bf.varchar
	reg->itemNF.qtd				= bf.vardbl
	reg->itemNF.unidade			= bf.varchar
	reg->itemNF.valor			= bf.vardbl
	reg->itemNF.desconto		= bf.vardbl
	reg->itemNF.indMovFisica	= bf.varint
	reg->itemNF.cstICMS			= bf.varint
	reg->itemNF.cfop			= bf.varint
	reg->itemNF.codNatureza		= bf.varchar
	reg->itemNF.bcICMS			= bf.vardbl
	reg->itemNF.aliqICMS		= bf.vardbl
	reg->itemNF.ICMS			= bf.vardbl
	reg->itemNF.bcICMSST		= bf.vardbl
	reg->itemNF.aliqICMSST		= bf.vardbl
	reg->itemNF.ICMSST			= bf.vardbl
	reg->itemNF.indApuracao		= bf.varint
	reg->itemNF.cstIPI			= bf.varint
	reg->itemNF.codEnqIPI		= bf.varchar
	reg->itemNF.bcIPI			= bf.vardbl
	reg->itemNF.aliqIPI			= bf.vardbl
	reg->itemNF.IPI				= bf.vardbl
	reg->itemNF.cstPIS			= bf.varint
	reg->itemNF.bcPIS			= bf.vardbl
	reg->itemNF.aliqPISPerc		= bf.vardbl
	reg->itemNF.qtdBcPIS		= bf.vardbl
	reg->itemNF.aliqPISMoed		= bf.vardbl
	reg->itemNF.PIS				= bf.vardbl
	reg->itemNF.cstCOFINS		= bf.varint
	reg->itemNF.bcCOFINS		= bf.vardbl
	reg->itemNF.aliqCOFINSPerc 	= bf.vardbl
	reg->itemNF.qtdBcCOFINS		= bf.vardbl
	reg->itemNF.aliqCOFINSMoed 	= bf.vardbl
	reg->itemNF.COFINS			= bf.vardbl
	bf.varchar					'' pular c�digo da conta
	if regMestre->mestre.versaoLayout >= 013 then
		bf.vardbl				'' pular VL_ABAT_NT
	end if

	documentoPai->nroItens 		+= 1
	
	reg->itemNF.itemRessarcStListHead = null
	reg->itemNF.itemRessarcStListTail = null

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocNFItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= documentoPai
	reg->itemAnal.num		= documentoPai->nf.itemAnalCnt
	documentoPai->nf.itemAnalCnt += 1
	
	reg->itemAnal.cst		= bf.varint
	reg->itemAnal.cfop		= bf.varint
	reg->itemAnal.aliq		= bf.vardbl
	reg->itemAnal.valorOp	= bf.vardbl
	reg->itemAnal.bc		= bf.vardbl
	reg->itemAnal.ICMS		= bf.vardbl
	reg->itemAnal.bcST		= bf.vardbl
	reg->itemAnal.ICMSST	= bf.vardbl
	reg->itemAnal.redBC		= bf.vardbl
	reg->itemAnal.IPI		= bf.vardbl
	bf.varchar					'' pular c�digo de observa��o

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function Efd.lerRegDocNFItemRessarcSt(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNFItem ptr) as Boolean

	bf.char1		'pular |

	reg->itemRessarcSt.documentoPai	= documentoPai
	
	reg->itemRessarcSt.modeloUlt 			= bf.int2
	bf.char1		'pular |
	reg->itemRessarcSt.numeroUlt 			= bf.varint
	reg->itemRessarcSt.serieUlt  			= bf.varchar
	reg->itemRessarcSt.dataUlt				= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->itemRessarcSt.idParticipanteUlt	= bf.varchar
	reg->itemRessarcSt.qtdUlt				= bf.vardbl
	reg->itemRessarcSt.valorUlt				= bf.vardbl
	reg->itemRessarcSt.valorBcST			= bf.vardbl
	
	if bf.peek1 <> 13 then
		reg->itemRessarcSt.chaveNFeUlt		= bf.varchar
		reg->itemRessarcSt.numItemNFeUlt	= bf.varint
		reg->itemRessarcSt.bcIcmsUlt		= bf.vardbl
		reg->itemRessarcSt.aliqIcmsUlt		= bf.vardbl
		reg->itemRessarcSt.limiteBcIcmsUlt	= bf.vardbl
		reg->itemRessarcSt.icmsUlt			= bf.vardbl
		reg->itemRessarcSt.aliqIcmsStUlt	= bf.vardbl
		reg->itemRessarcSt.res				= bf.vardbl
		reg->itemRessarcSt.responsavelRet	= bf.int1
		bf.char1		'pular |
		reg->itemRessarcSt.motivo			= bf.int1
		bf.char1		'pular |
		reg->itemRessarcSt.chaveNFeRet		= bf.varchar
		reg->itemRessarcSt.idParticipanteRet= bf.varchar
		reg->itemRessarcSt.serieRet			= bf.varchar
		reg->itemRessarcSt.numeroRet		= bf.varint
		reg->itemRessarcSt.numItemNFeRet 	= bf.varint
		reg->itemRessarcSt.tipDocArrecadacao= bf.int1
		bf.char1		'pular |
		reg->itemRessarcSt.numDocArrecadacao= bf.varchar
	end if
   
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function Efd.lerRegDocNFDifal(bf as bfile, reg as TRegistro ptr, documentoPai as TDocNF ptr) as Boolean

	bf.char1		'pular |

	documentoPai->difal.fcp			= bf.vardbl
	documentoPai->difal.icmsDest	= bf.vardbl
	documentoPai->difal.icmsOrigem	= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocCT(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->ct.operacao		= bf.int1
	bf.char1		'pular |
	reg->ct.emitente		= bf.int1
	bf.char1		'pular |
	reg->ct.idParticipante	= bf.varchar
	reg->ct.modelo			= bf.int2
	bf.char1		'pular |
	reg->ct.situacao		= bf.int2
	bf.char1		'pular |
	reg->ct.serie			= bf.varchar
	bf.varchar		'pular sub-s�rie
	reg->ct.numero			= bf.varint
	reg->ct.chave			= bf.varchar
	reg->ct.dataEmi			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ct.dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ct.tipoCTe			= bf.varint
	reg->ct.chaveRef		= bf.varchar
	reg->ct.valorTotal		= bf.vardbl
	reg->ct.valorDesconto	= bf.vardbl
	reg->ct.frete			= bf.varint
	reg->ct.valorServico	= bf.vardbl
	reg->ct.bcICMS			= bf.vardbl
	reg->ct.ICMS			= bf.vardbl
	reg->ct.valorNaoTributado = bf.vardbl
	reg->ct.codInfComplementar= bf.varchar
	bf.varchar		'pular c�digo Conta Analitica
	
	'' c�digos dos munic�pios de origem e de destino n�o aparecem em layouts antigos
	if bf.peek1 <> 13 and bf.peek1 <> 10 then 
		reg->ct.municipioOrigem	= bf.varint
		reg->ct.municipioDestino= bf.varint
	end if
	
	reg->ct.itemAnalListHead = null
	reg->ct.itemAnalListTail = null
	reg->ct.itemAnalCnt = 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocCTItemAnal(bf as bfile, reg as TRegistro ptr, docPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= docPai

	reg->itemAnal.cst			= bf.varint
	reg->itemAnal.cfop			= bf.varint
	reg->itemAnal.aliq			= bf.vardbl
	reg->itemAnal.valorOp		= bf.vardbl
	reg->itemAnal.bc			= bf.vardbl
	reg->itemAnal.ICMS			= bf.vardbl
	reg->itemAnal.redBc			= bf.vardbl
	bf.varchar					'' pular cod obs
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocCTDifal(bf as bfile, reg as TRegistro ptr, docPai as TDocCT ptr) as Boolean

	bf.char1		'pular |

	docPai->difal.fcp		= bf.vardbl
	docPai->difal.icmsDest	= bf.vardbl
	docPai->difal.icmsOrigem= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegEquipECF(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	var modelo = bf.varchar
	reg->equipECF.modelo	= iif(modelo = "2D", &h2D, valint(modelo))
	reg->equipECF.modeloEquip = bf.varchar
	reg->equipECF.numSerie 	= bf.varchar
	reg->equipECF.numCaixa	= bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocECF(bf as bfile, reg as TRegistro ptr, equipECF as TEquipECF ptr) as Boolean

	bf.char1		'pular |

	reg->ecf.equipECF		= equipECF
	reg->ecf.operacao		= SAIDA
	var modelo = bf.varchar
	reg->ecf.modelo			= iif(modelo = "2D", &h2D, valint(modelo))
	reg->ecf.situacao		= bf.int2
	bf.char1		'pular |
	reg->ecf.numero			= bf.varint
	reg->ecf.dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ecf.dataEntSaida	= reg->ecf.dataEmi
	reg->ecf.valorTotal		= bf.vardbl
	reg->ecf.PIS			= bf.vardbl
	reg->ecf.COFINS			= bf.vardbl
	reg->ecf.cpfCnpjAdquirente = bf.varchar
	reg->ecf.nomeAdquirente = bf.varchar
	reg->ecf.nroItens		= 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegECFReducaoZ(bf as bfile, reg as TRegistro ptr, equipECF as TEquipECF ptr) as Boolean

	bf.char1		'pular |

	reg->ecfRedZ.equipECF	= equipECF
	reg->ecfRedZ.dataMov	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ecfRedZ.cro		= bf.varint
	reg->ecfRedZ.crz		= bf.varint
	reg->ecfRedZ.numOrdem	= bf.varint
	reg->ecfRedZ.valorFinal	= bf.vardbl
	reg->ecfRedZ.valorBruto	= bf.vardbl

	reg->ecfRedZ.numIni		= 2^20
	reg->ecfRedZ.numFim		= -1
	reg->ecfRedZ.itemAnalListHead = null
	reg->ecfRedZ.itemAnalListTail = null

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocECFItem(bf as bfile, reg as TRegistro ptr, documentoPai as TDocECF ptr) as Boolean

	bf.char1		'pular |

	reg->itemECF.documentoPai	= documentoPai
   
	documentoPai->nroItens 		+= 1

	reg->itemECF.numItem		= documentoPai->nroItens
	reg->itemECF.itemId			= bf.varchar
	reg->itemECF.qtd			= bf.vardbl
	reg->itemECF.qtdCancelada	= bf.vardbl
	reg->itemECF.unidade		= bf.varchar
	reg->itemECF.valor			= bf.vardbl
	reg->itemECF.cstICMS		= bf.varint
	reg->itemECF.cfop			= bf.varint
	reg->itemECF.aliqICMS		= bf.vardbl
	reg->itemECF.PIS			= bf.vardbl
	reg->itemECF.COFINS			= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocECFItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= documentoPai
   
	reg->itemAnal.cst		= bf.varint
	reg->itemAnal.cfop		= bf.varint
	reg->itemAnal.aliq		= bf.vardbl
	reg->itemAnal.valorOp	= bf.vardbl
	reg->itemAnal.bc		= bf.vardbl
	reg->itemAnal.ICMS		= bf.vardbl
	bf.varchar					'' pular c�digo de observa��o

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function Efd.lerRegDocSAT(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->sat.operacao		= SAIDA
	reg->sat.modelo			= valint(bf.varchar)
	reg->sat.situacao		= bf.int2
	bf.char1		'pular |
	reg->sat.numero			= bf.varint
	reg->sat.dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->sat.valorTotal		= bf.vardbl
	reg->sat.PIS			= bf.vardbl
	reg->sat.COFINS			= bf.vardbl
	reg->sat.cpfCnpjAdquirente = bf.varchar
	reg->sat.serieEquip		= bf.varchar
	reg->sat.chave 			= bf.varchar
	reg->sat.descontos		= bf.vardbl
	reg->sat.valorMerc 		= bf.vardbl
	reg->sat.despesasAcess	= bf.vardbl
	reg->sat.icms			= bf.vardbl
	reg->sat.pisST			= bf.vardbl
	reg->sat.cofinsST		= bf.vardbl
	reg->sat.nroItens		= 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocSATItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= documentoPai
   
	reg->itemAnal.cst		= bf.varint
	reg->itemAnal.cfop		= bf.varint
	reg->itemAnal.aliq		= bf.vardbl
	reg->itemAnal.valorOp	= bf.vardbl
	reg->itemAnal.bc		= bf.vardbl
	reg->itemAnal.ICMS		= bf.vardbl
	bf.varchar					'' pular c�digo de observa��o

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function


''''''''
function Efd.lerRegDocNFSCT(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->nf.operacao		= bf.int1
	bf.char1		'pular |
	reg->nf.emitente		= bf.int1
	bf.char1		'pular |
	reg->nf.idParticipante	= bf.varchar
	reg->nf.modelo			= bf.int2
	bf.char1		'pular |
	reg->nf.situacao		= bf.int2
	bf.char1		'pular |
	reg->nf.serie			= bf.varchar
	reg->nf.subserie		= bf.varchar
	reg->nf.numero			= bf.varint
	reg->nf.dataEmi			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.valorTotal		= bf.vardbl
	reg->nf.valorDesconto	= bf.vardbl
	bf.vardbl		'pular valorServico
	bf.vardbl 		'pular valorServicoNT
	bf.vardbl 		'pular reg->nf.valorTerceiro
	bf.vardbl 		'pular reg->nf.valorDesp
	reg->nf.bcICMS			= bf.vardbl
	reg->nf.ICMS			= bf.vardbl
	bf.varchar		'pular cod_inf
	reg->nf.PIS				= bf.vardbl
	reg->nf.COFINS			= bf.vardbl
	bf.varchar		'pular cod_cta
	bf.varint		'pular tp_assinante
	reg->nf.nroItens		= 0

	reg->nf.itemAnalListHead = null
	reg->nf.itemAnalListTail = null
	reg->nf.itemAnalCnt = 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocNFSCTItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= documentoPai
   
	reg->itemAnal.cst		= bf.varint
	reg->itemAnal.cfop		= bf.varint
	reg->itemAnal.aliq		= bf.vardbl
	reg->itemAnal.valorOp	= bf.vardbl
	reg->itemAnal.bc		= bf.vardbl
	reg->itemAnal.ICMS		= bf.vardbl
	bf.vardbl		'pular VL_BC_ICMS_UF
	bf.vardbl		'pular VL_ICMS_UF
	reg->itemAnal.redBC		= bf.vardbl
	bf.varchar		'pular COD_OBS

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function Efd.lerRegDocNFElet(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->nf.operacao		= bf.int1
	bf.char1		'pular |
	reg->nf.emitente		= bf.int1
	bf.char1		'pular |
	reg->nf.idParticipante	= bf.varchar
	reg->nf.modelo			= bf.int2
	bf.char1		'pular |
	reg->nf.situacao		= bf.int2
	bf.char1		'pular |
	reg->nf.serie			= bf.varchar
	reg->nf.subserie		= bf.varchar
	bf.varchar		'pular cod_cons
	reg->nf.numero			= bf.varint
	reg->nf.dataEmi			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.dataEntSaida	= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->nf.valorTotal		= bf.vardbl
	reg->nf.valorDesconto	= bf.vardbl
	bf.varchar		'pular vl_forn
	bf.varchar 		'pular vl_serv_nt
	bf.varchar		'pular vl_terc
	bf.varchar		'pular vl_da
	reg->nf.bcICMS			= bf.vardbl
	reg->nf.ICMS			= bf.vardbl
	reg->nf.bcICMSST		= bf.vardbl
	reg->nf.ICMSST			= bf.vardbl
	bf.varchar		'pular cod_inf
	reg->nf.PIS				= bf.vardbl
	reg->nf.COFINS			= bf.vardbl
	bf.varchar		'pular tp_ligacao
	bf.varchar		'pular cod_grupo_tensao
	if regMestre->mestre.versaoLayout >= 014 then
		reg->nf.chave		= bf.varchar		
		bf.varchar		'pular fin_doce
		bf.varchar		'pular chv_doce_ref
		bf.varchar		'pular ind_dest
		bf.varchar		'pular cod_mun_dest
		bf.varchar		'pular cod_cta
	end if
	reg->nf.nroItens		= 0

	reg->nf.itemAnalListHead = null
	reg->nf.itemAnalListTail = null
	reg->nf.itemAnalCnt = 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegDocNFEletItemAnal(bf as bfile, reg as TRegistro ptr, documentoPai as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemAnal.documentoPai	= documentoPai
   
	reg->itemAnal.cst		= bf.varint
	reg->itemAnal.cfop		= bf.varint
	reg->itemAnal.aliq		= bf.vardbl
	reg->itemAnal.valorOp	= bf.vardbl
	reg->itemAnal.bc		= bf.vardbl
	reg->itemAnal.ICMS		= bf.vardbl
	reg->itemAnal.bcST		= bf.vardbl
	reg->itemAnal.ICMSST	= bf.vardbl
	reg->itemAnal.redBC		= bf.vardbl
	bf.varchar		'pular COD_OBS
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if
	
	function = true

end function

''''''''
function Efd.lerRegItemId(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->itemId.id			  	= bf.varchar
	reg->itemId.descricao	  	= bf.varchar
	reg->itemId.codBarra		= bf.varchar
	reg->itemId.codAnterior	  	= bf.varchar
	reg->itemId.unidInventario 	= bf.varchar
	reg->itemId.tipoItem		= bf.varint
	reg->itemId.ncm			  	= bf.varint
	reg->itemId.exIPI		  	= bf.varchar
	reg->itemId.codGenero	  	= bf.varint
	reg->itemId.codServico	  	= bf.varchar
	reg->itemId.aliqICMSInt	  	= bf.vardbl
	'CEST s� � obrigat�rio a partir de 2017
	if bf.peek1 <> 13 and bf.peek1 <> 10 then 
	  reg->itemId.CEST		  	= bf.varint
	end if

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegBemCiap(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->bemCiap.id			  	= bf.varchar
	reg->bemCiap.tipoMerc		= bf.varint
	reg->bemCiap.descricao	  	= bf.varchar
	reg->bemCiap.principal		= bf.varchar
	reg->bemCiap.codAnal	  	= bf.varchar
	reg->bemCiap.parcelas		= bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegInfoCompl(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->infoCompl.id				= bf.varchar
	reg->infoCompl.descricao	  	= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegApuIcmsPeriodo(bf as bfile, reg as TRegistro ptr) as Boolean

   bf.char1		'pular |

   reg->apuIcms.dataIni		  = ddMmYyyy2YyyyMmDd(bf.varchar)
   reg->apuIcms.dataFim		  = ddMmYyyy2YyyyMmDd(bf.varchar)

   'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

   function = true

end function

''''''''
function Efd.lerRegApuIcmsProprio(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->apuIcms.totalDebitos			= bf.vardbl
	reg->apuIcms.ajustesDebitos			= bf.vardbl
	reg->apuIcms.totalAjusteDeb			= bf.vardbl
	reg->apuIcms.estornosCredito		= bf.vardbl
	reg->apuIcms.totalCreditos			= bf.vardbl
	reg->apuIcms.ajustesCreditos		= bf.vardbl
	reg->apuIcms.totalAjusteCred		= bf.vardbl
	reg->apuIcms.estornoDebitos			= bf.vardbl
	reg->apuIcms.saldoCredAnterior		= bf.vardbl
	reg->apuIcms.saldoDevedorApurado	= bf.vardbl
	reg->apuIcms.totalDeducoes			= bf.vardbl
	reg->apuIcms.icmsRecolher			= bf.vardbl
	reg->apuIcms.saldoCredTransportar	= bf.vardbl
	reg->apuIcms.debExtraApuracao		= bf.vardbl

	reg->apuIcms.ajustesListHead 		= null
	reg->apuIcms.ajustesListTail 		= null
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegApuIcmsAjuste(bf as bfile, reg as TRegistro ptr, pai as TApuracaoIcmsPeriodo ptr) as Boolean

	bf.char1		'pular |
	
	reg->apuIcmsAjust.codigo 	= bf.varchar
	reg->apuIcmsAjust.descricao = bf.varchar
	reg->apuIcmsAjust.valor 	= bf.vardbl
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegApuIcmsSTPeriodo(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->apuIcmsST.UF		 	 = bf.varchar
	reg->apuIcmsST.dataIni		 = ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->apuIcmsST.dataFim		 = ddMmYyyy2YyyyMmDd(bf.varchar)

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegApuIcmsST(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->apuIcmsST.mov						= bf.varint
	reg->apuIcmsST.saldoCredAnterior		= bf.vardbl
	reg->apuIcmsST.devolMercadorias			= bf.vardbl
	reg->apuIcmsST.totalRessarciment		= bf.vardbl
	reg->apuIcmsST.totalOutrosCred			= bf.vardbl
	reg->apuIcmsST.ajustesCreditos			= bf.vardbl
	reg->apuIcmsST.totalRetencao			= bf.vardbl
	reg->apuIcmsST.totalOutrosDeb			= bf.vardbl
	reg->apuIcmsST.ajustesDebitos			= bf.vardbl
	reg->apuIcmsST.saldoAntesDed			= bf.vardbl
	reg->apuIcmsST.totalDeducoes			= bf.vardbl
	reg->apuIcmsST.icmsRecolher				= bf.vardbl
	reg->apuIcmsST.saldoCredTransportar		= bf.vardbl
	reg->apuIcmsST.debExtraApuracao			= bf.vardbl

	reg->apuIcmsST.ajustesListHead 			= null
	reg->apuIcmsST.ajustesListTail 			= null
	
	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegInventarioTotais(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->invTotais.dataInventario 	 = ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->invTotais.valorTotalEstoque = bf.vardbl
	reg->invTotais.motivoInventario	 = bf.varint

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegInventarioItem(bf as bfile, reg as TRegistro ptr, inventarioPai as TInventarioTotais ptr) as Boolean

	bf.char1		'pular |

	reg->invItem.dataInventario 	= inventarioPai->dataInventario
	reg->invItem.itemId 	 		= bf.varchar
	reg->invItem.unidade 			= bf.varchar
	reg->invItem.qtd	 			= bf.vardbl
	reg->invItem.valorUnitario		= bf.vardbl
	reg->invItem.valorItem			= bf.vardbl
	reg->invItem.indPropriedade		= bf.varint
	reg->invItem.idParticipante		= bf.varchar
	reg->invItem.txtComplementar	= bf.varchar
	reg->invItem.codConta			= bf.varchar
	reg->invItem.valorItemIR		= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegCiapTotal(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->ciapTotal.dataIni 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ciapTotal.dataFim 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ciapTotal.saldoInicialICMS = bf.vardbl
	reg->ciapTotal.parcelasSoma 	= bf.vardbl
	reg->ciapTotal.valorTributExpSoma = bf.vardbl
	reg->ciapTotal.valorTotalSaidas = bf.vardbl
	reg->ciapTotal.indicePercSaidas = bf.vardbl
	reg->ciapTotal.valorIcmsAprop 	= bf.vardbl
	reg->ciapTotal.valorOutrosCred 	= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegCiapItem(bf as bfile, reg as TRegistro ptr, pai as TCiapTotal ptr) as Boolean

	bf.char1		'pular |

	reg->ciapItem.pai				= pai
	reg->ciapItem.bemId 	 		= bf.varchar
	reg->ciapItem.dataMov 			= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->ciapItem.tipoMov 			= bf.varchar
	reg->ciapItem.valorIcms	 		= bf.vardbl
	reg->ciapItem.valorIcmsSt		= bf.vardbl
	reg->ciapItem.valorIcmsFrete	= bf.vardbl
	reg->ciapItem.valorIcmsDifal	= bf.vardbl
	reg->ciapItem.parcela			= bf.varint
	reg->ciapItem.valorParcela		= bf.vardbl
	reg->ciapItem.docCnt			= 0

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegCiapItemDoc(bf as bfile, reg as TRegistro ptr, pai as TCiapItem ptr) as Boolean

	bf.char1		'pular |

	reg->ciapItemDoc.pai			= pai
	reg->ciapItemDoc.indEmi 		= bf.varint
	reg->ciapItemDoc.idParticipante = bf.varchar
	reg->ciapItemDoc.modelo			= bf.varint
	reg->ciapItemDoc.serie			= bf.varchar
	reg->ciapItemDoc.numero			= bf.varint
	reg->ciapItemDoc.chaveNFe		= bf.varchar
	reg->ciapItemDoc.dataEmi		= ddMmYyyy2YyyyMmDd(bf.varchar)
	pai->docCnt += 1

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegEstoquePeriodo(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.char1		'pular |

	reg->estPeriod.dataIni 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)
	reg->estPeriod.dataFim 	 		= ddMmYyyy2YyyyMmDd(bf.varchar)

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegEstoqueItem(bf as bfile, reg as TRegistro ptr, pai as TEstoquePeriodo ptr) as Boolean

	bf.char1		'pular |

	reg->estItem.pai				= pai
	bf.varchar		'pular DT_EST (� a mesma do DT_FIN do K100)
	reg->estItem.itemId 	 		= bf.varchar
	reg->estItem.qtd 				= bf.vardbl
	reg->estItem.tipoEst			= bf.varint
	reg->estItem.idParticipante		= bf.varchar

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
function Efd.lerRegEstoqueOrdemProd(bf as bfile, reg as TRegistro ptr, pai as TEstoquePeriodo ptr) as Boolean

	bf.char1		'pular |

	reg->estOrdem.pai			= pai
	reg->estOrdem.dataIni 	 	= ddMmYyyy2YyyyMmDd(bf.varchar)
	var dtFim = bf.varchar
	reg->estOrdem.dataFim 	 	= iif(len(dtFim) > 0, ddMmYyyy2YyyyMmDd(dtFim), "99991231")
	reg->estOrdem.idOrdem		= bf.varchar
	reg->estOrdem.itemId 	 	= bf.varchar
	reg->estOrdem.qtd 			= bf.vardbl

	'pular \r\n
	if bf.peek1 = 13 then
		bf.char1
	end if
	if bf.peek1 <> 10 then
		onError("Erro: esperado \n, encontrado " & bf.peek1)
	else
		bf.char1
	end if

	function = true

end function

''''''''
private sub Efd.lerAssinatura(bf as bfile)

	'' verificar header
	var header = bf.nchar(len(ASSINATURA_P7K_HEADER))
	if header <> ASSINATURA_P7K_HEADER then
		onError("Erro: header da assinatura P7K n�o reconhecido")
	end if
	
	var lgt = (bf.tamanho - bf.posicao) + 1
	
	redim this.assinaturaP7K_DER(0 to lgt-1)
	
	bf.ler(assinaturaP7K_DER(), lgt)

end sub

''''''''
function Efd.filtrarPorCnpj(cnpj as const zstring ptr) as boolean
	
	for i as integer = 0 to ubound(opcoes.listaCnpj)
		if(*cnpj = opcoes.listaCnpj(i)) then
			return true
		end if
	next
	
	function = false
	
end function

''''''''
function Efd.filtrarPorChave(chave as const zstring ptr) as boolean
	
	for i as integer = 0 to ubound(opcoes.listaChaves)
		if(*chave = opcoes.listaChaves(i)) then
			return true
		end if
	next
	
	function = false
	
end function

''''''''
function Efd.lerRegistro(bf as bfile, reg as TRegistro ptr) as Boolean
	static as zstring * 4+1 tipo
	
	reg->tipo = lerTipo(bf, @tipo)
	reg->linha = nroLinha

	select case as const reg->tipo
	case DOC_NF
		if not lerRegDocNF(bf, reg) then
			return false
		end if
		
		ultimoReg = reg

	case DOC_NF_INFO
		if( ultimoReg <> null ) then
			if not lerRegDocNFInfo(bf, reg, @ultimoReg->nf) then
				return false
			end if
			
			if ultimoReg->nf.infoComplListHead = null then
				ultimoReg->nf.infoComplListHead = @reg->docInfoCompl
			else
				ultimoReg->nf.infoComplListTail->next_ = @reg->docInfoCompl
			end if
			
			ultimoReg->nf.infoComplListTail = @reg->docInfoCompl
			reg->docInfoCompl.next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
		
	case DOC_NF_ITEM
		if( ultimoReg <> null ) then
			if not lerRegDocNFItem(bf, reg, @ultimoReg->nf) then
				return false
			end if
			
			ultimoDocNFItem = @reg->itemNF
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_NF_ANAL
		if( ultimoReg <> null ) then
			if not lerRegDocNFItemAnal(bf, reg, ultimoReg) then
				return false
			end if
			
			if ultimoReg->nf.itemAnalListHead = null then
				ultimoReg->nf.itemAnalListHead = @reg->itemAnal
			else
				ultimoReg->nf.itemAnalListTail->next_ = @reg->itemAnal
			end if
			
			ultimoReg->nf.itemAnalListTail = @reg->itemAnal
			reg->itemAnal.next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
		
	case DOC_NF_DIFAL
		if( ultimoReg <> null ) then
			if not lerRegDocNFDifal(bf, reg, @ultimoReg->nf) then
				return false
			end if
			
			reg->tipo = DESCONHECIDO			'' deletar registro, j� que vamos reusar o registro pai
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
		
	case DOC_NF_ITEM_RESSARC_ST
		if( ultimoDocNFItem <> null ) then
			if not lerRegDocNFItemRessarcSt(bf, reg, ultimoDocNFItem) then
				return false
			end if
			
			if ultimoDocNFItem->itemRessarcStListHead = null then
				ultimoDocNFItem->itemRessarcStListHead = @reg->itemRessarcSt
			else
				ultimoDocNFItem->itemRessarcStListTail->next_ = @reg->itemRessarcSt
			end if
			
			ultimoDocNFItem->itemRessarcStListTail = @reg->itemRessarcSt
			reg->itemRessarcSt.next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_CT
		if not lerRegDocCT(bf, reg) then
			return false
		end if

		ultimoReg = reg

	case DOC_CT_ANAL
		if( ultimoReg <> null ) then
			if not lerRegDocCTItemAnal(bf, reg, ultimoReg) then
				return false
			end if

			if ultimoReg->ct.itemAnalListHead = null then
				ultimoReg->ct.itemAnalListHead = @reg->itemAnal
			else
				ultimoReg->ct.itemAnalListTail->next_ = @reg->itemAnal
			end if
			
			ultimoReg->ct.itemAnalListTail = @reg->itemAnal
			reg->itemAnal.next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
		
	case DOC_CT_DIFAL
		if( ultimoReg <> null ) then
			if not lerRegDocCTDifal(bf, reg, @reg->ct) then
				return false
			end if
			
			reg->tipo = DESCONHECIDO			'' deletar registro, j� que vamos reusar o registro pai
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_ECF
		if( ultimoEquipECF <> null ) then
			if not lerRegDocECF(bf, reg, ultimoEquipECF) then
				return false
			end if

			ultimoReg = reg
			
			if ultimoECFRedZ->ecfRedZ.numIni > reg->ecf.numero then
				ultimoECFRedZ->ecfRedZ.numIni = reg->ecf.numero
			end if

			if ultimoECFRedZ->ecfRedZ.numFim < reg->ecf.numero then
				ultimoECFRedZ->ecfRedZ.numFim = reg->ecf.numero
			end if
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
		
	case ECF_REDUCAO_Z
		if( ultimoEquipECF <> null ) then
			if not lerRegECFReducaoZ(bf, reg, ultimoEquipECF) then
				return false
			end if

			ultimoECFRedZ = reg
		else
			pularLinha(bf)
			ultimoECFRedZ = null
			reg->tipo = DESCONHECIDO
		end if
		
	case DOC_ECF_ITEM
		if( ultimoReg <> null ) then
			if not lerRegDocECFItem(bf, reg, @ultimoReg->ecf) then
				return false
			end if
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_ECF_ANAL
		if( ultimoECFRedZ <> null ) then
			if not lerRegDocECFItemAnal(bf, reg, ultimoECFRedZ) then
				return false
			end if
			
			if ultimoECFRedZ->ecfRedZ.itemAnalListHead = null then
				ultimoECFRedZ->ecfRedZ.itemAnalListHead = @reg->itemAnal
			else
				ultimoECFRedZ->ecfRedZ.itemAnalListTail->next_ = @reg->itemAnal
			end if
			
			ultimoECFRedZ->ecfRedZ.itemAnalListTail = @reg->itemAnal
			reg->itemAnal.next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case EQUIP_ECF
		if not lerRegEquipECF(bf, reg) then
			return false
		end if
		
		ultimoEquipECF = @reg->equipECF

	case DOC_SAT
		if not lerRegDocSAT(bf, reg) then
			return false
		end if

		ultimoReg = reg

	case DOC_SAT_ANAL
		if( ultimoReg <> null ) then
			if not lerRegDocSATItemAnal(bf, reg, ultimoReg) then
				return false
			end if

			if ultimoReg->sat.itemAnalListHead = null then
				ultimoReg->sat.itemAnalListHead = @reg->itemAnal
			else
				ultimoReg->sat.itemAnalListTail->next_ = @reg->itemAnal
			end if
			
			ultimoReg->sat.itemAnalListTail = @reg->itemAnal
			reg->itemAnal.next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if

	case DOC_NFSCT
		if not lerRegDocNFSCT(bf, reg) then
			return false
		end if

		ultimoReg = reg

	case DOC_NFSCT_ANAL
		if( ultimoReg <> null ) then
			if not lerRegDocNFSCTItemAnal(bf, reg, ultimoReg) then
				return false
			end if

			if ultimoReg->nf.itemAnalListHead = null then
				ultimoReg->nf.itemAnalListHead = @reg->itemAnal
			else
				ultimoReg->nf.itemAnalListTail->next_ = @reg->itemAnal
			end if
			
			ultimoReg->nf.itemAnalListTail = @reg->itemAnal
			reg->itemAnal.next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
	
	case DOC_NF_ELETRIC
		if not lerRegDocNFElet(bf, reg) then
			return false
		end if

		ultimoReg = reg

	case DOC_NF_ELETRIC_ANAL
		if( ultimoReg <> null ) then
			if not lerRegDocNFEletItemAnal(bf, reg, ultimoReg) then
				return false
			end if

			if ultimoReg->nf.itemAnalListHead = null then
				ultimoReg->nf.itemAnalListHead = @reg->itemAnal
			else
				ultimoReg->nf.itemAnalListTail->next_ = @reg->itemAnal
			end if
			
			ultimoReg->nf.itemAnalListTail = @reg->itemAnal
			reg->itemAnal.next_ = null
		else
			pularLinha(bf)
			reg->tipo = DESCONHECIDO
		end if
	
	case ITEM_ID
		if not lerRegItemId(bf, reg) then
			return false
		end if

		'adicionar ao dicion�rio
		if itemIdDict->lookup(reg->itemId.id) = null then
			itemIdDict->add(reg->itemId.id, @reg->itemId)
		end if

	case BEM_CIAP
		if not lerRegBemCiap(bf, reg) then
			return false
		end if

		'adicionar ao dicion�rio
		if bemCiapDict->lookup(reg->bemCiap.id) = null then
			bemCiapDict->add(reg->bemCiap.id, @reg->bemCiap)
		end if

	case INFO_COMPL
		if not lerRegInfoCompl(bf, reg) then
			return false
		end if

		'adicionar ao dicion�rio
		if infoComplDict->lookup(reg->infoCompl.id) = null then
			infoComplDict->add(reg->infoCompl.id, @reg->infoCompl)
		end if

	case PARTICIPANTE
		if not lerRegParticipante(bf, reg) then
			return false
		end if

		'adicionar ao dicion�rio
		if participanteDict->lookup(reg->part.id) = null then
			participanteDict->add(reg->part.id, @reg->part)
		end if

	case APURACAO_ICMS_PERIODO
		if not lerRegApuIcmsPeriodo(bf, reg) then
			return false
		end if

		ultimoReg = reg
		
	case APURACAO_ICMS_PROPRIO
		if not lerRegApuIcmsProprio(bf, ultimoReg) then
			return false
		end if
		
		reg->tipo = DESCONHECIDO			'' deletar registro, j� que vamos reusar o registro pai

	case APURACAO_ICMS_AJUSTE
		'' nota: como apuIcms e apuIcmsST estendem a mesma classe, pode-se acessar os campos comuns de qualquer classe filha
		if not lerRegApuIcmsAjuste(bf, reg, @ultimoReg->apuIcms) then
			return false
		end if

		if ultimoReg->apuIcms.ajustesListHead = null then
			ultimoReg->apuIcms.ajustesListHead = @reg->apuIcmsAjust
		else
			ultimoReg->apuIcms.ajustesListTail->next_ = @reg->apuIcmsAjust
		end if

		ultimoReg->apuIcms.ajustesListTail = @reg->apuIcmsAjust
		reg->apuIcmsAjust.next_ = null

	case APURACAO_ICMS_ST_PERIODO
		if not lerRegApuIcmsSTPeriodo(bf, reg) then
			return false
		end if

		ultimoReg = reg
		
	case APURACAO_ICMS_ST
		if not lerRegApuIcmsST(bf, ultimoReg) then
			return false
		end if
		
		reg->tipo = DESCONHECIDO			'' deletar registro, j� que vamos reusar o registro pai

	case INVENTARIO_TOTAIS
		if not lerRegInventarioTotais(bf, reg) then
			return false
		end if
		
		ultimoInventario = @reg->invTotais
	
	case INVENTARIO_ITEM
		if not lerRegInventarioItem(bf, reg, ultimoInventario) then
			return false
		end if
	
	case CIAP_TOTAL
		if not lerRegCiapTotal(bf, reg) then
			return false
		end if
		
		ultimoCiap = @reg->ciapTotal
	
	case CIAP_ITEM
		if not lerRegCiapItem(bf, reg, ultimoCiap) then
			return false
		end if
	
		ultimoCiapItem = @reg->ciapItem

	case CIAP_ITEM_DOC
		if not lerRegCiapItemDoc(bf, reg, ultimoCiapItem) then
			return false
		end if

	case ESTOQUE_PERIODO
		if not lerRegEstoquePeriodo(bf, reg) then
			return false
		end if
		
		ultimoEstoque = @reg->estPeriod
	
	case ESTOQUE_ITEM
		if not lerRegEstoqueItem(bf, reg, ultimoEstoque) then
			return false
		end if
		
	case ESTOQUE_ORDEM_PROD
		if not lerRegEstoqueOrdemProd(bf, reg, ultimoEstoque) then
			return false
		end if
	
	case MESTRE
		if not lerRegMestre(bf, reg) then
			return false
		end if
		
		regMestre = reg

	case FIM_DO_ARQUIVO
		pularLinha(bf)
		
		lerAssinatura(bf)
	
	case LUA_CUSTOM
		
		var luaFunc = cast(customLuaCb ptr, customLuaCbDict->lookup(tipo))->reader
		
		if luaFunc <> null then
			lua_getglobal(lua, luaFunc)
			lua_pushlightuserdata(lua, @bf)
			lua_newtable(lua)
			reg->lua.table = luaL_ref(lua, LUA_REGISTRYINDEX)
			lua_rawgeti(lua, LUA_REGISTRYINDEX, reg->lua.table)
			lua_call(lua, 2, 1)

			reg->tipo = LUA_CUSTOM
			reg->lua.tipo = tipo
			
		end if
	
	case else
		pularLinha(bf)
	end select

	function = true

end function

''''''''
private function situacaoSintegra2SituacaoEfd(sit as byte) as TipoSituacao
	select case sit
	case asc("N")
		return REGULAR
	case asc("S")
		return CANCELADO
	case asc("E")
		return EXTEMPORANEO
	case asc("X")
		return CANCELADO_EXT
	case asc("2")
		return DENEGADO
	case asc("4")
		return INUTILIZADO
	case else
		return REGULAR
	end select

end function

''''''''
private function lerRegSintegraDocumento(bf as bfile, reg as TRegistro ptr) as Boolean

	reg->docSint.cnpj 		= bf.nchar(14)
	reg->docSint.ie 		= bf.nchar(14)
	reg->docSint.dataEmi 	= bf.char8
	reg->docSint.uf 		= UF_SIGLA2COD(bf.char2)
	reg->docSint.modelo 	= bf.int2
	reg->docSint.serie 		= bf.nchar(3)
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("�") then
		bf.char1
		reg->docSint.numero = bf.int9
	else
		reg->docSint.numero = bf.int6
	end if
	reg->docSint.cfop 		= bf.int4
	reg->docSint.operacao 	= iif( bf.char1 = asc("T"), ENTRADA, SAIDA )
	reg->docSint.valorTotal = bf.dbl13_2
	reg->docSint.bcICMS 	= bf.dbl13_2
	reg->docSint.ICMS 		= bf.dbl13_2
	reg->docSint.valorIsento= bf.dbl13_2
	reg->docSint.valorOutras= bf.dbl13_2
	reg->docSint.aliqICMS 	= bf.dbl4_2
	reg->docSint.situacao 	= situacaoSintegra2SituacaoEfd( bf.char1 )

	'' ler chave NF-e no final da linha, se for um sintegra convertido pelo SAFI
	if bf.peek1 <> 13 then
		reg->docSint.chave 	= bf.nchar(44)
	end if

	'pular \r\n
	bf.char1
	bf.char1

	function = true
end function

''''''''
private function lerRegSintegraDocumentoST(bf as bfile, reg as TRegistro ptr) as Boolean

	reg->docSint.cnpj 		= bf.nchar(14)
	reg->docSint.ie 		= bf.nchar(14)
	reg->docSint.dataEmi	= bf.char8
	reg->docSint.uf 		= UF_SIGLA2COD(bf.char2)
	reg->docSint.modelo 	= bf.int2
	reg->docSint.serie 		= bf.nchar(3)
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("�") then
		bf.char1
		reg->docSint.numero = bf.int9
	else
		reg->docSint.numero = bf.int6
	end if
	reg->docSint.cfop 		= bf.int4
	reg->docSint.operacao 	= iif( bf.char1 = asc("T"), ENTRADA, SAIDA )
	reg->docSint.bcICMSST 	= bf.dbl13_2
	reg->docSint.ICMSST 	= bf.dbl13_2
	reg->docSint.despesasAcess = bf.dbl13_2
	reg->docSint.situacao 	= situacaoSintegra2SituacaoEfd( bf.char1 )
	bf.nchar(30)

	'pular \r\n
	bf.char1
	bf.char1

	function = true
end function

''''''''
private function lerRegSintegraDocumentoIPI(bf as bfile, reg as TRegistro ptr) as Boolean

	reg->docSint.cnpj 		= bf.nchar(14)
	reg->docSint.ie 		= bf.nchar(14)
	reg->docSint.dataEmi 	= bf.char8
	reg->docSint.uf 		= UF_SIGLA2COD(bf.char2)
	reg->docSint.serie 		= bf.nchar(3)
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("�") then
		bf.char1
		reg->docSint.numero = bf.int9
	else
		reg->docSint.numero = bf.int6
	end if
	reg->docSint.cfop 		= bf.int4
	reg->docSint.valorTotal = bf.dbl13_2
	reg->docSint.valorIPI 	= bf.dbl13_2
	reg->docSint.valorIsentoIPI = bf.dbl13_2
	reg->docSint.valorOutrasIPI = bf.dbl13_2
	bf.nchar(1+20)

	'pular \r\n
	bf.char1
	bf.char1

	function = true
end function

''''''''
private function lerRegSintegraMercadoria(bf as bfile, reg as TRegistro ptr) as Boolean

	bf.nchar(8+8)
	reg->itemId.id			  	= bf.nchar(14)
	reg->itemId.ncm			  	= vallng(bf.nchar(8))
	reg->itemId.descricao	  	= bf.nchar(53)
	reg->itemId.unidInventario 	= bf.nchar(6)
	reg->itemId.aliqIPI		  	= bf.dbl5_2
	reg->itemId.aliqICMSInt	  	= bf.dbl4_2
	reg->itemId.redBcICMS	  	= bf.dbl5_2
	reg->itemId.bcICMSST	  	= bf.dbl13_2

	'pular \r\n
	bf.char1
	bf.char1

	function = true
end function

''''''''
private function lerRegSintegraDocumentoItem(bf as bfile, reg as TRegistro ptr) as Boolean
	
	reg->docItemSint.cnpj 		= bf.nchar(14)
	bf.nchar(2)
	reg->docItemSint.serie 		= bf.nchar(3)
	'' formato de numero estendido do SAFI?
	if bf.peek1 = asc("�") then
		bf.char1
		reg->docItemSint.numero = bf.int9
	else
		reg->docItemSint.numero = bf.int6
	end if
	reg->docItemSint.cfop 		= bf.int4
	reg->docItemSint.CST 		= bf.nchar(3)
	reg->docItemSint.nroItem	= valint(bf.nchar(3))	
	reg->docItemSint.codMercadoria = bf.nchar(14)
	reg->docItemSint.qtd		= bf.dbl11_3
	reg->docItemSint.valor		= bf.dbl12_2
	reg->docItemSint.desconto	= bf.dbl12_2
	reg->docItemSint.bcICMS		= bf.dbl12_2
	reg->docItemSint.bcICMSST	= bf.dbl12_2
	reg->docItemSint.valorIPI	= bf.dbl12_2
	reg->docItemSint.aliqICMS	= bf.dbl4_2
	
	'pular \r\n
	bf.char1
	bf.char1

	function = true
end function

#define GENSINTEGRAKEY(r) ((r)->cnpj + (r)->serie + str((r)->numero) + str((r)->cfop))
  
''''''''
function Efd.lerRegistroSintegra(bf as bfile, reg as TRegistro ptr) as Boolean

	var tipo = bf.int2

	select case as const tipo
	case SINTEGRA_DOCUMENTO
		reg->tipo = SINTEGRA_DOCUMENTO
		if not lerRegSintegraDocumento(bf, reg) then
			return false
		end if

		'adicionar ao dicion�rio
		reg->docSint.chaveDict = GENSINTEGRAKEY(@reg->docSint)
		var antReg = cast(TRegistro ptr, sintegraDict->lookup(reg->docSint.chaveDict))
		if antReg = null then
			sintegraDict->add(reg->docSint.chaveDict, reg)
		else
			'' para cada al�quota diferente h� um novo registro 50, mas n�s s� queremos os valores totais
			''antReg->docSint.valorTotal	+= reg->docSint.valorTotal
			''antReg->docSint.bcICMS		+= reg->docSint.bcICMS
			''antReg->docSint.ICMS		+= reg->docSint.ICMS
			''antReg->docSint.valorIsento += reg->docSint.valorIsento
			''antReg->docSint.valorOutras += reg->docSint.valorOutras

			reg->tipo = DESCONHECIDO 
		end if

	case SINTEGRA_DOCUMENTO_ST
		reg->tipo = SINTEGRA_DOCUMENTO
		if not lerRegSintegraDocumentoST(bf, reg) then
			return false
		end if

		reg->docSint.chaveDict = GENSINTEGRAKEY(@reg->docSint)
		var antReg = cast(TRegistro ptr, sintegraDict->lookup(reg->docSint.chaveDict))
		'' NOTA: pode existir registro 53 sem o correspondente 50, para quando s� h� ICMS ST, sem destaque ICMS pr�prio
		if antReg = null then
			sintegraDict->add(reg->docSint.chaveDict, reg)
		else
			''antReg->docSint.bcICMSST		+= reg->docSint.bcICMSST
			''antReg->docSint.ICMSST			+= reg->docSint.ICMSST
			''antReg->docSint.despesasAcess	+= reg->docSint.despesasAcess
			reg->tipo = DESCONHECIDO
		end if
	  
	case SINTEGRA_DOCUMENTO_IPI
		reg->tipo = SINTEGRA_DOCUMENTO
		if not lerRegSintegraDocumentoIPI(bf, reg) then
			return false
		end if

		reg->docSint.chaveDict = GENSINTEGRAKEY(@reg->docSint)
		var antReg = cast(TRegistro ptr, sintegraDict->lookup(reg->docSint.chaveDict))
		if antReg = null then
			onError("ERRO: Sintegra 53 sem 50: " & reg->docSint.chaveDict)
		else
			antReg->docSint.valorIPI		= reg->docSint.valorIPI
			antReg->docSint.valorIsentoIPI	= reg->docSint.valorIsentoIPI
			antReg->docSint.valorOutrasIPI	= reg->docSint.valorOutrasIPI
		end if

		reg->tipo = DESCONHECIDO 
		
	case SINTEGRA_DOCUMENTO_ITEM
		reg->tipo = SINTEGRA_DOCUMENTO_ITEM
		if not lerRegSintegraDocumentoItem(bf, reg) then
			return false
		end if

		var chaveDict = GENSINTEGRAKEY(@reg->docItemSint)
		var doc = cast(TRegistro ptr, sintegraDict->lookup(chaveDict))
		if doc = null then
			onError("ERRO: Sintegra 54 sem 50: " & chaveDict)
		end if
		
		reg->docItemSint.doc = @(doc->docSint)
		
	case SINTEGRA_MERCADORIA
		reg->tipo = ITEM_ID
		if not lerRegSintegraMercadoria(bf, reg) then
			return false
		end if

		'adicionar ao dicion�rio
		if itemIdDict->lookup(reg->itemId.id) = null then
			itemIdDict->add(reg->itemId.id, @reg->itemId)
		end if
		
	case else
		pularLinha(bf)
		reg->tipo = DESCONHECIDO
	end select

	function = true

end function

''''''''
function Efd.carregarSintegra(bf as bfile) as Boolean
	
	var fsize = bf.tamanho
	
	dim as TRegistro ptr tail = null
	var nroLinha = 0

	try
		do while bf.temProximo()		 
			var reg = new TRegistro
			
			nroLinha += 1

			if lerRegistroSintegra( bf, reg ) then 
				onProgress(null, bf.posicao / fsize)
				
				if reg->tipo <> DESCONHECIDO then
					if tail = null then
					   regListHead = reg
					   tail = reg
					else
					   tail->next_ = reg
					   tail = reg
					end if

					nroRegs += 1
				else
					delete reg
				end if
			 
			else
				exit do
			end if
		loop
	catch
		onError(!"\r\nErro ao carregar o registro da linha (" & nroLinha & !") do arquivo\r\n")
	endtry
	   
	function = true

end function

''''''''
sub Efd.adicionarMestre(reg as TMestre ptr)

	'' (versao, original, dataIni, dataFim, nome, cnpj, uf, ie)
	db_mestreInsertStmt->reset()
	db_mestreInsertStmt->bind(1, reg->versaoLayout)
	db_mestreInsertStmt->bind(2, cint(reg->original))
	db_mestreInsertStmt->bind(3, reg->dataIni)
	db_mestreInsertStmt->bind(4, reg->dataFim)
	db_mestreInsertStmt->bind(5, reg->nome)
	db_mestreInsertStmt->bind(6, reg->cnpj)
	db_mestreInsertStmt->bind(7, reg->uf)
	db_mestreInsertStmt->bind(8, reg->ie)
	
	if not db->execNonQuery(db_mestreInsertStmt) then
		onError("Erro ao inserir registro na EFD_Mestre: " & *db->getErrorMsg())
	end if

end sub

''''''''
sub Efd.adicionarDocEscriturado(doc as TDocDF ptr)
	
	select case as const doc->situacao
	case REGULAR, EXTEMPORANEO
		var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipante) )
		
		var uf = iif(part->municip >= 1100000 and part->municip <= 5399999, part->municip \ 100000, 99)
		
		'' adicionar ao db
		if doc->operacao = ENTRADA then
			'' (periodo, cnpjEmit, ufEmit, serie, numero, modelo, chave, dataEmit, valorOp, IE)
			db_LREInsertStmt->reset()
			db_LREInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
			if len(part->cpf) > 0 then 
				db_LREInsertStmt->bind(2, part->cpf)
			else
				db_LREInsertStmt->bind(2, part->cnpj)
			end if
			db_LREInsertStmt->bind(3, uf)
			db_LREInsertStmt->bind(4, doc->serie)
			db_LREInsertStmt->bind(5, doc->numero)
			db_LREInsertStmt->bind(6, doc->modelo)
			db_LREInsertStmt->bind(7, doc->chave)
			db_LREInsertStmt->bind(8, doc->dataEmi)
			db_LREInsertStmt->bind(9, doc->valorTotal)
			if len(part->ie) > 0 then
				db_LREInsertStmt->bind(10, trim(part->ie))
			else
				db_LREInsertStmt->bindNull(10)
			end if
			
			if not db->execNonQuery(db_LREInsertStmt) then
				onError("Erro ao inserir registro na EFD_LRE: " & *db->getErrorMsg())
			end if
		else
			'' (periodo, cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp, IE)
			db_LRSInsertStmt->reset()
			db_LRSInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
			if len(part->cpf) > 0 then 
				db_LRSInsertStmt->bind(2, part->cpf)
			else
				db_LRSInsertStmt->bind(2, part->cnpj)
			end if
			db_LRSInsertStmt->bind(3, uf)
			db_LRSInsertStmt->bind(4, doc->serie)
			db_LRSInsertStmt->bind(5, doc->numero)
			db_LRSInsertStmt->bind(6, doc->modelo)
			db_LRSInsertStmt->bind(7, doc->chave)
			db_LRSInsertStmt->bind(8, doc->dataEmi)
			db_LRSInsertStmt->bind(9, doc->valorTotal)
			if len(part->ie) > 0 then
				db_LRSInsertStmt->bind(10, trim(part->ie))
			else
				db_LRSInsertStmt->bindNull(10)
			end if
		
			if not db->execNonQuery(db_LRSInsertStmt) then
				onError("Erro ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
			end if
		end if
	
	case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
		'' !!!TODO!!! inserir em outra tabela para fazermos an�lises posteriores
	
	case else
		'' !!!TODO!!! como tratar outras situa��es? os dados v�m completos?
	end select
	
end sub

''''''''
sub Efd.adicionarDocEscriturado(doc as TDocECF ptr)
	
	select case as const doc->situacao
	case REGULAR, EXTEMPORANEO
	
		'' s� existe de sa�da para ECF
		if doc->operacao = SAIDA then
			'' (periodo, cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp)
			db_LRSInsertStmt->reset()
			db_LRSInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
			db_LRSInsertStmt->bind(2, doc->cpfCnpjAdquirente)
			db_LRSInsertStmt->bind(3, 35)
			db_LRSInsertStmt->bind(4, 0)
			db_LRSInsertStmt->bind(5, doc->numero)
			db_LRSInsertStmt->bind(6, doc->modelo)
			db_LRSInsertStmt->bind(7, doc->chave)
			db_LRSInsertStmt->bind(8, doc->dataEmi)
			db_LRSInsertStmt->bind(9, doc->valorTotal)
			db_dfeSaidaInsertStmt->bindNull(10)
		
			if not db->execNonQuery(db_LRSInsertStmt) then
				onError("Erro ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
			end if
		end if
	
	case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
		'' !!!TODO!!! inserir em outra tabela para fazermos an�lises posteriores
	
	case else
		'' !!!TODO!!! como tratar outras situa��es? os dados v�m completos?
	end select
	
end sub

''''''''
sub Efd.adicionarDocEscriturado(doc as TDocSAT ptr)
	
	select case as const doc->situacao
	case REGULAR, EXTEMPORANEO
	
		'' s� existe de sa�da para SAT
		if doc->operacao = SAIDA then
			'' (periodo, cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp)
			db_LRSInsertStmt->reset()
			db_LRSInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
			db_LRSInsertStmt->bind(2, 0) '' n�o � poss�vel usar doc->cpfCnpjAdquirente, porque relat�rio do BO vem sem essa info
			db_LRSInsertStmt->bind(3, 35)
			db_LRSInsertStmt->bind(4, 0)
			db_LRSInsertStmt->bind(5, doc->numero)
			db_LRSInsertStmt->bind(6, doc->modelo)
			db_LRSInsertStmt->bind(7, doc->chave)
			db_LRSInsertStmt->bind(8, doc->dataEmi)
			db_LRSInsertStmt->bind(9, doc->valorTotal)
			db_dfeSaidaInsertStmt->bindNull(10)
		
			if not db->execNonQuery(db_LRSInsertStmt) then
				onError("Erro ao inserir registro na EFD_LRS: " & *db->getErrorMsg())
			end if
		end if
	
	case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
		'' !!!TODO!!! inserir em outra tabela para fazermos an�lises posteriores
	
	case else
		'' !!!TODO!!! como tratar outras situa��es? os dados v�m completos?
	end select
	
end sub

''''''''
sub Efd.adicionarItemNFEscriturado(item as TDocNFItem ptr)
	
	var doc = item->documentoPai
	select case as const doc->situacao
	case REGULAR, EXTEMPORANEO
		var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipante) )
		
		var uf = iif(part->municip >= 1100000 and part->municip <= 5399999, part->municip \ 100000, 99)

		'' (periodo, cnpjEmit, ufEmit, serie, numero, modelo, numItem, cst, cst_origem, cst_tribut, cfop, qtd, valorProd, valorDesc, bc, aliq, icms, bcIcmsST, aliqIcmsST, icmsST, itemId)
		db_itensNfLRInsertStmt->reset()
		db_itensNfLRInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
		db_itensNfLRInsertStmt->bind(2, iif(len(part->cpf) > 0, part->cpf, part->cnpj))
		db_itensNfLRInsertStmt->bind(3, uf)
		db_itensNfLRInsertStmt->bind(4, doc->serie)
		db_itensNfLRInsertStmt->bind(5, doc->numero)
		db_itensNfLRInsertStmt->bind(6, doc->modelo)
		db_itensNfLRInsertStmt->bind(7, item->numItem)
		db_itensNfLRInsertStmt->bind(8, item->cstIcms)
		db_itensNfLRInsertStmt->bind(9, item->cstIcms \ 100)
		db_itensNfLRInsertStmt->bind(10, item->cstIcms mod 100)
		db_itensNfLRInsertStmt->bind(11, item->cfop)
		db_itensNfLRInsertStmt->bind(12, item->qtd)
		db_itensNfLRInsertStmt->bind(13, item->valor)
		db_itensNfLRInsertStmt->bind(14, item->desconto)
		db_itensNfLRInsertStmt->bind(15, item->bcICMS)
		db_itensNfLRInsertStmt->bind(16, item->aliqICMS)
		db_itensNfLRInsertStmt->bind(17, item->icms)
		db_itensNfLRInsertStmt->bind(18, item->bcICMSST)
		db_itensNfLRInsertStmt->bind(19, item->aliqICMSST)
		db_itensNfLRInsertStmt->bind(20, item->icmsST)
		if opcoes.manterDb then
			db_itensNfLRInsertStmt->bind(21, item->itemId)
		else
			db_itensNfLRInsertStmt->bind(21, null)
		end if
		
		if not db->execNonQuery(db_itensNfLRInsertStmt) then
			onError("Erro ao inserir registro na EFD_Itens: " & *db->getErrorMsg())
		end if
	end select
	
end sub

''''''''
sub Efd.adicionarRessarcStEscriturado(doc as TDocNFItemRessarcSt ptr)

	var docPai = doc->documentoPai
	var docAvo = doc->documentoPai->documentoPai
	
	var part = cast( TParticipante ptr, participanteDict->lookup(docAvo->idParticipante) )
	var uf = iif(part->municip >= 1100000 and part->municip <= 5399999, part->municip \ 100000, 99)
	
	var partUlt = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipanteUlt) )
	var ufUlt = iif(partUlt->municip >= 1100000 and partUlt->municip <= 5399999, partUlt->municip \ 100000, 99)
	
	'' (periodo, cnpjEmit, ufEmit, serie, numero, modelo, nroItem, cnpjUlt, ufUlt, serieUlt, numeroUlt, modeloUlt, chaveUlt, dataUlt, valorUlt, bcSTUlt, qtdUlt, nroItemUlt)
	db_ressarcStItensNfLRSInsertStmt->reset()
	db_ressarcStItensNfLRSInsertStmt->bind(1, valint(regMestre->mestre.dataIni))
	db_ressarcStItensNfLRSInsertStmt->bind(2, iif(len(part->cpf) > 0, part->cpf, part->cnpj))
	db_ressarcStItensNfLRSInsertStmt->bind(3, uf)
	db_ressarcStItensNfLRSInsertStmt->bind(4, docAvo->serie)
	db_ressarcStItensNfLRSInsertStmt->bind(5, docAvo->numero)
	db_ressarcStItensNfLRSInsertStmt->bind(6, docAvo->modelo)
	db_ressarcStItensNfLRSInsertStmt->bind(7, docPai->numItem)
	db_ressarcStItensNfLRSInsertStmt->bind(8, partUlt->cnpj)
	db_ressarcStItensNfLRSInsertStmt->bind(9, ufUlt)
	db_ressarcStItensNfLRSInsertStmt->bind(10, doc->serieUlt)
	db_ressarcStItensNfLRSInsertStmt->bind(11, doc->numeroUlt)
	db_ressarcStItensNfLRSInsertStmt->bind(12, doc->modeloUlt)
	if len(doc->chaveNFeUlt) > 0 then
		db_ressarcStItensNfLRSInsertStmt->bind(13, doc->chaveNFeUlt)
	else
		db_ressarcStItensNfLRSInsertStmt->bindNull(13)
	end if
	db_ressarcStItensNfLRSInsertStmt->bind(14, doc->dataUlt)
	db_ressarcStItensNfLRSInsertStmt->bind(15, doc->valorUlt)
	db_ressarcStItensNfLRSInsertStmt->bind(16, doc->valorBcST)
	db_ressarcStItensNfLRSInsertStmt->bind(17, doc->qtdUlt)
	if doc->numItemNFeUlt > 0 then
		db_ressarcStItensNfLRSInsertStmt->bind(18, doc->numItemNFeUlt)
	else
		db_ressarcStItensNfLRSInsertStmt->bindNull(18)
	end if

	if not db->execNonQuery(db_ressarcStItensNfLRSInsertStmt) then
		onError("Erro ao inserir registro na EFD_Ressarc_Itens: " & *db->getErrorMsg())
	end if
	
end sub

''''''''
sub Efd.adicionarItemEscriturado(item as TItemId ptr)

	'' (id, descricao, ncm, cest, aliqInt)
	db_itensIdInsertStmt->reset()
	db_itensIdInsertStmt->bind(1, item->id)
	db_itensIdInsertStmt->bind(2, item->descricao)
	db_itensIdInsertStmt->bind(3, item->ncm)
	db_itensIdInsertStmt->bind(4, item->CEST)
	db_itensIdInsertStmt->bind(5, item->aliqICMSInt)
	
	if not db->execNonQuery(db_itensIdInsertStmt) then
		onError("Erro ao inserir registro na EFD_ItensId: " & *db->getErrorMsg())
	end if

end sub

''''''''
sub Efd.adicionarAnalEscriturado(anal as TDocItemAnal ptr)

	var doc = @anal->documentoPai->nf
	var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipante) )
	
	var uf = iif(part->municip >= 1100000 and part->municip <= 5399999, part->municip \ 100000, 99)

	'' (operacao, periodo, cnpj, uf, serie, numero, modelo, numReg, cst, cst_origem, cst_tribut, cfop, aliq, valorOp, bc, icms, bcIcmsST, icmsST, redBC, ipi)
	db_analInsertStmt->reset()
	db_analInsertStmt->bind(1, doc->operacao)
	db_analInsertStmt->bind(2, valint(regMestre->mestre.dataIni))
	db_analInsertStmt->bind(3, iif(len(part->cpf) > 0, part->cpf, part->cnpj))
	db_analInsertStmt->bind(4, uf)
	db_analInsertStmt->bind(5, doc->serie)
	db_analInsertStmt->bind(6, doc->numero)
	db_analInsertStmt->bind(7, doc->modelo)
	db_analInsertStmt->bind(8, anal->num)
	db_analInsertStmt->bind(9, anal->cst)
	db_analInsertStmt->bind(10, anal->cst \ 100)
	db_analInsertStmt->bind(11, anal->cst mod 100)
	db_analInsertStmt->bind(12, anal->cfop)
	db_analInsertStmt->bind(13, anal->aliq)
	db_analInsertStmt->bind(14, anal->valorOp)
	db_analInsertStmt->bind(15, anal->bc)
	db_analInsertStmt->bind(16, anal->ICMS)
	db_analInsertStmt->bind(17, anal->bcST)
	db_analInsertStmt->bind(18, anal->ICMSST)
	db_analInsertStmt->bind(19, anal->redBC)
	db_analInsertStmt->bind(20, anal->IPI)
	
	if not db->execNonQuery(db_analInsertStmt) then
		onError("Erro ao inserir registro na EDF_Anal: " & *db->getErrorMsg())
	end if

end sub

''''''''
sub Efd.addRegistroAoDB(reg as TRegistro ptr)

	select case as const reg->tipo
	case DOC_NF
		adicionarDocEscriturado(@reg->nf)
	case DOC_NF_ITEM
		adicionarItemNFEscriturado(@reg->itemNF)
	case DOC_NF_ANAL
		adicionarAnalEscriturado(@reg->itemAnal)
	case DOC_CT
		adicionarDocEscriturado(@reg->ct)
	case DOC_ECF
		adicionarDocEscriturado(@reg->ecf)
	case DOC_SAT
		adicionarDocEscriturado(@reg->sat)
	case DOC_NF_ITEM_RESSARC_ST
		adicionarRessarcStEscriturado(@reg->itemRessarcSt)
	case ITEM_ID
		if opcoes.manterDb then
			adicionarItemEscriturado(@reg->itemId)
		end if
	case MESTRE
		adicionarMestre(@reg->mestre)
	end select
	
end sub

''''''''
private function yyyyMmDd2Days(d as const zstring ptr) as uinteger

	if d = null then
		d = @"19000101"
	end if
	
	var days = (cuint(d[0] - asc("0")) * 1000 + _
				cuint(d[1] - asc("0")) * 0100 + _
				cuint(d[2] - asc("0")) * 0010 + _
				cuint(d[3] - asc("0")) * 0001) * (31*12)
	
	days = days + _
			   ((cuint(d[4] - asc("0")) * 10 + _
				 cuint(d[5] - asc("0")) * 01) - 1) * 31

	days = days + _
			   (cuint(d[6] - asc("0")) * 10 + _
				cuint(d[7] - asc("0")) * 01) 
				
	function = days

end function

''''''''
private function mergeLists(pSrc1 as TRegistro ptr, pSrc2 as TRegistro ptr) as TRegistro ptr
	dim as TRegistro ptr pDst = NULL
	dim as TRegistro ptr ptr ppDst = @pDst
    if pSrc1 = NULL then
        return pSrc2
	end if
    if pSrc2 = NULL then
        return pSrc1
	end if
    
	dim as zstring ptr dReg = null
	dim as longint nro = 0

	do while true
		select case as const pSrc1->tipo
		case DOC_NF
			dReg = @pSrc1->nf.dataEntSaida
			nro = pSrc1->nf.numero
		case DOC_CT
			dReg = @pSrc1->ct.dataEntSaida
			nro = pSrc1->ct.numero
		case DOC_NF_ITEM
			dReg = @pSrc1->itemNF.documentoPai->dataEntSaida
			nro = pSrc1->itemNF.documentoPai->numero
		case ECF_REDUCAO_Z
			dReg = @pSrc1->ecfRedZ.dataMov
			nro = pSrc1->ecfRedZ.numIni
		case DOC_SAT
			dReg = @pSrc1->sat.dataEntSaida
			nro = pSrc1->sat.numero
		case else
			dReg = null
			nro = 0
		end select
		
		var date1 = yyyyMmDd2Days(dReg) shl 32 + nro

		select case as const pSrc2->tipo
		case DOC_NF
			dReg = @pSrc2->nf.dataEntSaida
			nro = pSrc2->nf.numero
		case DOC_CT
			dReg = @pSrc2->ct.dataEntSaida
			nro = pSrc2->ct.numero
		case DOC_NF_ITEM
			dReg = @pSrc2->itemNF.documentoPai->dataEntSaida
			nro = pSrc2->itemNF.documentoPai->numero
		case ECF_REDUCAO_Z
			dReg = @pSrc2->ecfRedZ.dataMov
			nro = pSrc2->ecfRedZ.numIni
		case DOC_SAT
			dReg = @pSrc2->sat.dataEntSaida
			nro = pSrc2->sat.numero
		case else
			dReg = null
			nro = 0
		end select

		var date2 = yyyyMmDd2Days(dReg) shl 32 + nro

		if date2 < date1 then
			*ppDst = pSrc2
			ppDst = @pSrc2->next_
			pSrc2 = *ppDst
			if pSrc2 = NULL then
				*ppDst = pSrc1
				exit do
			end if
		else
			*ppDst = pSrc1
			ppDst = @pSrc1->next_
			pSrc1 = *ppDst
			if pSrc1 = NULL then
				*ppDst = pSrc2
				exit do
			end if
		end if
    loop
	
    function = pDst
end function

''''''''
private function ordenarRegistrosPorData(head as TRegistro ptr) as TRegistro ptr

	const NUMLISTS = 1000
	dim as TRegistro ptr aList(0 to NUMLISTS-1)
    
	if head = NULL then
        return NULL
	end if
    
	var n = head
	do while n <> NULL
        var nn = n->next_
        n->next_ = NULL
		var i = 0
        do while (i < NUMLISTS) and (aList(i) <> NULL)
            n = mergeLists(aList(i), n)
            aList(i) = NULL
			i += 1
        loop
        if i = NUMLISTS then
            i -= 1
		end if
        aList(i) = n
        n = nn
    loop
	
    n = NULL
    for i as integer = 0 to NUMLISTS-1
        n = mergeLists(aList(i), n)
	next
    
	function = n
	
end function

''''''''
function Efd.carregarTxt(nomeArquivo as String) as Boolean

	dim bf as bfile
   
	if not bf.abrir( nomeArquivo ) then
		return false
	end if

	participanteDict = new TDict(2^20)
	itemIdDict = new TDict(2^20)	 
	bemCiapDict = new TDict(2^16)
	infoComplDict = new TDict(2^16)
	sintegraDict = new TDict(2^20)

	regListHead = null
	nroRegs = 0
	
	dim as TArquivoInfo ptr arquivo = arquivos->add()
	arquivo->nome = nomeArquivo
	
	if bf.peek1 <> asc("|") then
		tipoArquivo = TIPO_ARQUIVO_SINTEGRA
		function = carregarSintegra(bf)
	else
		try
			tipoArquivo = TIPO_ARQUIVO_EFD
			var fsize = bf.tamanho - 6500 			'' descontar certificado digital no final do arquivo
			nroLinha = 1
			
			dim as TRegistro ptr tail = null

			do while bf.temProximo()		 
				var reg = new TRegistro
				reg->arquivo = arquivo

				onProgress(null, (bf.posicao / fsize) * 0.66)
				
				if lerRegistro( bf, reg ) then 
					if reg->tipo <> DESCONHECIDO then
						select case as const reg->tipo
						'' fim de arquivo?
						case FIM_DO_ARQUIVO
							delete reg
							onProgress(null, 1)
							exit do

						'' adicionar ao DB
						case DOC_NF, _
							 DOC_NF_ITEM, _
							 DOC_NF_ANAL, _
							 DOC_CT, _
							 ECF_REDUCAO_Z, _
							 DOC_SAT, _
							 DOC_NF_ITEM_RESSARC_ST, _
							 ITEM_ID, _
							 MESTRE
							addRegistroAoDB(reg)
						end select
						
						'' adicionar ao fim da lista
						if tail = null then
							regListHead = reg
							tail = reg
						else
							tail->next_ = reg
							tail = reg
						end if

						nroRegs += 1
					else
						delete reg
					end if
				 
					nroLinha += 1
				else
					exit do
				end if
			loop
		
		catch
			onError(!"\r\nErro ao carregar o registro da linha (" & nroLinha & !") do arquivo\r\n")
		endtry
		
		regListHead = ordenarRegistrosPorData(regListHead)
		
		onProgress(null, 1)

		function = true
	  
	end if

	bf.fechar()
   
end function

''''''''
function Efd.carregarCsvNFeDestSAFI(bf as bfile, emModoOutrasUFs as boolean) as TDFe ptr
	
	var dfe = new TDFe
	
	dfe->operacao			= ENTRADA
	
	if not emModoOutrasUFs then
		dfe->chave				= bf.charCsv
		dfe->dataEmi			= csvDate2YYYYMMDD(bf.charCsv)
		dfe->cnpjEmit			= bf.charCsv
		dfe->nomeEmit			= bf.charCsv
		dfe->nfe.ieEmit			= trim(bf.charCsv)
		dfe->cnpjDest			= bf.charCsv
		dfe->ufDest				= UF_SIGLA2COD(bf.charCsv)
		dfe->nomeDest			= bf.charCsv
		dfe->nfe.bcICMSTotal	= bf.dblCsv
		dfe->nfe.ICMSTotal		= bf.dblCsv
		dfe->nfe.bcICMSSTTotal	= bf.dblCsv
		dfe->nfe.ICMSSTTotal	= bf.dblCsv
		dfe->valorOperacao		= bf.dblCsv
		dfe->ufEmit				= UF_SIGLA2COD(bf.charCsv)
		dfe->numero				= bf.intCsv
		dfe->serie				= bf.intCsv
		dfe->modelo				= bf.intCsv
	else
		dfe->chave				= bf.charCsv
		dfe->cnpjDest			= bf.charCsv
		dfe->nomeDest			= bf.charCsv
		dfe->dataEmi			= csvDate2YYYYMMDD(bf.charCsv)
		dfe->ufDest				= 35
		dfe->cnpjEmit			= bf.charCsv
		dfe->nomeEmit			= bf.charCsv
		dfe->ufEmit				= UF_SIGLA2COD(bf.charCsv)
		dfe->nfe.bcICMSTotal	= bf.dblCsv
		dfe->nfe.ICMSTotal		= bf.dblCsv
		dfe->nfe.bcICMSSTTotal	= bf.dblCsv
		dfe->nfe.ICMSSTTotal	= bf.dblCsv
		dfe->valorOperacao		= bf.dblCsv
		dfe->modelo				= bf.intCsv
		dfe->serie				= bf.intCsv
		dfe->numero				= bf.intCsv
	end if

	'' pular \r\n
	bf.char1
	bf.char1
	
	function = dfe
	
end function

''''''''
function Efd.carregarCsvNFeEmitSAFI(bf as bfile) as TDFe ptr
	
	var chave = bf.charCsv
	var dfe = cast(TDFe ptr, chaveDFeDict->lookup(chave))	
	if dfe = null then
		dfe = new TDFe
	end if
	
	dfe->chave				= chave
	dfe->dataEmi			= csvDate2YYYYMMDD(bf.charCsv)
	dfe->cnpjEmit			= bf.charCsv
	dfe->nomeEmit			= bf.charCsv
	dfe->nfe.ieEmit			= trim(bf.charCsv)
	dfe->ufEmit				= 35
	dfe->cnpjDest			= bf.charCsv
	dfe->ufDest				= UF_SIGLA2COD(bf.charCsv)
	dfe->nomeDest			= bf.charCsv
	dfe->nfe.bcICMSTotal	= bf.dblCsv
	dfe->nfe.ICMSTotal		= bf.dblCsv
	dfe->nfe.bcICMSSTTotal	= bf.dblCsv
	dfe->nfe.ICMSSTTotal	= bf.dblCsv
	dfe->valorOperacao		= bf.dblCsv
	var op = bf.charCsv
	dfe->operacao			= iif(op[0] = asc("S"), SAIDA, ENTRADA)
	dfe->numero				= bf.intCsv
	dfe->serie				= bf.intCsv
	dfe->modelo				= bf.intCsv
	
	'' devolu��o? inverter emit <-> dest
	if dfe->operacao = ENTRADA then
		swap dfe->cnpjEmit, dfe->cnpjDest
		swap dfe->ufEmit, dfe->ufDest
	end if
	
	'' pular \r\n
	bf.char1
	bf.char1
	
	function = dfe
	
end function

''''''''
function Efd.carregarCsvNFeEmitItensSAFI(bf as bfile, chave as string) as TDFe_NFeItem ptr
	
	var item = new TDFe_NFeItem
	
	bf.charCsv				'' pular vers�o
	bf.charCsv				'' pular cnpj emitente
	bf.charCsv				'' pular ie emitente
	bf.charCsv				'' pular cnpj dest
	item->modelo 			= bf.intCsv
	item->serie				= bf.intCsv
	item->numero			= bf.intCsv
	bf.charCsv				'' pular data emi
	item->cfop				= bf.intCsv
	item->nroItem			= bf.intCsv
	item->codProduto		= bf.charCsv
	item->descricao			= bf.charCsv
	item->qtd				= bf.dblCsv
	item->unidade			= bf.charCsv
	item->valorProduto		= bf.dblCsv
	item->desconto			= bf.dblCsv
	item->despesasAcess		= bf.dblCsv
	item->bcICMS			= bf.dblCsv
	item->aliqICMS			= bf.dblCsv
	item->ICMS				= bf.dblCsv
	item->bcICMSST			= bf.dblCsv
	item->IPI				= bf.dblCsv
	item->next_ = null
	
	chave = bf.charCsv
	
	'' pular \r\n
	bf.char1
	bf.char1
	
	function = item
end function

''''''''
function Efd.carregarCsvCTeSAFI(bf as bfile, emModoOutrasUFs as boolean) as TDFe ptr
	var dfe = new TDFe
	
	'' NOTA: s� ser� poss�vel saber se � opera��o de entrada ou sa�da quando pegarmos 
	''       o CNPJ base do contribuinte, que s� vem no final do arquivo.......
	dfe->operacao			= DESCONHECIDA			
	
	bf.charCsv				'' pular chave quebrada
	dfe->serie				= bf.intCsv
	dfe->numero				= bf.intCsv
	dfe->cnpjEmit			= bf.charCsv
	dfe->dataEmi			= csvDate2YYYYMMDD(bf.charCsv)
	dfe->nomeEmit			= bf.charCsv
	dfe->ufEmit				= UF_SIGLA2COD(bf.charCsv)
	dfe->cte.cnpjToma		= bf.charCsv
	dfe->cte.nomeToma		= bf.charCsv
	dfe->cte.ufToma			= bf.charCsv
	dfe->cte.cnpjRem		= bf.charCsv
	dfe->cte.nomeRem		= bf.charCsv
	dfe->cte.ufRem			= bf.charCsv
	dfe->cnpjDest			= bf.charCsv
	dfe->nomeDest			= bf.charCsv
	dfe->ufDest				= UF_SIGLA2COD(bf.charCsv)
	dfe->cte.cnpjExp		= bf.charCsv
	dfe->cte.ufExp			= bf.charCsv
	dfe->cte.cnpjReceb		= bf.charCsv
	dfe->cte.ufReceb		= bf.charCsv
	dfe->cte.tipo			= valint(left(bf.charCsv,1))
	dfe->chave				= bf.charCsv
	dfe->valorOperacao		= bf.dblCsv
	dfe->cte.valorReceber	= bf.dblCsv
	dfe->cte.qtdCCe			= bf.dblCsv
	dfe->cte.cfop			= bf.intCsv
	dfe->cte.nomeMunicIni	= bf.charCsv
	dfe->cte.ufIni			= bf.charCsv
	dfe->cte.nomeMunicFim	= bf.charCsv
	dfe->cte.ufFim			= bf.charCsv
	dfe->modelo				= 57
	
	'' pular \r\n
	bf.char1
	bf.char1
	
	''
	if cteListHead = null then
		cteListHead = @dfe->cte
	else
		cteListTail->next_ = @dfe->cte
	end if
	
	cteListTail = @dfe->cte
	dfe->cte.next_ = null
	dfe->cte.parent = dfe
	
	function = dfe
	
end function

const BO_CSV_SEP = asc(!"\t")
const BO_CSV_DIG = asc(".")

''''''''
function Efd.carregarCsvNFeEmitItens(bf as bfile, chave as string) as TDFe_NFeItem ptr
	
	var item = new TDFe_NFeItem
	
	'' chave_nfe	num_doc_fiscal	cod_serie_doc_fiscal	cod_modelo	ind_tipo_documento_fiscal	ind_situacao_doc_fiscal	data_emissao	
	'' nome_rsocial_emit	num_cnpj_emit	num_ie_emit	cod_drt_emit	cod_est_emit	nome_rsocial_dest	num_cnpj_dest	num_cpf_dest	
	'' num_ie_dest	cod_drt_dest	cod_est_dest	num_item	descr_prod	cod_prod_servico	cod_gtin	cod_ncm	cod_cfop	
	'' cod_tributacao_icms	cod_csosn	perc_aliquota_icms	perc_aliquota_base_calc	perc_aliquota_icms_st	perc_reduc_icms_st	
	'' quant_comercial	unid_comercial	valor_produto_servico	valor_base_calc_icms	valor_icms	valor_base_calc_icms_st	valor_icms_st	
	'' valor_bc_icms_st_retido	valor_icms_st_retido	valor_ipi	valor_desconto	valor_frete	ind_modalidade_frete	valor_seguro	
	'' valor_outras_desp	valor_pis	valor_cofins	num_docto_importacao	num_fci	data_desembaraco	cod_est_desembaraco	
	'' descr_inf_adic_produto	ind_origem_mercadoria	cod_cnae

	chave 					= bf.varchar(BO_CSV_SEP)

	item->numero			= bf.varint(BO_CSV_SEP) ''vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->serie				= bf.varint(BO_CSV_SEP)
	item->modelo 			= bf.varint(BO_CSV_SEP)
	bf.varchar(BO_CSV_SEP) '' tipo
	bf.varchar(BO_CSV_SEP)	'' situa��o
	bf.varchar(BO_CSV_SEP) '' data emi
	bf.varchar(BO_CSV_SEP) '' raz�o social emi
	bf.varchar(BO_CSV_SEP) '' cnpj emi
	bf.varchar(BO_CSV_SEP) '' ie emi
	bf.varchar(BO_CSV_SEP) '' drt emi
	bf.varchar(BO_CSV_SEP)	'' uf emi
	bf.varchar(BO_CSV_SEP)	'' raz�o social dest
	bf.varchar(BO_CSV_SEP) '' cnpj dest
	bf.varchar(BO_CSV_SEP) '' cpf dest
	bf.varchar(BO_CSV_SEP) '' ie dest
	bf.varchar(BO_CSV_SEP) '' drt dest
	bf.varchar(BO_CSV_SEP)	'' uf dest
	item->nroItem			= bf.varint(BO_CSV_SEP)
	item->descricao			= bf.varchar(BO_CSV_SEP)
	item->codProduto		= bf.varchar(BO_CSV_SEP)
	bf.varchar(BO_CSV_SEP)	'' GTIN
	item->ncm				= bf.varint(BO_CSV_SEP)
	item->cfop				= bf.varint(BO_CSV_SEP)
	item->cst				= bf.varint(BO_CSV_SEP)
	bf.varchar(BO_CSV_SEP) '' CSOSN
	item->aliqICMS			= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	bf.varchar(BO_CSV_SEP) '' redu��o bc
	bf.varchar(BO_CSV_SEP) '' al�q ICMS ST
	bf.varchar(BO_CSV_SEP) '' redu��o bc ST
	item->qtd				= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->unidade			= bf.varchar(BO_CSV_SEP)
	item->valorProduto		= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->bcICMS			= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->ICMS				= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->bcICMSST			= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	bf.varchar(BO_CSV_SEP) '' ICMS ST
	bf.varchar(BO_CSV_SEP) '' bc ICMS ST anterior
	bf.varchar(BO_CSV_SEP) '' ICMS ST anterior
	item->IPI				= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	item->desconto			= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	bf.varchar(BO_CSV_SEP) '' frete
	bf.varchar(BO_CSV_SEP) '' indicador frete
	bf.varchar(BO_CSV_SEP) '' seguro
	item->despesasAcess		= bf.vardbl(BO_CSV_SEP, BO_CSV_DIG)
	bf.varchar(BO_CSV_SEP) '' pis
	bf.varchar(BO_CSV_SEP) '' cofins
	bf.varchar(BO_CSV_SEP) '' num doc importacao
	bf.varchar(BO_CSV_SEP) '' num fci
	bf.varchar(BO_CSV_SEP) '' data desembaraco
	bf.varchar(BO_CSV_SEP) '' uf desembaraco
	bf.varchar(BO_CSV_SEP) '' info adicional
	bf.varchar(BO_CSV_SEP) '' origem mercadoria
	bf.varchar(BO_CSV_SEP) '' cnae
	item->next_ = null
	
	'' pular \r\n
	bf.char1
	bf.char1
	
	function = item
end function

''''''''
sub Efd.adicionarDFe(dfe as TDFe ptr)
	
	if dfeListHead = null then
		dfeListHead = dfe
	else
		dfeListTail->next_ = dfe
	end if
	
	dfeListTail = dfe
	dfe->next_ = null
	
	if chaveDFeDict->lookup(dfe->chave) = null then
		chaveDFeDict->add(dfe->chave, dfe)
	end if

	'' adicionar ao db
	select case dfe->operacao
	case ENTRADA
		'' (cnpjEmit, ufEmit, serie, numero, modelo, chave, dataEmit, valorOp, ieEmit)
		db_dfeEntradaInsertStmt->reset()
		db_dfeEntradaInsertStmt->bind(1, dfe->cnpjEmit)
		db_dfeEntradaInsertStmt->bind(2, dfe->ufEmit)
		db_dfeEntradaInsertStmt->bind(3, dfe->serie)
		db_dfeEntradaInsertStmt->bind(4, dfe->numero)
		db_dfeEntradaInsertStmt->bind(5, dfe->modelo)
		db_dfeEntradaInsertStmt->bind(6, dfe->chave)
		db_dfeEntradaInsertStmt->bind(7, dfe->dataEmi)
		db_dfeEntradaInsertStmt->bind(8, dfe->valorOperacao)
		if len(dfe->nfe.ieEmit) > 0 then
			db_dfeEntradaInsertStmt->bind(9, dfe->nfe.ieEmit)
		else
			db_dfeEntradaInsertStmt->bindNull(9)
		end if
		
		if not db->execNonQuery(db_dfeEntradaInsertStmt) then
			onError("Erro ao inserir DFe de entrada: " & *db->getErrorMsg())
		end if
	
	case SAIDA
		'' (cnpjDest, ufDest, serie, numero, modelo, chave, dataEmit, valorOp, ieDest)
		db_dfeSaidaInsertStmt->reset()

		db_dfeSaidaInsertStmt->bind(1, dfe->cnpjDest)
		db_dfeSaidaInsertStmt->bind(2, dfe->ufDest)
		db_dfeSaidaInsertStmt->bind(3, dfe->serie)
		db_dfeSaidaInsertStmt->bind(4, dfe->numero)
		db_dfeSaidaInsertStmt->bind(5, dfe->modelo)
		db_dfeSaidaInsertStmt->bind(6, dfe->chave)
		db_dfeSaidaInsertStmt->bind(7, dfe->dataEmi)
		db_dfeSaidaInsertStmt->bind(8, dfe->valorOperacao)
		if len(dfe->nfe.ieDest) > 0 then
			db_dfeSaidaInsertStmt->bind(9, dfe->nfe.ieDest)
		else
			db_dfeSaidaInsertStmt->bindNull(9)
		end if
	
		if not db->execNonQuery(db_dfeSaidaInsertStmt) then
			onError("Erro ao inserir DFe de sa�da: " & *db->getErrorMsg())
		end if
	end select
	
	nroDfe += 1

end sub

''''''''
sub Efd.adicionarItemDFe(chave as const zstring ptr, item as TDFe_NFeItem ptr)
		'' (serie, numero, modelo, numItem, chave, cfop, valorProd, valorDesc, valorAcess, bc, aliq, icms, bcIcmsST, ncm, cst, qtd, unidade, codProduto, descricao) 
		db_itensDfeSaidaInsertStmt->reset()
		db_itensDfeSaidaInsertStmt->bind(1, item->serie)
		db_itensDfeSaidaInsertStmt->bind(2, item->numero)
		db_itensDfeSaidaInsertStmt->bind(3, item->modelo)
		db_itensDfeSaidaInsertStmt->bind(4, item->nroItem)
		db_itensDfeSaidaInsertStmt->bind(5, chave)
		db_itensDfeSaidaInsertStmt->bind(6, item->cfop)
		db_itensDfeSaidaInsertStmt->bind(7, item->valorProduto)
		db_itensDfeSaidaInsertStmt->bind(8, item->desconto)
		db_itensDfeSaidaInsertStmt->bind(9, item->despesasAcess)
		db_itensDfeSaidaInsertStmt->bind(10, item->bcICMS)
		db_itensDfeSaidaInsertStmt->bind(11, item->aliqICMS)
		db_itensDfeSaidaInsertStmt->bind(12, item->icms)
		db_itensDfeSaidaInsertStmt->bind(13, item->bcIcmsST)
		db_itensDfeSaidaInsertStmt->bind(14, item->ncm)
		db_itensDfeSaidaInsertStmt->bind(15, item->cst)
		db_itensDfeSaidaInsertStmt->bind(16, item->qtd)
		if opcoes.manterDb then
			db_itensDfeSaidaInsertStmt->bind(17, item->unidade)
			db_itensDfeSaidaInsertStmt->bind(18, item->codProduto)
			db_itensDfeSaidaInsertStmt->bind(19, item->descricao)
		else
			db_itensDfeSaidaInsertStmt->bind(17, null)
			db_itensDfeSaidaInsertStmt->bind(18, null)
			db_itensDfeSaidaInsertStmt->bind(19, null)
		end if
	
		if not db->execNonQuery(db_itensDfeSaidaInsertStmt) then
			onError("Erro ao inserir Item DFe de entrada: " & *db->getErrorMsg())
		end if
end sub

''''''''
function Efd.carregarCsv(nomeArquivo as String) as Boolean

	dim bf as bfile
   
	if not bf.abrir( nomeArquivo ) then
		return false
	end if
	
	dim as integer tipoArquivo
	dim as boolean isSafi = true
	if instr( nomeArquivo, "BO_NFe_Destinatario" ) > 0 then
		tipoArquivo = BO_NFe_Dest
		nfeDestSafiFornecido = true
	
	elseif instr( nomeArquivo, "BO_NFe_Emitente_Itens" ) > 0 then
		tipoArquivo = BO_NFe_Emit_Itens
		itemNFeSafiFornecido = true
	
	elseif instr( nomeArquivo, "BO_NFe_Emitente" ) > 0 then
		tipoArquivo = BO_NFe_Emit
		nfeEmitSafiFornecido = true
	
	elseif instr( nomeArquivo, "BO_CTe_CNPJ" ) > 0 then
		tipoArquivo = BO_CTe
		cteListHead = null
		cteListTail = null
		cteSafiFornecido = true
		
	elseif instr( nomeArquivo, "NFE_Emitente_Itens_SP_OSF" ) > 0 then
		tipoArquivo = BO_NFe_Emit_Itens
		isSafi = false
		itemNFeSafiFornecido = true
	
	else
		onError("Erro: imposs�vel resolver tipo de arquivo pelo nome")
		return false
	end if

	var nroLinha = 1
		
	try
		var fsize = bf.tamanho

		'' pular header
		pularLinha(bf)
		nroLinha += 1
		
		var emModoOutrasUFs = false
		
		do while bf.temProximo()		 
			onProgress(null, bf.posicao / fsize)
			
			if isSafi then
				'' outro header?
				if bf.peek1 <> asc("""") then
					'' final de arquivo?
					
					var linha = lcase(lerLinha(bf))
					if left(linha, 22) = "cnpj base contribuinte" or left(linha, 26) = "cnpj/cpf base contribuinte" then
						onProgress(null, 1)
						nroLinha += 1
						
						'' se for CT-e, temos que ler o CNPJ base do contribuinte para fazer um 
						'' patch em todos os tipos de opera��o (sa�da ou entrada)
						if tipoArquivo = BO_CTe then
							var cnpjBase = bf.charCsv
							var cte = cteListHead
							do while cte <> null 
								if left(cte->parent->cnpjEmit,8) = cnpjBase then
									cte->parent->operacao = SAIDA
								elseif left(cte->cnpjToma,8) = cnpjBase then
									cte->parent->operacao = ENTRADA
								end if
								adicionarDFe(cte->parent)
								cte = cte->next_
							loop
						end if
						exit do
					else
						emModoOutrasUFs = true
					end if
				end if
			end if
		
			select case as const tipoArquivo  
			case BO_NFe_Dest
				var dfe = carregarCsvNFeDestSAFI( bf, emModoOutrasUFs )
				if dfe <> null then
					adicionarDFe(dfe)
				end if
			
			case BO_NFe_Emit
				var dfe = carregarCsvNFeEmitSAFI( bf )
				if dfe <> null then
					adicionarDFe(dfe)
				end if
				
			case BO_NFe_Emit_Itens
				var chave = ""
				var nfeItem = iif(isSafi, _
					carregarCsvNFeEmitItensSAFI( bf, chave ), _
					carregarCsvNFeEmitItens( bf, chave ))
				if nfeItem <> null then
					adicionarItemDFe(chave, nfeItem)

					var dfe = cast(TDFe ptr, chaveDFeDict->lookup(chave))
					'' nf-e n�o encontrada? pode acontecer se processarmos o csv de itens antes do csv de nf-e
					if dfe = null then
						dfe = new TDFe
						'' s� adicionar ao dicion�rio, depois ser� adicionado por adicionarDFe() no case acima
						dfe->chave = chave
						chaveDFeDict->add(dfe->chave, dfe)
					end if
					
					if dfe->nfe.itemListHead = null then
						dfe->nfe.itemListHead = nfeItem
					else
						dfe->nfe.itemListTail->next_ = nfeItem
					end if
					
					dfe->nfe.itemListTail = nfeItem
				end if
			
			case BO_CTe
				var dfe = carregarCsvCTeSAFI( bf, emModoOutrasUFs )
			end select
			
			nroLinha += 1
		loop
		
		if not isSafi then
			onProgress(null, 1)
		end if
		
		function = true
	
	catch
		onError(!"\r\n\tErro ao carregar linha " & nroLinha & !"\r\n")
		function = false
	endtry
	   
	bf.fechar()
	
end function

private function dbl2Cnpj(valor as double) as string
	return iif(valor <> 0, right("00000000000000" + str(valor), 14), "")
end function 

private function limparCNPJ(valor as string) as string
	return iif(len(valor) > 0, right("00000000000000" + strreplace(strreplace(strreplace(valor, ".", ""), "/", ""), "-", ""), 14), "")
end function

#define limparIE(valor) strreplace(valor, ".", "")

''''''''
function Efd.carregarXlsxNFeDest(rd as ExcelReader ptr) as TDFe ptr
	
	'' Chave Acesso NFe,	N�mero,	S�rie,	Modelo,	Data Emiss�o,	Raz�o Social Emitente
	'' CNPJ Emitente,	N�mero CPF Emitente,	Inscri��o Estadual Emitente,	CRT,	DRT Emitente
	'' UF Emit,	Raz�o Social Destinat�rio,	CNPJ Destinat�rio,	Inscri��o Estadual Destinat�rio,	DRT Destinat�rio
	'' UF Dest,	Tipo Doc Fiscal,	Descri��o Natureza Opera��o,	Peso Liquido,	Peso Bruto	
	'' Informa��es Interesse Fisco,	Informa��es Complementares Interesse Contribuinte,	Indicador Modalidade Frete,	Situa��o Documento,	Dt. Cancelamento	
	'' Mercadoria - Valor,	Raz�o Social Transportador	CNPJ do Transportador,	Inscri��o Estadual Transportador,	Placa Ve�culo Transportador	
	'' UF Ve�culo Transportador,	Total BC  ICMS,	Total ICMS,	Total BC ICMS-ST,	Total ICMS-ST	
	'' Total NFe,	Valor Total Frete,	Valor Total Seguro,	Quantidade Cartas de Corre��o Eletr�nicas,	Quantidade Manifesta��es Destinat�rio

	var chave				= rd->read
	if len(chave) <> 44 then
		return null
	end if
		
	if chaveDFeDict->lookup(chave) <> null then
		return null
	end if

	var dfe = new TDFe
	
	dfe->loader				= LOADER_NFE_DEST
	dfe->operacao			= ENTRADA
	dfe->chave				= chave
	dfe->numero				= rd->readDbl
	dfe->serie				= rd->readInt
	dfe->modelo				= rd->readInt
	dfe->dataEmi			= rd->readDate
	dfe->nomeEmit			= rd->read(true)
	dfe->cnpjEmit			= dbl2Cnpj(rd->readDbl)
	rd->skip '' cpf emit
	dfe->nfe.ieEmit			= trim(limparIE(rd->read))
	rd->skip '' crt emit
	rd->skip '' drt emit
	dfe->ufEmit				= UF_SIGLA2COD(rd->read)
	dfe->nomeDest			= rd->read(true)
	dfe->cnpjDest			= dbl2Cnpj(rd->readDbl)
	dfe->nfe.ieDest			= trim(limparIE(rd->read))
	rd->skip '' drt dest
	dfe->ufDest				= UF_SIGLA2COD(rd->read)
	rd->skip '' tipo doc
	rd->skip '' descri��o op
	rd->skip '' peso liq
	rd->skip '' peso bruto
	rd->skip '' info fisco
	rd->skip '' info contrib
	rd->skip '' frete
	rd->skip '' situa��o doc
	rd->skip '' data canc
	rd->skip '' merc valor
	rd->skip '' transportador
	rd->skip '' cnpj transportador
	rd->skip '' ie transportador
	rd->skip '' placa transportador
	rd->skip '' uf transportador
	dfe->nfe.bcICMSTotal	= rd->readDbl
	dfe->nfe.ICMSTotal		= rd->readDbl
	dfe->nfe.bcICMSSTTotal	= rd->readDbl
	dfe->nfe.ICMSSTTotal	= rd->readDbl
	dfe->valorOperacao		= rd->readDbl
	
	function = dfe

end function

''''''''
function Efd.carregarXlsxNFeDestItens(rd as ExcelReader ptr) as TDFe ptr
	
	'' Chave de Acesso NFe, N�mero Documento Fiscal, S�rie Documento Fiscal, Modelo Documento Fiscal, Tipo Documento Fiscal, Situa��o Documento Fiscal, 
	'' Data Emiss�o, Raz�o Social Emitente, CNPJ Emitente, CPF Emitente, Inscri��o Estadual Emitente, DRT Emitente,	UF Emitente, Raz�o Social Destinat�rio,
	'' CNPJ Destinat�rio, CPF Destinat�rio,	Inscri��o Estadual Destinat�rio, DRT Destinat�rio, UF Destinat�rio, Item, Descri��o Produto, C�digo Produto, 
	'' GTIN, NCM, CFOP, CST, O/CSOSN, Al�quota ICMS, Percentual Redu��o Base de C�lculo ICMS, Al�quota ICMS-ST, Percentual Redu��o Base de C�lculo ICMS-ST, 
	'' Quantidade Comercial, Unidade Comercial, Valor Produto ou Servi�o, Valor Base de C�lculo ICMS, Valor ICMS, Valor Base C�lculo ICMS-ST, Valor ICMS-ST
	'' Valor Base C�lculo ICMS-ST Retido Opera��o Anterior, Valor ICMS-ST Retido Opera��o Anterior, Valor IPI, Valor Desconto,
	'' Valor Frete, Indicador Modalidade Frete, Valor Seguro, Valor Outras Despesas Acess�rias, Valor PIS, Valor COFINS, 
	'' Percentual Al�quota Cr�dito Simples Nacional, Valor Cr�dito Simples Nacional,
	'' N�mero DI, N�mero FCI, Data Desembara�o, C�digo UF Desembara�o, Descri��o Informa��es Adicionais Produto
	
	var chave				= rd->read
	if len(chave) <> 44 then
		return null
	end if
	
	var dfe = cast(TDFe ptr, chaveDFeDict->lookup(chave))	
	if dfe = null then
		dfe = new TDFe
	else
		if dfe->loader <> LOADER_NFE_DEST_ITENS then
			return null
		end if
	end if
	
	dfe->loader				= LOADER_NFE_DEST_ITENS
	dfe->operacao			= ENTRADA
	dfe->chave				= chave
	dfe->numero				= rd->readDbl
	dfe->serie				= rd->readInt
	dfe->modelo				= rd->readInt
	rd->skip '' tipo
	rd->skip '' situa��o
	dfe->dataEmi			= rd->readDate
	dfe->nomeEmit			= rd->read(true)
	dfe->cnpjEmit			= dbl2Cnpj(rd->readDbl)
	rd->skip '' cpf emit
	dfe->nfe.ieEmit			= trim(limparIE(rd->read))
	rd->skip '' drt emit
	dfe->ufEmit				= UF_SIGLA2COD(rd->read)
	dfe->nomeDest			= rd->read(true)
	dfe->cnpjDest			= dbl2Cnpj(rd->readDbl)
	if dfe->cnpjDest = "" then
		dfe->cnpjDest 		= rd->read
	else
		rd->skip '' cpf dest
	end if
	rd->skip '' ie dest
	rd->skip '' drt dest
	dfe->ufDest				= UF_SIGLA2COD(rd->read)
	rd->skip '' item
	rd->skip '' descri��o prod
	rd->skip '' c�digo prod
	rd->skip '' GTIN
	rd->skip '' NCM
	rd->skip '' CFOP
	rd->skip '' CST
	rd->skip '' CSOSN
	rd->skip '' aliq
	rd->skip '' red bc icms
	rd->skip '' aliq ST
	rd->skip '' red bc icms ST
	rd->skip '' qtd
	rd->skip '' unidade
	dfe->valorOperacao		+= rd->readDbl
	dfe->nfe.bcICMSTotal	+= rd->readDbl
	dfe->nfe.ICMSTotal		+= rd->readDbl
	dfe->nfe.bcICMSSTTotal	+= rd->readDbl
	dfe->nfe.ICMSSTTotal	+= rd->readDbl

	function = dfe

end function

''''''''
function Efd.carregarXlsxNFeEmit(rd as ExcelReader ptr) as TDFe ptr
	
	var chave = rd->read
	if len(chave) <> 44 then
		return null
	end if
	
	var dfe = cast(TDFe ptr, chaveDFeDict->lookup(chave))	
	if dfe = null then
		dfe = new TDFe
	end if
	
	'' Chave Acesso NFe,	N�mero,	S�rie,	Modelo,	Data Emiss�o,	Raz�o Social Emitente
	'' CNPJ Emitente,	Inscri��o Estadual Emitente,	CRT, DRT Emit,	UF Emit
	'' Raz�o Social Destinat�rio,	CNPJ ou CPF do Destinat�rio,	Inscri��o Estadual Destinat�rio,	CNAE Destinat�rio,	Cod Cnae Destinat�rio (Cadesp)	
	'' DRT Dest,	UF Dest,	Tipo Doc Fiscal,	Descri��o Natureza Opera��o,	Peso Liquido(NFe SP Volume)
	'' Peso Bruto(NFe SP Volume),	Informa��es Interesse Fisco,	Informa��es Complementares Interesse Contribuinte,	Indicador Modalidade Frete,	Situa��o Documento
	'' Dt. Cancelamento,	Mercadoria - Valor,	Raz�o Social Transportador,	CNPJ do Transportador,	Inscri��o Estadual Transportador
	'' Placa Ve�culo Transportador,	UF Ve�culo Transportador,	Total BC  ICMS,	Total ICMSv,	Total BC ICMS-ST
	'' Total ICMS-ST,	Total NFe,	Valor Total Frete,	Valor Total Seguro,	Valor ICMS Inter. UF Destino	
	'' Valor ICMS Inter. UF Remetente,	Quantidade Cartas de Corre��o Eletr�nicas,	Quantidade Manifesta��es Detinat�rio

	dfe->loader				= LOADER_NFE_EMIT
	dfe->chave				= chave
	dfe->numero				= rd->readDbl
	dfe->serie				= rd->readInt
	dfe->modelo				= rd->readInt
	dfe->dataEmi			= rd->readDate
	dfe->nomeEmit			= rd->read(true)
	dfe->cnpjEmit			= dbl2Cnpj(rd->readDbl)
	dfe->nfe.ieEmit			= trim(limparIE(rd->read))
	rd->skip '' crt emit
	rd->skip '' drt emit
	dfe->ufEmit				= UF_SIGLA2COD(rd->read)
	dfe->nomeDest			= rd->read(true)
	dfe->cnpjDest			= limparCNPJ(rd->read)
	dfe->nfe.ieDest			= trim(limparIE(rd->read))
	rd->skip '' cnae dest
	rd->skip '' cnae dest cadesp
	rd->skip '' drt dest
	dfe->ufDest				= UF_SIGLA2COD(rd->read)
	var op = rd->read
	dfe->operacao			= iif(op[0] = asc("S"), SAIDA, ENTRADA)
	rd->skip '' descri��o op
	rd->skip '' peso liq
	rd->skip '' peso bruto
	rd->skip '' info fisco
	rd->skip '' info contrib
	rd->skip '' frete
	rd->skip '' situa��o doc
	rd->skip '' data canc
	rd->skip '' merc valor
	rd->skip '' transportador
	rd->skip '' cnpj transportador
	rd->skip '' ie transportador
	rd->skip '' placa transportador
	rd->skip '' uf transportador
	dfe->nfe.bcICMSTotal	= rd->readDbl
	dfe->nfe.ICMSTotal		= rd->readDbl
	dfe->nfe.bcICMSSTTotal	= rd->readDbl
	dfe->nfe.ICMSSTTotal	= rd->readDbl
	dfe->valorOperacao		= rd->readDbl
	
	'' devolu��o? inverter emit <-> dest
	if dfe->operacao = ENTRADA then
		swap dfe->cnpjEmit, dfe->cnpjDest
		swap dfe->ufEmit, dfe->ufDest
	end if

	function = dfe
	
end function

''''''''
function Efd.carregarXlsxNFeEmitItens(rd as ExcelReader ptr, chave as string) as TDFe_NFeItem ptr
	
	'' Chave de Acesso NFe,	N�mero Documento Fiscal,	 S�rie Documento Fiscal,	Modelo Documento Fiscal, Tipo Documento Fiscal,	
	'' Situa��o Documento Fiscal,	Data Emiss�o,	Raz�o Social Emitente,	CNPJ Emitente,	Inscri��o Estadual Emitente,	
	'' DRT Emitente,	UF Emitente,	Raz�o Social Destinat�rio,	CNPJ Destinat�rio,	CPF Destinat�rio,	
	'' Inscri��o Estadual Destinat�rio,	DRT Destinat�rio,	UF Destinat�rio,	Item,	Descri��o Produto,	
	'' C�digo Produto,	GTIN,	NCM,	CFOP,	CST,	
	'' O/CSOSN,	Al�quota ICMS,	Percentual Redu��o Base de C�lculo ICMS,	Al�quota ICMS-ST,	Percentual Redu��o Base de C�lculo ICMS-ST,	
	'' Quantidade Comercial,	Unidade Comercial,	 Valor Produto ou Servi�o ,	 Valor Base de C�lculo ICMS,	 Valor ICMS, 	
	'' Valor Base C�lculo ICMS-ST,	Valor ICMS-ST,	Valor Base C�lculo ICMS-ST Retido Opera��o Anterior,	Valor ICMS-ST Retido Opera��o Anterior,	Valor IPI,	
	'' Valor Desconto,	Valor Frete,	Indicador Modalidade Frete,	Valor Seguro,	Valor Outras Despesas Acess�rias, 
	'' Valor PIS,	Valor COFINS,	N�mero DI,	N�mero FCI,	Data Desembara�o
	'' C�digo UF Desembara�o,	Descri��o Informa��es Adicionais Produto
		
	chave = rd->read
	if len(chave) <> 44 then
		return null
	end if
	
	var item = new TDFe_NFeItem
	
	item->numero			= rd->readDbl
	item->serie				= rd->readInt
	item->modelo 			= rd->readInt
	rd->skip '' tipo
	rd->skip	'' situa��o
	rd->skip '' data emi
	rd->skip '' raz�o social emi
	rd->skip '' cnpj emi
	rd->skip '' ie emi
	rd->skip '' drt emi
	rd->skip	'' uf emi
	rd->skip	'' raz�o social dest
	rd->skip '' cnpj dest
	rd->skip '' cpf dest
	rd->skip '' ie dest
	rd->skip '' drt dest
	rd->skip	'' uf dest
	item->nroItem			= rd->readInt
	item->descricao			= rd->read(true)
	item->codProduto		= rd->read
	rd->skip	'' GTIN
	item->ncm				= rd->readInt
	item->cfop				= rd->readInt
	item->cst				= rd->readInt
	rd->skip '' CSOSN
	item->aliqICMS			= rd->readDbl
	rd->skip '' redu��o bc
	rd->skip '' al�q ICMS ST
	rd->skip '' redu��o bc ST
	item->qtd				= rd->readDbl
	item->unidade			= rd->read
	item->valorProduto		= rd->readDbl
	item->bcICMS			= rd->readDbl
	item->ICMS				= rd->readDbl
	item->bcICMSST			= rd->readDbl
	rd->skip '' ICMS ST
	rd->skip '' bc ICMS ST anterior
	rd->skip '' ICMS ST anterior
	item->IPI				= rd->readDbl
	item->desconto			= rd->readDbl
	rd->skip '' frete
	rd->skip '' indicador frete
	rd->skip '' seguro
	item->despesasAcess		= rd->readDbl
	item->next_ = null
		
	function = item
	
end function

''''''''
function Efd.carregarXlsxCTe(rd as ExcelReader ptr, op as TipoOperacao) as TDFe ptr
	
	'' ---em branco---,	Chave Acesso CT-e (char),	S�rie,	Num CTe,	Data Emiss�o	
	'' CNPJ Emitente,	Num. Inscr. Est. Emitente,	Raz�o Social Emitente,	UF Emitente,	CNPJ Tomador,	
	'' Num Inscr. Est. Tomador,	Raz�o Social Tomador,	Indicador Tomador Servi�o,	UF Tomador,	CNPJ Remetente,	
	'' Raz�o Social Remetente,	UF Remetente,	CNPJ Destinat�rio,	Raz�o Social Destinat�rio,	UF Destinat�rio,	
	'' CNPJ Expedidor,	UF Expedidor,	CNPJ Recebedor,	UF Recebedor,	Tipo CT-e,	
	'' indSN,	C�digo CFOP,	Descr. Nat. Opera��o,	Descr. Modal,	Descr. Servico,	
	'' Descr. Cst,	Munic�pio Inicial,	UF Inicial,	Munic�pio Final,	UF Final,	
	'' Aliq�ota Icms,	Perc. Redu��o Bc,	Valor Bc St Retido,	Valor Icms St Retido,	Valor Icms OutrasUF,	
	'' Valor Cr�dito Outorgado/Presumido,	Valor Total Prest. Servi�o,	Valor Icms,	Valor Bc ICMS,	Quantidade de CCE,	
	'' Quantidade de manifesta��es do tomador
	
	rd->skip '' ---em branco---
	var chave 				= rd->read
	if len(chave) <> 44 then
		return null
	end if
	
	var dfe = new TDFe

	dfe->operacao			= op
	dfe->chave				= chave
	dfe->serie				= rd->readInt
	dfe->numero				= rd->readInt
	dfe->dataEmi			= rd->readDate
	dfe->cnpjEmit			= dbl2Cnpj(rd->readDbl)
	rd->skip '' ie emit
	dfe->nomeEmit			= rd->read(true)
	dfe->ufEmit				= UF_SIGLA2COD(rd->read)
	dfe->cte.cnpjToma		= dbl2Cnpj(rd->readDbl)
	rd->skip '' ie toma
	dfe->cte.nomeToma		= rd->read(true)
	rd->skip '' ind toma
	dfe->cte.ufToma			= rd->read
	dfe->cte.cnpjRem		= dbl2Cnpj(rd->readDbl)
	dfe->cte.nomeRem		= rd->read(true)
	dfe->cte.ufRem			= rd->read
	dfe->cnpjDest			= dbl2Cnpj(rd->readDbl)
	dfe->nomeDest			= rd->read(true)
	dfe->ufDest				= UF_SIGLA2COD(rd->read)
	dfe->cte.cnpjExp		= dbl2Cnpj(rd->readDbl)
	dfe->cte.ufExp			= rd->read
	dfe->cte.cnpjReceb		= dbl2Cnpj(rd->readDbl)
	dfe->cte.ufReceb		= rd->read
	dfe->cte.tipo			= valint(left(rd->read,1))
	rd->skip '' indSN
	dfe->cte.cfop			= rd->readInt
	rd->skip '' Descr. Nat. Opera��o
	rd->skip '' Descr. Modal
	rd->skip '' Descr. Servico
	rd->skip '' Descr. Cst
	dfe->cte.nomeMunicIni	= rd->read
	dfe->cte.ufIni			= rd->read
	dfe->cte.nomeMunicFim	= rd->read
	dfe->cte.ufFim			= rd->read
	rd->skip '' Aliq�ota Icms
	rd->skip '' Perc. Redu��o Bc
	rd->skip '' Valor Bc St Retido
	rd->skip '' Valor Icms St Retido
	rd->skip '' Valor Icms OutrasUF,	
	rd->skip '' Valor Cr�dito Outorgado/Presumido
	dfe->valorOperacao		= rd->readDbl
	rd->skip '' Valor Icms
	rd->skip '' Valor Bc ICMS
	dfe->cte.valorReceber	= dfe->valorOperacao
	dfe->cte.qtdCCe			= rd->readInt
	dfe->modelo				= 57
	
	function = dfe
	
end function

''''''''
function Efd.carregarXlsxSATItens(rd as ExcelReader ptr, chave as string) as TDFe_NFeItem ptr
	
	'' ---em branco---, Num Inscr. Estadual Emitente,	Data Emiss�o,	Identifica��o CF-e,	N�mero Cupom CF-e,	Indicador Cupom Cancelado	
	'' N�mero S�rie,	Valor ICMS,	N�mero Item,	C�digo Produto,	C�digo EAN,	
	'' Descri��o Produto,	C�digo NCM,	C�digo CFOP 04 Posi��es,	Unidade Comercial,	Quantidade Comercial,	
	'' Indicador Regra C�lculo,	Valor Unit�rio Comercializa��o,	Valor Produtos,	Valor Desconto,	Valor Outro,	
	'' Valor Item,	Valor Rateio Desconto,	Valor Rateio Acrescimo,	Indicador Origem,	C�digo CST/CSOSN,	
	'' Al�quota ICMS,	C�digo CST PIS,	Valor Base C�lculo PIS,	Al�quota PIS,	Valor PIS,	
	'' Quantidade Vendida PIS,	Valor Al�quota PIS,	Valor Base C�lculo PIS-ST,	Al�quota PIS-ST,	Quantidade Vendida PIS-ST,	
	'' Valor Al�quota PIS-ST,	Valor PIS-ST,	C�digo CST COFINS,	Valor Base C�lculo COFINS,	Al�quota COFINS	Valor COFINS,	
	'' Quantidade Vendida COFINS,	Valor Al�quota COFINS,	Valor Base C�lculo COFINS-ST,	Al�quota COFINS-ST,	Quantidade Vendida COFINS-ST,	
	'' Valor Al�quota COFINS-ST,	Valor COFINS-ST,	Informa��es Adicicionais,	Descri��o Campo,	Descri��o Texto Campo
		
	rd->skip '' ---em branco---
	var ie = trim(rd->read)
	if len(ie) = 0 then
		return null
	end if
	if ie[0] < asc("0") or ie[0] > asc("9") then
		return null
	end if
	
	rd->skip '' data emi
	chave = rd->read
	if len(chave) <> 3+44 then
		return null
	end if
	
	chave = right(chave, 44)
	
	var item = new TDFe_NFeItem

	item->modelo 			= 59
	item->numero			= rd->readInt
	rd->skip '' situa��o
	item->serie				= rd->readInt
	item->ICMS				= rd->readDbl
	item->nroItem			= rd->readInt
	item->codProduto		= rd->read
	rd->skip '' EAN
	item->descricao			= rd->read(true)
	item->ncm				= rd->readInt
	item->cfop				= rd->readInt
	item->unidade			= rd->read
	item->qtd				= rd->readDbl
	rd->skip '' Indicador Regra C�lculo
	rd->skip '' Valor Unit�rio Comercializa��o
	rd->skip '' Valor Produtos
	item->desconto			= rd->readDbl
	item->despesasAcess		= rd->readDbl
	item->valorProduto		= rd->readDbl
	rd->skip '' Valor Rateio Desconto
	rd->skip '' Valor Rateio Acrescimo
	rd->skip '' Indicador Origem
	item->cst				= rd->readInt
	item->aliqICMS			= rd->readDbl
	rd->skip '' C�digo CST PIS
	rd->skip '' Valor Base C�lculo PIS
	rd->skip '' Al�quota PIS
	item->IPI				= rd->readDbl
	item->bcICMS			= item->valorProduto
	item->bcICMSST			= 0
	item->next_ = null
	
	function = item
	
end function

''''''''
function Efd.carregarXlsxSAT(rd as ExcelReader ptr) as TDFe ptr
	
	'' ---em branco---, Num Inscr. Estadual Emitente,	N�mero de S�rie do SAT,	Data Emiss�o,	Hora Emiss�o,	
	'' Indicador Cupom Cancelado,	Identifica��o CF-e,	Data Recep��o Cupom,	N�mero Cupom CF-e,	Indicador Possui Destinat�rio,	
	'' Valor Total CF-e,	Valor Total ICMS,	Valor Total Produtos,	Valor Total Desconto,	Valor Total Pis,	Valor Total Cofins,	
	'' Valor Total Pis-ST,	Valor Total Cofins-ST,	Valor Total Outros,	Valor Acrescimo/Desconto Subtotal,	Valor Cfe Lei 12741
	
	rd->skip '' ---em branco---
	var ie = rd->read
	if len(ie) = 0 then
		return null
	end if
	if ie[0] < asc("0") or ie[0] > asc("9") then
		return null
	end if
	
	rd->skip '' N�mero de S�rie do SAT
	var dEmi 				= rd->readDate
	rd->skip '' Hora Emiss�o
	rd->skip '' Indicador Cupom Cancelado
	
	var chave = rd->read
	if len(chave) <> 3+44 then
		return null
	end if
	
	chave = right(chave, 44)
	
	var dfe = cast(TDFe ptr, chaveDFeDict->lookup(chave))	
	if dfe = null then
		dfe = new TDFe
	end if
	
	dfe->chave				= chave
	dfe->dataEmi			= dEmi
	dfe->nfe.ieEmit			= str(cdbl(ie))
	rd->skip '' Data Recep��o Cupom
	dfe->numero				= rd->readInt
	dfe->serie				= 0
	dfe->modelo				= 59
	rd->skip '' Indicador Possui Destinat�rio
	dfe->valorOperacao		= rd->readDbl
	dfe->nfe.ICMSTotal		= rd->readDbl
	dfe->nfe.bcICMSTotal	= dfe->valorOperacao
	dfe->ufEmit				= 35
	dfe->cnpjDest			= "00000000000000"
	dfe->ufDest				= 35
	dfe->operacao			= SAIDA
	dfe->nfe.bcICMSSTTotal	= 0
	dfe->nfe.ICMSSTTotal	= 0
	
	function = dfe
	
end function

''''''''
function Efd.carregarXlsx(nomeArquivo as String) as Boolean

	if left(nomeArquivo, 1) = "~" then
		return true
	elseif left(nomeArquivo, 7) = "SpedEFD" then
		return true
	elseif nomeArquivo = "__efd__.xlsx" then
		return true
	elseif instr(nomeArquivo, "NFe_Destinatario_Itens_OSF") > 0 then
		onProgress(null, 1)
		return true
	end if

	dim as integer tipoArquivo
	dim as string nomePlanilhas(0 to 1)
	
	if instr( nomeArquivo, "NFe_Destinatario_OSF" ) > 0 then
		tipoArquivo = BO_NFe_Dest
		nfeDestSafiFornecido = true
		nomePlanilhas(0) = "Planilha NF-e por Destinatário"

	elseif instr( nomeArquivo, "NFe_Emitente_Itens_OSF" ) > 0 then
		tipoArquivo = BO_NFe_Emit_Itens
		itemNFeSafiFornecido = true
		nomePlanilhas(0) = "Planilha"
	
	elseif instr( nomeArquivo, "NFe_Emitente_OSF" ) > 0 then
		tipoArquivo = BO_NFe_Emit
		nfeEmitSafiFornecido = true
		nomePlanilhas(0) = "Planilha NF-e por Emitente"
	
	elseif instr( nomeArquivo, "CTe_CNPJ_Emitente_Tomador_Remetente_Destinatario_OSF" ) > 0 then
		tipoArquivo = BO_CTe
		nomePlanilhas(0) = "CT-e por Emitente"
		nomePlanilhas(1) = "CT-e por Tomador"
		cteSafiFornecido = true
	
	elseif instr( nomeArquivo, "SAT_-_CuponsEmitidosPorContribuinteCNPJ_OSF" ) > 0 then
		tipoArquivo = BO_SAT
		nfeEmitSafiFornecido = true
		nomePlanilhas(0) = "Cupons emitidos em dado periodo"
	
	elseif instr( nomeArquivo, "SAT_-_ItensDeCuponsCNPJ_OSF" ) > 0 then
		tipoArquivo = BO_SAT_Itens
		itemNFeSafiFornecido = true
		nomePlanilhas(0) = "Itens de Cupons"
	
	elseif instr( nomeArquivo, "NFC-e_itens_OSF" ) > 0 then
		tipoArquivo = BO_NFCe_Itens
		itemNFeSafiFornecido = true
		nomePlanilhas(0) = "Itens"
		onError(!"\n\tErro: relat�rio n�o suportado ainda")
		return false
		
	elseif instr( nomeArquivo, "REDF_consulta_Cupons_Fiscais_ECF" ) > 0 then
		tipoArquivo = SAFI_ECF
		nfeEmitSafiFornecido = true
		nomePlanilhas(0) = "REDF - Cupons Fiscais"
		onError(!"\n\tErro: relat�rio n�o suportado ainda")
		return false
	
	elseif instr( nomeArquivo, "REDF_-_Consulta_Cupons_Fiscais_ECF_e_itens_do_CF" ) > 0 then
		tipoArquivo = BO_ECF_Itens
		itemNFeSafiFornecido = true
		nomePlanilhas(0) = "REDF - Itens dos Cupons Fiscais"
		onError(!"\n\tErro: relat�rio n�o suportado ainda")
		return false
	
	else
		onError(!"\n\tErro: imposs�vel resolver tipo de arquivo pelo nome")
		return false
	end if
	
	var reader = new ExcelReader()
	
	if not reader->open(nomeArquivo) then
		onError(!"\n\tErro: arquivo n�o encontrado ou inv�lido")
		delete reader
		return false
	end if
	
	var plan = 0
	do
		var nomePlanilha = nomePlanilhas(plan)
		if nomePlanilha = "" then
			exit do
		end if
		
		if not reader->setSheet(nomePlanilha) then
			onError(!"\n\tErro: planilha n�o encontrada (" + nomePlanilha + ")")
			delete reader
			return false
		end if
		
		var nroLinha = 1

		try
			do while (reader->nextRow()) 
				if nroLinha > 1 then
					select case as const tipoArquivo  
					case BO_NFe_Dest
						var dfe = carregarXlsxNFeDest(reader)
						if dfe <> null then
							adicionarDFe(dfe)
						end if
					
					case BO_NFe_Emit
						var dfe = carregarXlsxNFeEmit( reader )
						if dfe <> null then
							adicionarDFe(dfe)
						end if
						
					case BO_NFe_Emit_Itens
						var chave = ""
						var nfeItem = carregarXlsxNFeEmitItens( reader, chave )
						if nfeItem <> null then
							adicionarItemDFe(chave, nfeItem)

							var dfe = cast(TDFe ptr, chaveDFeDict->lookup(chave))
							'' nf-e n�o encontrada? pode acontecer se processarmos o csv de itens antes do csv de nf-e
							if dfe = null then
								dfe = new TDFe
								'' s� adicionar ao dicion�rio, depois ser� adicionado por adicionarDFe() no case acima
								dfe->chave = chave
								chaveDFeDict->add(dfe->chave, dfe)
							end if
							
							if dfe->nfe.itemListHead = null then
								dfe->nfe.itemListHead = nfeItem
							else
								dfe->nfe.itemListTail->next_ = nfeItem
							end if
							
							dfe->nfe.itemListTail = nfeItem
						end if
					
					case BO_CTe
						var dfe = carregarXlsxCTe( reader, iif(plan = 0, SAIDA, ENTRADA) )
						if dfe <> null then
							adicionarDFe(dfe)
						end if
						
					case BO_SAT
						var dfe = carregarXlsxSAT( reader )
						if dfe <> null then
							adicionarDFe(dfe)
						end if
						
					case BO_SAT_Itens
						var chave = ""
						var satItem = carregarXlsxSATItens( reader, chave )
						if satItem <> null then
							adicionarItemDFe(chave, satItem)

							var dfe = cast(TDFe ptr, chaveDFeDict->lookup(chave))
							'' sat n�o encontrado? pode acontecer se processarmos o csv de itens antes do csv de nf-e
							if dfe = null then
								dfe = new TDFe
								'' s� adicionar ao dicion�rio, depois ser� adicionado por adicionarDFe() no case acima
								dfe->chave = chave
								chaveDFeDict->add(dfe->chave, dfe)
							end if
							
							if dfe->nfe.itemListHead = null then
								dfe->nfe.itemListHead = satItem
							else
								dfe->nfe.itemListTail->next_ = satItem
							end if
							
							dfe->nfe.itemListTail = satItem
						end if
						
					case BO_NFCe_Itens
						''var dfe = carregarXlsxNFCeItens( reader )
						''if dfe <> null then
						''end if

					case BO_ECF_Itens
						''var dfe = carregarXlsxECFItens( reader )
						''if dfe <> null then
						''end if
						
					end select
				end if
				
				nroLinha += 1
			loop
			
			function = true
		
		catch
			onError(!"\r\n\tErro ao carregar linha " & nroLinha & !"\r\n")
			function = false
		endtry
	
		plan += 1
	loop while plan <= ubound(nomePlanilhas)
	
	onProgress(null, 1)
	
	delete reader
	
end function

''''''''
private sub adicionarColunasComuns(sheet as ExcelWorksheet ptr, ehEntrada as Boolean, itemNFeSafiFornecido as boolean)

	var row = sheet->addRow(true)
	row->addCell("CNPJ " + iif(ehEntrada, "Emitente", "Destinatario"))
	row->addCell("IE " + iif(ehEntrada, "Emitente", "Destinatario"))
	row->addCell("UF " + iif(ehEntrada, "Emitente", "Destinatario"))
	row->addCell("Razao Social " + iif(ehEntrada, "Emitente", "Destinatario"))
	row->addCell("Modelo")
	row->addCell("Serie")
	row->addCell("Numero")
	row->addCell("Data Emissao")
	row->addCell("Data " + iif(ehEntrada, "Entrada", "Saida"))
	row->addCell("Chave")
	row->addCell("Situacao")
	row->addCell("BC ICMS")
	row->addCell("Aliq ICMS")
	row->addCell("Valor ICMS")
	row->addCell("BC ICMS ST")
	row->addCell("Aliq ICMS ST")
	row->addCell("Valor ICMS ST")
	row->addCell("Valor IPI")
	row->addCell("Valor Item")
	row->addCell("Nro Item")
	row->addCell("Qtd")
	row->addCell("Unidade")
	row->addCell("CFOP")
	row->addCell("CST")
	row->addCell("NCM")
	row->addCell("Codigo Item")
	row->addCell("Descricao Item")
	
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING, 45)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 30)
   
	if not ehEntrada then
		row->addCell("DifAl FCP")
		row->addCell("DifAl ICMS Orig")
		row->addCell("DifAl ICMS Dest")
		
		sheet->AddCellType(CT_MONEY)
		sheet->AddCellType(CT_MONEY)
		sheet->AddCellType(CT_MONEY)
	end if
	
	row->addCell("Info. complementares")
	sheet->AddCellType(CT_STRING, 40)
end sub
   
''''''''
private sub lua_setarGlobal overload (lua as lua_State ptr, varName as const zstring ptr, value as integer)
	lua_pushnumber(lua, value)
	lua_setglobal(lua, varName)
end sub

''''''''
private sub lua_setarGlobal overload (lua as lua_State ptr, varName as const zstring ptr, value as any ptr)
	lua_pushlightuserdata(lua, value)
	lua_setglobal(lua, varName)
end sub

private sub criarColunasApuracaoIcms(sheet as ExcelWorksheet ptr)
	var row = sheet->addRow(true)
	row->addCell("Inicio")
	row->addCell("Fim")
	row->addCell("Total Debitos")
	row->addCell("Ajustes Debitos")
	row->addCell("Total Ajuste Deb")
	row->addCell("Estornos Credito")
	row->addCell("Total Creditos")
	row->addCell("Ajustes Creditos")
	row->addCell("Total Ajuste Cred")
	row->addCell("Estornos Debito")
	row->addCell("Saldo Cred Anterior")
	row->addCell("Saldo Devedor Apurado")
	row->addCell("Total Deducoes")
	row->addCell("ICMS a Recolher")
	row->addCell("Saldo Credor a Transportar")
	row->addCell("Deb Extra Apuracao")
	for i as integer = 1 to MAX_AJUSTES
		row->addCell("Detalhe Ajuste " & i)
	next
	
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	for i as integer = 1 to MAX_AJUSTES
		sheet->AddCellType(CT_STRING, 80)
	next
end sub

private sub criarColunasApuracaoIcmsST(sheet as ExcelWorksheet ptr)
	var row = sheet->addRow(true)
	row->addCell("Inicio")
	row->addCell("Fim")
	row->addCell("UF")
	row->addCell("Movimentacao")
	row->addCell("Saldo Credor Anterior")
	row->addCell("Total Devolucao Merc")
	row->addCell("Total Ressarcimentos")
	row->addCell("Total Ajustes Cred")
	row->addCell("Total Ajustes Cred Docs")
	row->addCell("Total Retencao")
	row->addCell("Total Ajustes Deb")
	row->addCell("Total Ajustes Deb Docs")
	row->addCell("Saldo Devedor ant. Deducoes")
	row->addCell("Total Deducoes")
	row->addCell("ICMS a Recolher")
	row->addCell("Saldo Credor a Transportar")
	row->addCell("Deb Extra Apuracao")
	for i as integer = 1 to MAX_AJUSTES
		row->addCell("Detalhe Ajuste " & i)
	next

	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	for i as integer = 1 to MAX_AJUSTES
		sheet->AddCellType(CT_STRING, 80)
	next
end sub

private sub criarColunasInventario(sheet as ExcelWorksheet ptr)
	var row = sheet->addRow(true)
	row->addCell("Data Inventario")
	row->addCell("Codigo")
	row->addCell("NCM")
	row->addCell("Tipo")
	row->addCell("Tipo (Descricao)")
	row->addCell("Descricao")
	row->addCell("Unidade")
	row->addCell("Qtd")
	row->addCell("Valor Unitario")
	row->addCell("Valor Item")
	row->addCell("Ind. Propriedade")
	row->addCell("CNPJ Proprietario")
	row->addCell("Texto Complementar")
	row->addCell("Codigo Conta Contabil")
	row->addCell("Valor Item IR")

	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_MONEY)
end sub

private sub criarColunasCIAP(sheet as ExcelWorksheet ptr)
	var row = sheet->addRow(true)
	row->addCell("Data Inicial")
	row->addCell("Data Final")
	row->addCell("Soma Total Saidas Tributadas")
	row->addCell("Soma Total Saidas")
	row->addCell("Codigo Bem")
	row->addCell("Descricao Bem")
	row->addCell("Data Movimentacao")
	row->addCell("Tipo Movimentacao")
	row->addCell("Valor ICMS")
	row->addCell("Valor ICMS ST")
	row->addCell("Valor ICMS Frete")
	row->addCell("Valor ICMS Difal")
	row->addCell("Num. Parcela")
	row->addCell("Valor Parcela")
	row->addCell("Modelo")
	row->addCell("Serie")
	row->addCell("Numero")
	row->addCell("Data Emissao")
	row->addCell("Chave NF-e")
	row->addCell("CNPJ")
	row->addCell("IE")
	row->addCell("UF")
	row->addCell("Razao Social")
	
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 30)
end sub

private sub criarColunasEstoque(sheet as ExcelWorksheet ptr)
	var row = sheet->addRow(true)
	row->addCell("Data Inicial")
	row->addCell("Data Final")
	row->addCell("Codigo Item")
	row->addCell("NCM Item")
	row->addCell("Tipo Item")
	row->addCell("Tipo Item (Descricao)")
	row->addCell("Descricao Item")
	row->addCell("Qtd")
	row->addCell("Tipo")
	row->addCell("Prop CNPJ")
	row->addCell("Prop IE")
	row->addCell("Prop UF")
	row->addCell("Prop Razao Social")
	
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 30)
end sub

private sub criarColunasProducao(sheet as ExcelWorksheet ptr)
	var row = sheet->addRow(true)
	row->addCell("Data Inicial")
	row->addCell("Data Final")
	row->addCell("Codigo Item")
	row->addCell("NCM Item")
	row->addCell("Tipo Item")
	row->addCell("Tipo Item (Descricao)")
	row->addCell("Descricao Item")
	row->addCell("Qtd")
	row->addCell("Codigo Ordem")
	
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_STRING)
end sub

private sub criarColunasRessarcST(sheet as ExcelWorksheet ptr)
	var row = sheet->addRow(true)
	row->addCell("CNPJ Emitente Ult NF-e Ent")
	row->addCell("IE Emitente Ult NF-e Ent")
	row->addCell("UF Emitente Ult NF-e Ent")
	row->addCell("Razao Social Emitente Ult NF-e Ent")
	row->addCell("Modelo Ult NF-e Ent")
	row->addCell("Serie Ult NF-e Ent")
	row->addCell("Numero Ult NF-e Ent")
	row->addCell("Data Emissao Ult NF-e Ent")
	row->addCell("Qtd Ult Ent")
	row->addCell("Valor Ult Ent")
	row->addCell("BC ICMS ST")
	row->addCell("Chave Ult NF-e Ent")
	row->addCell("Num Item Ult NF-e Ent")
	row->addCell("BC ICMS")
	row->addCell("Aliq ICMS")
	row->addCell("Lim BC ICMS")
	row->addCell("ICMS")
	row->addCell("Aliq ICMS ST")
	row->addCell("Ressarcimento")
	row->addCell("Responsavel")
	row->addCell("Motivo")
	row->addCell("Tipo Doc Arrecad")
	row->addCell("Num Doc Arrecad")
	row->addCell("Chave NF-e Saida")
	row->addCell("Num Item NF-e Saida")
	
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 30)
	sheet->AddCellType(CT_STRING, 4)
	sheet->AddCellType(CT_STRING, 6)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_DATE)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_STRING, 45)
	sheet->AddCellType(CT_INTNUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_NUMBER)
	sheet->AddCellType(CT_MONEY)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING)
	sheet->AddCellType(CT_STRING, 45)
	sheet->AddCellType(CT_INTNUMBER)
end sub

''''''''
sub Efd.criarPlanilhas()
	'' planilha de entradas
	entradas = ew->AddWorksheet("Entradas")
	adicionarColunasComuns(entradas, true, itemNFeSafiFornecido)

	'' planilha de sa�das
	saidas = ew->AddWorksheet("Saidas")
	adicionarColunasComuns(saidas, false, itemNFeSafiFornecido)

	'' apura��o do ICMS
	apuracaoIcms = ew->AddWorksheet("Apuracao ICMS")
	criarColunasApuracaoIcms(apuracaoIcms)
   
	'' apura��o do ICMS ST
	apuracaoIcmsST = ew->AddWorksheet("Apuracao ICMS ST")
	criarColunasApuracaoIcmsST(apuracaoIcmsST)
	
	'' Invent�rio
	inventario = ew->AddWorksheet("Inventario")
	criarColunasInventario(inventario)

	'' CIAP
	ciap = ew->AddWorksheet("CIAP")
	criarColunasCIAP(ciap)

	'' Estoque
	estoque = ew->AddWorksheet("Estoque")
	criarColunasEstoque(estoque)

	'' Producao
	producao = ew->AddWorksheet("Producao")
	criarColunasProducao(producao)

	'' Ressarcimento ST
	ressarcST = ew->AddWorksheet("Ressarcimento ST")
	criarColunasRessarcST(ressarcST)
	
	'' Inconsistencias LRE
	inconsistenciasLRE = ew->AddWorksheet("Inconsistencias LRE")

	'' Inconsistencias LRS
	inconsistenciasLRS = ew->AddWorksheet("Inconsistencias LRS")
	
	'' Resumos LRE
	resumosLRE = ew->AddWorksheet("Resumos LRE")

	'' Resumos LRS
	resumosLRS = ew->AddWorksheet("Resumos LRS")
	
	''
	lua_getglobal(lua, "criarPlanilhas")
	lua_call(lua, 0, 0)

	lua_setarGlobal(lua, "efd_plan_entradas", entradas)
	lua_setarGlobal(lua, "efd_plan_saidas", saidas)
	
end sub

''''''''
type HashCtx
	bf				as bfile ptr
	tamanhoSemSign	as longint
	bytesLidosTotal	as longint
end type

private function HashReadCB cdecl(ctx_ as any ptr, buffer as ubyte ptr, maxLen as long) as long
	var ctx = cast(HashCtx ptr, ctx_)
	
	if ctx->bytesLidosTotal + maxLen > ctx->tamanhoSemSign then
		maxLen = ctx->tamanhoSemSign - ctx->bytesLidosTotal
	end if
	
	var bytesLidos = ctx->bf->ler(buffer, maxLen)
	ctx->bytesLidosTotal += bytesLidos
	
	function = bytesLidos
	
end function

''''''''
function Efd.lerInfoAssinatura(nomeArquivo as string, assinaturaP7K_DER() as byte) as InfoAssinatura ptr
	
	try
		var res = new InfoAssinatura
		
		var sh = new SSL_Helper
		var tamanhoAssinatura = ubound(assinaturaP7K_DER)+1
		var p7k = sh->Load_P7K(@assinaturaP7K_DER(0), tamanhoAssinatura)
		
		''
		var s = sh->Get_CommonName(p7k)
		if s <> null then
			res->assinante = *s
			deallocate s
		end if
		
		''
		s = sh->Get_AttributeFromAltName(p7k, AN_ATT_CPF)
		if s <> null then
			res->cpf = *s
			deallocate s
		else
			res->cpf = "00000000000"
		end if

		''
		var bf = new bfile()
		bf->abrir(nomeArquivo)
		var ctx = new HashCtx
		ctx->bf = bf
		ctx->tamanhoSemSign = bf->tamanho() - (tamanhoAssinatura + len(ASSINATURA_P7K_HEADER))
		ctx->bytesLidosTotal = 0
		
		s = sh->Compute_SHA1(@HashReadCB, ctx)
		if s <> null then
			res->hashDoArquivo = *s
			deallocate s
		end if
		
		bf->fechar()

		''
		sh->Free(p7k)
		delete sh
		
		function = res
	catch
		onError("Erro ao ler assinatura digital. As informa��es relativas � assinatura estar�o em branco nos relat�rios gerados")
		function = null
	endtry
	
end function

''''''''
function Efd.processar(nomeArquivo as string) as Boolean
   
	if opcoes.formatoDeSaida <> FT_NULL then
		gerarPlanilhas(nomeArquivo)
	else
		onProgress(null, 1)
	end if
	
	if opcoes.gerarRelatorios then
		if tipoArquivo = TIPO_ARQUIVO_EFD then
			infAssinatura = lerInfoAssinatura(nomeArquivo, assinaturaP7K_DER())
		
			gerarRelatorios(nomeArquivo)
			
			if infAssinatura <> NULL then
				delete infAssinatura
			end if
		end if
	end if
	
	do while regListHead <> null
		var next_ = regListHead->next_
		delete regListHead
		regListHead = next_
	loop

	regListHead = null

	delete sintegraDict
	delete infoComplDict
	delete bemCiapDict
	delete itemIdDict
	delete participanteDict

	function = true
end function

private function efd.getInfoCompl(info as TDocInfoCompl ptr) as string
	var res = ""
	
	do while info <> null
		var compl = cast( TInfoCompl ptr, infoComplDict->lookup(info->idCompl))
		res += iif(len(res) > 0, "|", "") + _
			compl->descricao + _
			iif(len(info->extra) > 0, ":" + info->extra, "")
		info = info->next_
	loop
	
	function = res
end function

''''''''
sub Efd.gerarPlanilhas(nomeArquivo as string)
	
	if entradas = null then
		criarPlanilhas
	end if
	
	onProgress(!"\tGerando planilhas", 0)
	
	dim as TRegistro ptr reg = null
	try
		var regCnt = 0
		reg = regListHead
		do while reg <> null
			'para cada registro..
			select case as const reg->tipo
			'item de NF-e?
			case DOC_NF_ITEM
				var doc = reg->itemNF.documentoPai
				select case as const doc->situacao
				case REGULAR, EXTEMPORANEO
					var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipante) )

					var emitirLinha = iif(doc->operacao = SAIDA, not opcoes.pularLrs, not opcoes.pularLre)
					if opcoes.filtrarCnpj andalso emitirLinha then
						if part <> null then
							emitirLinha = filtrarPorCnpj(part->cnpj)
						end if
					end if

					if opcoes.filtrarChaves andalso emitirLinha then
						emitirLinha = filtrarPorChave(doc->chave)
					end if
					
					if opcoes.somenteRessarcimentoST andalso emitirLinha then
						emitirLinha = reg->itemNF.itemRessarcStListHead <> null
					end if
					
					if emitirLinha then
						's� existe item para entradas (exceto quando h� ressarcimento ST)
						dim as ExcelRow ptr row
						if doc->operacao = ENTRADA then
							row = entradas->AddRow()
						else
							row = saidas->AddRow()
						end if

						if part <> null then
							row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
							row->addCell(part->ie)
							row->addCell(MUNICIPIO2SIGLA(part->municip))
							row->addCell(part->nome)
						else
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell("")
						end if
						row->addCell(doc->modelo)
						row->addCell(doc->serie)
						row->addCell(doc->numero)
						row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
						row->addCell(YyyyMmDd2Datetime(doc->dataEntSaida))
						row->addCell(doc->chave)
						row->addCell(codSituacao2Str(doc->situacao))
						row->addCell(reg->itemNF.bcICMS)
						row->addCell(reg->itemNF.aliqICMS)
						row->addCell(reg->itemNF.ICMS)
						row->addCell(reg->itemNF.bcICMSST)
						row->addCell(reg->itemNF.aliqICMSST)
						row->addCell(reg->itemNF.ICMSST)
						row->addCell(reg->itemNF.IPI)
						row->addCell(reg->itemNF.valor)
						row->addCell(reg->itemNF.numItem)
						row->addCell(reg->itemNF.qtd)
						row->addCell(reg->itemNF.unidade)
						row->addCell(reg->itemNF.cfop)
						row->addCell(reg->itemNF.cstICMS)
						var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->itemNF.itemId) )
						if itemId <> null then 
							row->addCell(itemId->ncm)
							row->addCell(itemId->id)
							row->addCell(itemId->descricao)
						end if
						row->addCell(getInfoCompl(doc->infoComplListHead))
					end if
				end select

			'NF-e?
			case DOC_NF, DOC_NFSCT, DOC_NF_ELETRIC
				select case as const reg->nf.situacao
				case REGULAR, EXTEMPORANEO
					'' NOTA: n�o existe itemDoc para sa�das (exceto quando h� ressarcimento ST), s� temos informa��es b�sicas do DF-e, 
					'' 	     a n�o ser que sejam carregados os relat�rios .csv do SAFI vindos do infoview
					if reg->nf.operacao = SAIDA or (reg->nf.operacao = ENTRADA and reg->nf.nroItens = 0) or reg->tipo <> DOC_NF then
						dim as TDFe_NFeItem ptr item = null
						if itemNFeSafiFornecido and opcoes.acrescentarDados then
							if len(reg->nf.chave) > 0 then
								var dfe = cast( TDFe ptr, chaveDFeDict->lookup(reg->nf.chave) )
								if dfe <> null then
									item = dfe->nfe.itemListHead
								end if
							end if
						end if

						var part = cast( TParticipante ptr, participanteDict->lookup(reg->nf.idParticipante) )

						var emitirLinhas = (opcoes.somenteRessarcimentoST = false) and _
							iif(reg->nf.operacao = SAIDA, not opcoes.pularLrs, not opcoes.pularLre)
						if opcoes.filtrarCnpj andalso emitirLinhas then
							if part <> null then
								emitirLinhas = filtrarPorCnpj(part->cnpj)
							end if
						end if

						if opcoes.filtrarChaves andalso emitirLinhas then
							emitirLinhas = filtrarPorChave(reg->nf.chave)
						end if

						var anal = iif(item = null, reg->nf.itemAnalListHead, null)
						var analCnt = 1
						
						if emitirLinhas then
							do
								dim as ExcelRow ptr row
								if reg->nf.operacao = SAIDA then
									row = saidas->AddRow()
								else
									row = entradas->AddRow()
								end if
							
								if part <> null then
									row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
									row->addCell(part->ie)
									row->addCell(MUNICIPIO2SIGLA(part->municip))
									row->addCell(part->nome)
								else
									row->addCell("")
									row->addCell("")
									row->addCell("")
									row->addCell("")
								end if
								row->addCell(reg->nf.modelo)
								row->addCell(reg->nf.serie)
								row->addCell(reg->nf.numero)
								row->addCell(YyyyMmDd2Datetime(reg->nf.dataEmi))
								row->addCell(YyyyMmDd2Datetime(reg->nf.dataEntSaida))
								row->addCell(reg->nf.chave)
								row->addCell(codSituacao2Str(reg->nf.situacao))

								if ((itemNFeSafiFornecido and opcoes.acrescentarDados) or _
								   cbool((reg->nf.operacao = ENTRADA) and (reg->tipo = DOC_NF))) and _
								   cbool(item <> null) then
									row->addCell(item->bcICMS)
									row->addCell(item->aliqICMS)
									row->addCell(item->ICMS)
									row->addCell(item->bcICMSST)
									row->addCell("")
									row->addCell("")
									row->addCell(item->IPI)
									row->addCell(item->valorProduto)
									row->addCell(item->nroItem)
									row->addCell(item->qtd)
									row->addCell(item->unidade)
									row->addCell(item->cfop)
									row->addCell(item->cst)
									row->addCell(item->ncm)
									row->addCell(item->codProduto)
									row->addCell(item->descricao)

								else
									if anal = null then
										row->addCell(reg->nf.bcICMS)
										row->addCell("")
										row->addCell(reg->nf.ICMS)
										row->addCell(reg->nf.bcICMSST)
										row->addCell(reg->nf.ICMSST)
										row->addCell("")
										row->addCell(reg->nf.IPI)
										row->addCell(reg->nf.valorTotal)
										for cell as integer = 1 to 16-8
											row->addCell("")
										next
									else
										row->addCell(anal->bc)
										row->addCell(anal->aliq)
										row->addCell(anal->ICMS)
										row->addCell(anal->bcST)
										row->addCell("")
										row->addCell(anal->ICMSST)
										row->addCell(anal->IPI)
										row->addCell(anal->valorOp)
										row->addCell(analCnt)
										row->addCell(0)
										row->addCell("")
										row->addCell(anal->cfop)
										row->addCell(anal->cst)
										for cell as integer = 1 to 3
											row->addCell("")
										next
										analCnt += 1
									end if
								end if

								if reg->nf.operacao = SAIDA then
									row->addCell(reg->nf.difal.fcp)
									row->addCell(reg->nf.difal.icmsOrigem)
									row->addCell(reg->nf.difal.icmsDest)
								end if
								
								row->addCell(getInfoCompl(reg->nf.infoComplListHead))
							
								if item = null then
									if anal = null then
										exit do
									end if
									
									anal = anal->next_
								else
									item = item->next_
								end if
								
							loop while (item <> null) or (anal <> null)
						end if
					
					end if
			   
				case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
					if reg->nf.operacao = SAIDA then
						var emitirLinha = (opcoes.somenteRessarcimentoST = false) and _
							(not opcoes.pularLrs)
						
						if emitirLinha then
							var row = saidas->AddRow()

							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell(reg->nf.modelo)
							row->addCell(reg->nf.serie)
							row->addCell(reg->nf.numero)
							'' NOTA: cancelados e inutilizados n�o v�m com a data preenchida, ent�o retiramos a data da chave ou do registro mestre
							var dataEmi = iif( len(reg->nf.chave) = 44, "20" + mid(reg->nf.chave,3,2) + mid(reg->nf.chave,5,2) + "01", regMestre->mestre.dataIni )
							row->addCell(YyyyMmDd2Datetime(dataEmi))
							row->addCell("")
							row->addCell(reg->nf.chave)
							row->addCell(codSituacao2Str(reg->nf.situacao))
						end if
					end if

				end select
				
			'ressarcimento st?
			case DOC_NF_ITEM_RESSARC_ST
				var doc = @reg->itemRessarcSt
				var part = cast( TParticipante ptr, participanteDict->lookup(doc->idParticipanteUlt) )

				var emitirLinha = iif(reg->ct.operacao = SAIDA, not opcoes.pularLrs, not opcoes.pularLre)
				if opcoes.filtrarCnpj andalso emitirLinha then
					if part <> null then
						emitirLinha = filtrarPorCnpj(part->cnpj)
					end if
				end if

				if emitirLinha then
					var row = ressarcST->AddRow()

					if part <> null then
						row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
						row->addCell(part->ie)
						row->addCell(MUNICIPIO2SIGLA(part->municip))
						row->addCell(part->nome)
					else
						row->addCell("")
						row->addCell("")
						row->addCell("")
						row->addCell("")
					end if
					row->addCell(doc->modeloUlt)
					row->addCell(doc->serieUlt)
					row->addCell(doc->numeroUlt)
					row->addCell(YyyyMmDd2Datetime(doc->dataUlt))
					row->addCell(doc->qtdUlt)
					row->addCell(doc->valorUlt)
					row->addCell(doc->valorBcST)
					row->addCell(doc->chaveNFeUlt)
					row->addCell(doc->numItemNFeUlt)
					row->addCell(doc->bcIcmsUlt)
					row->addCell(doc->aliqIcmsUlt)
					row->addCell(doc->limiteBcIcmsUlt)
					row->addCell(doc->icmsUlt)
					row->addCell(doc->aliqIcmsStUlt)
					row->addCell(doc->res)
					row->addCell(doc->responsavelRet)
					row->addCell(doc->motivo)
					row->addCell(doc->tipDocArrecadacao)
					row->addCell(doc->numDocArrecadacao)
					row->addCell(doc->documentoPai->documentoPai->chave)
					row->addCell(doc->documentoPai->numItem)
				end if

			'CT-e?
			case DOC_CT
				select case as const reg->ct.situacao
				case REGULAR, EXTEMPORANEO
					var part = cast( TParticipante ptr, participanteDict->lookup(reg->ct.idParticipante) )

					var emitirLinhas = (opcoes.somenteRessarcimentoST = false) and _
						iif(reg->ct.operacao = SAIDA, not opcoes.pularLrs, not opcoes.pularLre)
					
					if opcoes.filtrarCnpj andalso emitirLinhas then
						if part <> null then
							emitirLinhas = filtrarPorCnpj(part->cnpj)
						end if
					end if

					if opcoes.filtrarChaves andalso emitirLinhas then
						emitirLinhas = filtrarPorChave(reg->ct.chave)
					end if
						
					if emitirLinhas then
						dim as TDFe_CTe ptr cte = null
						if cteSafiFornecido then
							if len(reg->ct.chave) > 0 then
								var dfe = cast( TDFe ptr, chaveDFeDict->lookup(reg->ct.chave) )
								if dfe <> null then
									cte = @dfe->cte
								end if
							end if
						end if
						
						dim as TDocItemAnal ptr item = null
						if reg->ct.operacao = ENTRADA then
							item = reg->ct.itemAnalListHead
						end if
						
						var itemCnt = 1
						do
							dim as ExcelRow ptr row 
							if reg->ct.operacao = SAIDA then
								row = saidas->AddRow()
							else
								row = entradas->AddRow()
							end if
							
							if part <> null then
								row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
								row->addCell(part->ie)
								row->addCell(MUNICIPIO2SIGLA(part->municip))
								row->addCell(part->nome)
							else
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
							end if
							row->addCell(reg->ct.modelo)
							row->addCell(reg->ct.serie)
							row->addCell(reg->ct.numero)
							row->addCell(YyyyMmDd2Datetime(reg->ct.dataEmi))
							row->addCell(YyyyMmDd2Datetime(reg->ct.dataEntSaida))
							row->addCell(reg->ct.chave)
							row->addCell(codSituacao2Str(reg->ct.situacao))
							
							if item <> null then
								row->addCell(item->bc)
								row->addCell(item->aliq)
								row->addCell(item->ICMS)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell(item->valorOp)
								row->addCell(itemCnt)
								row->addCell("")
								row->addCell("")
								row->addCell(item->cfop)
								row->addCell(item->cst)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								
								item = item->next_
								itemCnt += 1
							else
								row->addCell(reg->ct.bcICMS)
								row->addCell("")
								row->addCell(reg->ct.ICMS)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell(reg->ct.valorServico)
								row->addCell(1)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								
							end if

							if reg->ct.operacao = SAIDA then
								row->addCell(reg->ct.difal.fcp)
								row->addCell(reg->ct.difal.icmsOrigem)
								row->addCell(reg->ct.difal.icmsDest)
							end if
							
						loop while item <> null
					end if
				
				case CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
					if reg->ct.operacao = SAIDA then
						if opcoes.somenteRessarcimentoST = false then
							var row = saidas->AddRow()

							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell("")
							row->addCell(reg->ct.modelo)
							row->addCell(reg->ct.serie)
							row->addCell(reg->ct.numero)
							'' NOTA: cancelados e inutilizados n�o v�m com a data preenchida, ent�o retiramos a data da chave ou do registro mestre
							var dataEmi = iif( len(reg->ct.chave) = 44, "20" + mid(reg->ct.chave,3,2) + mid(reg->ct.chave,5,2) + "01", regMestre->mestre.dataIni )
							row->addCell(YyyyMmDd2Datetime(dataEmi))
							row->addCell("")
							row->addCell(reg->ct.chave)
							row->addCell(codSituacao2Str(reg->ct.situacao))
						end if
					end if
				
				end select
				
			'item de ECF?
			case DOC_ECF_ITEM
				if not opcoes.pularLrs then
					var doc = reg->itemECF.documentoPai
					select case as const doc->situacao
					case REGULAR, EXTEMPORANEO
						's� existe cupom para sa�da
						if doc->operacao = SAIDA then
							var emitirLinha = (opcoes.somenteRessarcimentoST = false)
							if opcoes.filtrarCnpj andalso emitirLinha then
								emitirLinha = filtrarPorCnpj(doc->cpfCnpjAdquirente)
							end if

							if opcoes.filtrarChaves andalso emitirLinha then
								emitirLinha = filtrarPorChave(doc->chave)
							end if
							
							if emitirLinha then
								var row = saidas->AddRow()

								row->addCell(doc->cpfCnpjAdquirente)
								row->addCell("")
								row->addCell("SP")
								row->addCell(doc->nomeAdquirente)
								row->addCell(iif(doc->modelo = &h2D, "2D", str(doc->modelo)))
								row->addCell("")
								row->addCell(doc->numero)
								row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
								row->addCell(YyyyMmDd2Datetime(doc->dataEntSaida))
								row->addCell(doc->chave)
								row->addCell(codSituacao2Str(doc->situacao))
								row->addCell("")
								row->addCell(reg->itemECF.aliqICMS)
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell("")
								row->addCell(reg->itemECF.valor)
								row->addCell(reg->itemECF.numItem)
								row->addCell(reg->itemECF.qtd)
								row->addCell(reg->itemECF.unidade)
								row->addCell(reg->itemECF.cfop)
								row->addCell(reg->itemECF.cstICMS)
								var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->itemECF.itemId) )
								if itemId <> null then 
									row->addCell(itemId->ncm)
									row->addCell(itemId->id)
									row->addCell(itemId->descricao)
								end if
							end if
						end if
					end select
				end if
				
			'SAT?
			case DOC_SAT
				if not opcoes.pularLrs then
					var doc = @reg->sat
					select case as const doc->situacao
					case REGULAR, EXTEMPORANEO
						's� existe cupom para sa�da
						if doc->operacao = SAIDA then
							var emitirLinha = (opcoes.somenteRessarcimentoST = false)
							if opcoes.filtrarCnpj andalso emitirLinha then
								emitirLinha = filtrarPorCnpj(doc->cpfCnpjAdquirente)
							end if
							
							if opcoes.filtrarChaves andalso emitirLinha then
								emitirLinha = filtrarPorChave(doc->chave)
							end if
							
							if emitirLinha then
								dim as TDFe_NFeItem ptr item = null
								if itemNFeSafiFornecido and opcoes.acrescentarDados then
									var dfe = cast( TDFe ptr, chaveDFeDict->lookup(doc->chave) )
									if dfe <> null then
										item = dfe->nfe.itemListHead
									end if
								end if
								
								var anal = iif(item = null, doc->itemAnalListHead, null)
								
								var analCnt = 1
								do
									var row = saidas->AddRow()

									row->addCell(doc->cpfCnpjAdquirente)
									row->addCell("")
									row->addCell("SP")
									row->addCell("")
									row->addCell(str(doc->modelo))
									row->addCell("")
									row->addCell(doc->numero)
									row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
									row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
									row->addCell(doc->chave)
									row->addCell(codSituacao2Str(doc->situacao))
									if item <> null then
										row->addCell(item->bcICMS)
										row->addCell(item->aliqICMS)
										row->addCell(item->ICMS)
										row->addCell("")
										row->addCell("")
										row->addCell("")
										row->addCell("")
										row->addCell(item->valorProduto)
										row->addCell(item->nroItem)
										row->addCell(item->qtd)
										row->addCell(item->unidade)
										row->addCell(item->cfop)
										row->addCell(item->cst)
										row->addCell(item->ncm)
										row->addCell(item->codProduto)
										row->addCell(item->descricao)
										
										item = item->next_
										if item = null then
											exit do
										end if
										
									else
										if anal = null then
											exit do
										end if
											
										row->addCell("")
										row->addCell(anal->aliq)
										row->addCell("")
										row->addCell("")
										row->addCell("")
										row->addCell("")
										row->addCell("")
										row->addCell(anal->valorOp)
										row->addCell(analCnt)
										row->addCell("")
										row->addCell("")
										row->addCell(anal->cfop)
										row->addCell(anal->cst)
										row->addCell("")
										row->addCell("")
										row->addCell("")
										
										analCnt += 1
										anal = anal->next_
										if anal = null then
											exit do
										end if
									end if
								loop
							end if
						end if
					end select
				end if
				
			case APURACAO_ICMS_PERIODO
				if not opcoes.pularLraicms then
					var row = apuracaoIcms->AddRow()

					row->addCell(YyyyMmDd2Datetime(reg->apuIcms.dataIni))
					row->addCell(YyyyMmDd2Datetime(reg->apuIcms.dataFim))
					row->addCell(reg->apuIcms.totalDebitos)
					row->addCell(reg->apuIcms.ajustesDebitos)
					row->addCell(reg->apuIcms.totalAjusteDeb)
					row->addCell(reg->apuIcms.estornosCredito)
					row->addCell(reg->apuIcms.totalCreditos)
					row->addCell(reg->apuIcms.ajustesCreditos)
					row->addCell(reg->apuIcms.totalAjusteCred)
					row->addCell(reg->apuIcms.estornoDebitos)
					row->addCell(reg->apuIcms.saldoCredAnterior)
					row->addCell(reg->apuIcms.saldoDevedorApurado)
					row->addCell(reg->apuIcms.totalDeducoes)
					row->addCell(reg->apuIcms.icmsRecolher)
					row->addCell(reg->apuIcms.saldoCredTransportar)
					row->addCell(reg->apuIcms.debExtraApuracao)
					
					var detalhe = ""
					var ajuste = reg->apuIcms.ajustesListHead
					var cnt = 1
					do while ajuste <> null andalso cnt <= MAX_AJUSTES
						row->addCell("{'codigo':'" & ajuste->codigo & "', 'valor':'" & DBL2MONEYBR(ajuste->valor) & "', 'descricao':'" & ajuste->descricao) & "'}"
						ajuste = ajuste->next_
						cnt += 1
					loop
					
				end if
				
			case APURACAO_ICMS_ST_PERIODO
				if not opcoes.pularLraicms then
					var row = apuracaoIcmsST->AddRow()

					row->addCell(YyyyMmDd2Datetime(reg->apuIcmsST.dataIni))
					row->addCell(YyyyMmDd2Datetime(reg->apuIcmsST.dataFim))
					row->addCell(reg->apuIcmsST.UF)
					row->addCell(iif(reg->apuIcmsST.mov=0, "N", "S"))
					row->addCell(reg->apuIcmsST.saldoCredAnterior)
					row->addCell(reg->apuIcmsST.devolMercadorias)
					row->addCell(reg->apuIcmsST.totalRessarciment)
					row->addCell(reg->apuIcmsST.totalOutrosCred)
					row->addCell(reg->apuIcmsST.ajustesCreditos)
					row->addCell(reg->apuIcmsST.totalRetencao)
					row->addCell(reg->apuIcmsST.totalOutrosDeb)
					row->addCell(reg->apuIcmsST.ajustesDebitos)
					row->addCell(reg->apuIcmsST.saldoAntesDed)
					row->addCell(reg->apuIcmsST.totalDeducoes)
					row->addCell(reg->apuIcmsST.icmsRecolher)
					row->addCell(reg->apuIcmsST.saldoCredTransportar)
					row->addCell(reg->apuIcmsST.debExtraApuracao)
				end if

			case INVENTARIO_ITEM
				var row = inventario->AddRow()
				
				row->addCell(YyyyMmDd2Datetime(reg->invItem.dataInventario))

				var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->invItem.itemId) )
				if itemId <> null then 
					row->addCell(itemId->id)
					row->addCell(itemId->ncm)
					row->addCell(itemId->tipoItem)
					row->addCell(tipoItem2Str(itemId->tipoItem))
					row->addCell(itemId->descricao)
				else
					row->addCell(reg->invItem.itemId)
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if
				
				row->addCell(reg->invItem.unidade)
				row->addCell(reg->invItem.qtd)
				row->addCell(reg->invItem.valorUnitario)
				row->addCell(reg->invItem.valorItem)
				row->addCell(reg->invItem.indPropriedade)
				var part = cast( TParticipante ptr, participanteDict->lookup(reg->invItem.idParticipante) )
				if part <> null then
					row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
				else
					row->addCell("")
				end if
				row->addCell(reg->invItem.txtComplementar)
				row->addCell(reg->invItem.codConta)
				row->addCell(reg->invItem.valorItemIR)

			case CIAP_ITEM
				if reg->ciapItem.docCnt = 0 then
					var row = ciap->AddRow()
					
					var pai = reg->ciapItem.pai
					row->addCell(YyyyMmDd2Datetime(pai->dataIni))
					row->addCell(YyyyMmDd2Datetime(pai->dataFim))
					row->addCell(pai->valorTributExpSoma)
					row->addCell(pai->valorTotalSaidas)
					
					var bemCiap = cast( TBemCiap ptr, bemCiapDict->lookup(reg->ciapItem.bemId) )
					if bemCiap <> null then 
						row->addCell(bemCiap->id)
						row->addCell(bemCiap->descricao)
					else
						row->addCell(reg->ciapItem.bemId)
						row->addCell("")
					end if
					
					row->addCell(YyyyMmDd2Datetime(reg->ciapItem.dataMov))
					row->addCell(reg->ciapItem.tipoMov)
					row->addCell(reg->ciapItem.valorIcms)
					row->addCell(reg->ciapItem.valorIcmsSt)
					row->addCell(reg->ciapItem.valorIcmsFrete)
					row->addCell(reg->ciapItem.valorIcmsDifal)
					row->addCell(reg->ciapItem.parcela)
					row->addCell(reg->ciapItem.valorParcela)
				end if

			case CIAP_ITEM_DOC
				var row = ciap->AddRow()
				
				var pai = reg->ciapItemDoc.pai
				var avo = pai->pai
				row->addCell(YyyyMmDd2Datetime(avo->dataIni))
				row->addCell(YyyyMmDd2Datetime(avo->dataFim))
				row->addCell(avo->valorTributExpSoma)
				row->addCell(avo->valorTotalSaidas)
				
				var bemCiap = cast( TBemCiap ptr, bemCiapDict->lookup(pai->bemId) )
				if bemCiap <> null then 
					row->addCell(bemCiap->id)
					row->addCell(bemCiap->descricao)
				else
					row->addCell(pai->bemId)
					row->addCell("")
				end if
				
				row->addCell(YyyyMmDd2Datetime(pai->dataMov))
				row->addCell(pai->tipoMov)
				row->addCell(pai->valorIcms)
				row->addCell(pai->valorIcmsSt)
				row->addCell(pai->valorIcmsFrete)
				row->addCell(pai->valorIcmsDifal)
				row->addCell(pai->parcela)
				row->addCell(pai->valorParcela)
				
				row->addCell(reg->ciapItemDoc.modelo)
				row->addCell(reg->ciapItemDoc.serie)
				row->addCell(reg->ciapItemDoc.numero)
				row->addCell(YyyyMmDd2Datetime(reg->ciapItemDoc.dataEmi))
				row->addCell(reg->ciapItemDoc.chaveNfe)
				
				var part = cast( TParticipante ptr, participanteDict->lookup(reg->ciapItemDoc.idParticipante) )
				if part <> null then
					row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
					row->addCell(part->ie)
					row->addCell(MUNICIPIO2SIGLA(part->municip))
					row->addCell(part->nome)
				else
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if

			case ESTOQUE_ITEM
				var row = estoque->AddRow()
				
				var pai = reg->estItem.pai
				row->addCell(YyyyMmDd2Datetime(pai->dataIni))
				row->addCell(YyyyMmDd2Datetime(pai->dataFim))
				
				var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->estItem.itemId) )
				if itemId <> null then 
					row->addCell(itemId->id)
					row->addCell(itemId->ncm)
					row->addCell(itemId->tipoItem)
					row->addCell(tipoItem2Str(itemId->tipoItem))
					row->addCell(itemId->descricao)
				else
					row->addCell(reg->estItem.itemId)
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if
				
				row->addCell(reg->estItem.qtd)
				row->addCell(reg->estItem.tipoEst)

				var part = cast( TParticipante ptr, participanteDict->lookup(reg->estItem.idParticipante) )
				if part <> null then
					row->addCell(iif(len(part->cpf) > 0, part->cpf, part->cnpj))
					row->addCell(part->ie)
					row->addCell(MUNICIPIO2SIGLA(part->municip))
					row->addCell(part->nome)
				else
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if

			case ESTOQUE_ORDEM_PROD
				var row = producao->AddRow()
				
				row->addCell(YyyyMmDd2Datetime(reg->estOrdem.dataIni))
				row->addCell(YyyyMmDd2Datetime(reg->estOrdem.dataFim))
				
				var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->estOrdem.itemId) )
				if itemId <> null then 
					row->addCell(itemId->id)
					row->addCell(itemId->ncm)
					row->addCell(itemId->tipoItem)
					row->addCell(tipoItem2Str(itemId->tipoItem))
					row->addCell(itemId->descricao)
				else
					row->addCell(reg->estOrdem.itemId)
					row->addCell("")
					row->addCell("")
					row->addCell("")
					row->addCell("")
				end if
				
				row->addCell(reg->estOrdem.qtd)
				row->addCell(reg->estOrdem.idOrdem)

			'item de documento do sintegra?
			case SINTEGRA_DOCUMENTO_ITEM
				if not opcoes.pularLrs then
					var doc = reg->docItemSint.doc
					
					select case as const doc->situacao
					case REGULAR, EXTEMPORANEO, CANCELADO, CANCELADO_EXT, DENEGADO, INUTILIZADO
						dim as ExcelRow ptr row 
						if doc->operacao = SAIDA then
							row = saidas->AddRow()
						else
							row = entradas->AddRow()
						end if
						
						var itemId = cast( TItemId ptr, itemIdDict->lookup(reg->docItemSint.codMercadoria) )
						  
						row->addCell(doc->cnpj)
						row->addCell(doc->ie)
						row->addCell(ufCod2Sigla(doc->uf))
						row->addCell("")
						row->addCell(doc->modelo)
						row->addCell(doc->serie)
						row->addCell(doc->numero)
						row->addCell(YyyyMmDd2Datetime(doc->dataEmi))
						row->addCell("")
						row->addCell("")
						row->addCell(codSituacao2Str(doc->situacao))
						row->addCell(reg->docItemSint.bcICMS)
						row->addCell(reg->docItemSint.aliqICMS)
						row->addCell(reg->docItemSint.bcICMS * reg->docItemSint.aliqICMS / 100)
						row->addCell(reg->docItemSint.bcICMSST)
						row->addCell("")
						row->addCell(reg->docSint.ICMSST)
						row->addCell(reg->docItemSint.valorIPI)
						row->addCell(reg->docItemSint.valor)
						row->addCell(reg->docItemSint.nroItem)
						row->addCell(reg->docItemSint.qtd)
						if itemId <> null then 
							row->addCell(rtrim(itemId->unidInventario))
						else
							row->addCell("")
						end if
						row->addCell(reg->docItemSint.cfop)
						row->addCell(reg->docItemSint.cst)
						if itemId <> null then 
							row->addCell(itemId->ncm)
							row->addCell(rtrim(itemId->id))
							row->addCell(rtrim(itemId->descricao))
						end if
						
					end select
				end if

			case LUA_CUSTOM
				
				var luaFunc = cast(customLuaCb ptr, customLuaCbDict->lookup(reg->lua.tipo))->writer
				
				if luaFunc <> null then
					lua_getglobal(lua, luaFunc)
					lua_rawgeti(lua, LUA_REGISTRYINDEX, reg->lua.table)
					lua_call(lua, 1, 0)
				end if
			
			end select

			regCnt += 1
			onProgress(null, regCnt / nroRegs)
			
			reg = reg->next_
		loop
	catch
		onError(!"\r\nErro ao tratar o registro de tipo (" & reg->tipo & !") carregado na linha (" & reg->linha & !")\r\n")
	endtry
	
	onProgress(null, 1)
	
end sub

''''''''
function EFd.getPlanilha(nome as const zstring ptr) as ExcelWorksheet ptr
		dim as ExcelWorksheet ptr plan = null
		select case lcase(*nome)
		case "entradas"
			plan = entradas
		case "saidas"
			plan = saidas
		case "inconsistencias LRE"
			plan = inconsistenciasLRE
		case "inconsistencias LRS"
			plan = inconsistenciasLRS
		case "resumos LRE"
			plan = resumosLRE
		case "resumos LRS"
			plan = resumosLRS
		case "ciap"
			plan = ciap
		case "estoque"
			plan = estoque
		case "producao"
			plan = producao
		case "inventario"
			plan = inventario
		end select
		function = plan
end function

''''''''
private function luacb_efd_plan_get cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)
	
	lua_getglobal(L, "efd")
	var g_efd = cast(Efd ptr, lua_touserdata(L, -1))
	lua_pop(L, 1)
	
	if args = 1 then
		var planName = lua_tostring(L, 1)

		var plan = g_efd->getPlanilha(planName)
		if plan <> null then
			lua_pushlightuserdata(L, plan)
		else
			lua_pushnil(L)
		end if
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
static function Efd.luacb_efd_participante_get cdecl(byval L as lua_State ptr) as long
	var args = lua_gettop(L)

	lua_getglobal(L, "efd")
	var g_efd = cast(Efd ptr, lua_touserdata(L, -1))
	lua_pop(L, 1)
	
	if args = 2 then
		var idParticipante = lua_tostring(L, 1)
		var formatar = lua_toboolean(L, 2) <> 0

		var part = cast( TParticipante ptr, g_efd->participanteDict->lookup(idParticipante) )
		if part <> null then
			lua_newtable(L)
			lua_pushstring(L, "cnpj")
			lua_pushstring(L, iif(formatar, iif(len(part->cpf) > 0, STR2CPF(part->cpf), STR2CNPJ(part->cnpj)), iif(len(part->cpf) > 0, part->cpf, part->cnpj)))
			lua_settable(L, -3)
			lua_pushstring(L, "ie")
			lua_pushstring(L, iif(formatar, STR2IE(part->ie), part->ie))
			lua_settable(L, -3)
			lua_pushstring(L, "uf")
			lua_pushstring(L, MUNICIPIO2SIGLA(part->municip))
			lua_settable(L, -3)
			lua_pushstring(L, "municip")
			lua_pushstring(L, g_efd->codMunicipio2Nome(part->municip))
			lua_settable(L, -3)			
			lua_pushstring(L, "nome")
			lua_pushstring(L, iif(formatar, left(part->nome, 64), part->nome))
			lua_settable(L, -3)
		else
			lua_pushnil(L)
		end if
	else
		 lua_pushnil(L)
	end if
	
	function = 1
	
end function

''''''''
sub Efd.exportAPI(L as lua_State ptr)
	
	lua_setarGlobal(L, "TI_ESCRIT_FALTA", TI_ESCRIT_FALTA)
	lua_setarGlobal(L, "TI_ESCRIT_FANTASMA", TI_ESCRIT_FANTASMA)
	lua_setarGlobal(L, "TI_ALIQ", TI_ALIQ)
	lua_setarGlobal(L, "TI_DUP", TI_DUP)
	lua_setarGlobal(L, "TI_DIF", TI_DIF)
	lua_setarGlobal(L, "TI_RESSARC_ST", TI_RESSARC_ST)
	lua_setarGlobal(L, "TI_CRED", TI_CRED)
	lua_setarGlobal(L, "TI_SEL", TI_SEL)
	lua_setarGlobal(L, "TI_DEB", TI_DEB)
	
	lua_setarGlobal(L, "TL_ENTRADAS", TL_ENTRADAS)
	lua_setarGlobal(L, "TL_SAIDAS", TL_SAIDAS)

	lua_setarGlobal(L, "TR_CFOP", TR_CFOP)
	lua_setarGlobal(L, "TR_CST", TR_CST)
	lua_setarGlobal(L, "TR_CST_CFOP", TR_CST_CFOP)

	lua_setarGlobal(L, "DFE_NFE_DEST_FORNECIDO", MASK_BO_NFe_DEST_FORNECIDO)
	lua_setarGlobal(L, "DFE_NFE_EMIT_FORNECIDO", MASK_BO_NFe_EMIT_FORNECIDO)
	lua_setarGlobal(L, "DFE_ITEM_NFE_FORNECIDO", MASK_BO_ITEM_NFE_FORNECIDO)
	lua_setarGlobal(L, "DFE_CTE_FORNECIDO", MASK_BO_CTe_FORNECIDO)
	
	lua_setarGlobal(L, "efd", @this)
	
	lua_register(L, "efd_plan_get", @luacb_efd_plan_get)
	lua_register(L, "efd_participante_get", @luacb_efd_participante_get)
	lua_register(L, "efd_rel_addItemAnalitico", @luacb_efd_rel_addItemAnalitico)
	
end sub
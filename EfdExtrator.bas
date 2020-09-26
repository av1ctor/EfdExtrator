'' Extrator de EFD
'' Copyleft 2017-2020 André Vicentini (avtvicentini)
'' fbc.exe EfdExtrator.bas Efd.bas Efd-analises.bas Efd-relatorios.bas Efd-misc.bas bfile.bas ExcelReader.bas ExcelWriter.bas list.bas Dict.bas Pdfer.bas DB.bas VarBox.bas trycatch.bas -d WITH_PARSER

#include once "EFD.bi"

declare sub main()
declare sub importarGia()
declare sub importarCadContribuinte()
declare sub importarCadContribuinteRegime()
declare sub importarCadInidoneo()

on error goto exceptionReport

'''''''''''
sub mostrarUso()
	print wstr("Modo de usar:")
	print wstr("EfdExtrator.exe Opções efd-ou-sintegra.txt [relatorio-bo.csv] [relatorio-bo.xlsx]")
	print wstr("Notas:")
	print wstr(!" 1. No lugar do nome dos arquivos, podem ser usadas máscaras,")
	print wstr(!"    como por exemplo: *.txt *.csv *.xlsx")
	print wstr(!" 2. Os arquivos .txt podem ser em formato Sintegra ou EFD")
	print wstr(!" 3. Os arquivos .csv do SAFI devem manter o padrão de nome dado pelo")
	print wstr(!"    Infoview BO, mas podem ser usados prefixos e sufixos no nome,")
	print wstr(!"    como por exemplo: \"2017 SAFI_NFe_Emitente_Itens parte 1.csv\"")
	print wstr(!" 4. Os arquivos .xlsx devem manter o padrão de nome dado pelo")
	print wstr(!"    Infoview BO, mas podem ser usados prefixos e sufixos no nome,")
	print wstr(!"    como por exemplo: \"2019 NFe_Emitente_Itens_OSF parte 1.xlsx\"")
	print wstr(!" 5. No final da extração será gerado um arquivo .xlsx para ser aberto")
	print wstr(!"    no Excel 2003 ou superior (exceto se o formato de saída for null)")
	print wstr("Opções:")
	print wstr(!" -gerarRelatorios:")
	print wstr(!"  Gera os relatórios do EFD-ICMS-IPI no formato PDF.")
	print wstr(!" -filtrarCnpjs cnpj1,cnpj2,...:")
	print wstr(!"  Extrai somente os registros com os mesmos CNPJs (de emitentes ou")
	print wstr(!"  destinatários) dos contidos na lista de CNPJs informada (separada por")
	print wstr(!"  vírgula; zeros à esq).")
	print wstr(!" -filtrarChaves chave1,chave2,... ou @arquivo:")
	print wstr(!"  Extrai somente os registros com as mesmas chaves das contidas na lista")
	print wstr(!"  (utilizar @arquivo.txt para carregar as chaves de um arquivo, com uma")
	print wstr(!"  chave por linha, sem linhas vazias ou espaços entre as chaves).")
	print wstr(!" -realcar:")
	print wstr(!"  Cria um realce, nos relatórios em PDF, nos registros que corresponderem")
	print wstr(!"  à -filtrarCnpjs ou -filtrarChaves.")
	print wstr(!" -naoGerarLre, -naoGerarLrs e -naoGerarLraicms:")
	print wstr(!"  Deixam de gerar os respectivos livros quando -gerarRelatorios é utilizada.")
	print wstr(!" -formatoDeSaida xml|csv|xlsx|null:")
	print wstr(!"  Altera o formato de saída do padrão xlsx para csv ou XML.")
	print wstr(!" -complementarDados:")
	print wstr(!"  Inclui dados complementares na planilha (aba Saídas ou Entradas para docs")
	print wstr(!"  de emissão própria) que será gerada e que não constam na EFD, caso os")
	print wstr(!"  arquivos .csv do SAFI ou os .xlsx do Infoview BO sejam fornecidos. AVISO:")
	print wstr(!"  não utilize as informações relacionadas ao ICMS (alíquota, BC, valor, etc),")
	print wstr(!"  pois esses dados serão retirados dos DF-e's e não da escrituração.")
	print wstr(!" -somenteRessarcimentoST:")
	print wstr(!"  Extrai somente documentos do LRS que contenham o registro C176 relativo ao")
	print wstr(!"  ressarcimento ST.")
	print wstr(!" -dbEmDisco:")
	print wstr(!"  Grava os dados intermediários em disco, poupando memória.")
	print wstr(!" -manterDB:")
	print wstr(!"  Preserva o arquivo de dados intermediários (formato SQLite3).")
	
	print 
end sub

'''''''''''   
sub mostrarCopyright()
	print wstr("Extrator de EFD/Sintegra para Excel, versão 0.8 beta")
	print wstr("Copyleft 2017-2020 by André Vicentini (avtvicentini)")
	print
end sub

'''''''''''
sub mostrarProgresso(estagio as const wstring ptr, porCompleto as double)
	static as double ultPorCompleto = 0
	
	if estagio <> null then
		print *estagio;
	end if
	
	if porCompleto = 0 then
		ultPorCompleto = 0
		return
	end if
	
	do while porCompleto >= ultPorCompleto + 0.05
		print ".";
		ultPorCompleto += 0.05
	loop
	
	if porCompleto = 1 then
		print "OK!"
	end if
	
end sub

'''''''''''
sub main()
	dim as OpcoesExtracao opcoes
	
	mostrarCopyright()
   
	if len(command(1)) = 0 then
		mostrarUso()
		exit sub
	end if
   
	'' verificar opções
	var nroOpcoes = 0
	var i = 1
	do 
		var arg = command(i)
		if len(arg) = 0 then
			exit do
		end if
		
		if arg[0] = asc("-") then
			select case lcase(arg)
			case "-gerarrelatorios"
				opcoes.gerarRelatorios = true
				nroOpcoes += 1
			case "-naogerarlre"
				opcoes.pularLreAoGerarRelatorios = true
				nroOpcoes += 1
			case "-naogerarlrs"
				opcoes.pularLrsAoGerarRelatorios = true
				nroOpcoes += 1
			case "-naogerarlrelrs"
				opcoes.pularLreAoGerarRelatorios = true
				opcoes.pularLrsAoGerarRelatorios = true
				nroOpcoes += 1
			case "-naogerarlraicms"
				opcoes.pularLRaicmsAoGerarRelatorios = true
				nroOpcoes += 1
			case "-realcar"
				opcoes.highlight = true
			case "-filtrarcnpjs"
				i += 1
				var listaCnpj = command(i)
				if( len(listaCnpj) > 0 ) then
					splitstr(listaCnpj, ",", opcoes.listaCnpj())
					opcoes.filtrarCnpj = true
				else
					opcoes.filtrarCnpj = false
				end if
				nroOpcoes += 2
			case "-filtrarchaves"
				i += 1
				var listaChaves = command(i)
				if( len(listaChaves) > 0 ) then
					if left(listaChaves, 1) = "@" then
						var lista = mid(listaChaves, 2)
						if not loadstrings(lista, opcoes.listaChaves()) then
							print wstr("Erro: ao carregar arquivo: " + lista)
							exit sub
						end if
					else
						splitstr(listaChaves, ",", opcoes.listaChaves())
					end if
					opcoes.filtrarChaves = true
				else
					opcoes.filtrarChaves = false
				end if
				nroOpcoes += 2
			case "-complementardados"
				opcoes.acrescentarDados = true
				nroOpcoes += 1
			case "-importargia"
				importarGia()
				exit sub
			case "-importarcadcontribuinte"
				importarCadContribuinte()
				exit sub
			case "-importarcadregime"
				importarCadContribuinteRegime()
				exit sub
			case "-importarcadinidoneo"
				importarCadInidoneo()
				exit sub
			case "-formatodesaida"
				i += 1
				select case command(i)
				case "xml" 
					opcoes.formatoDeSaida = FT_XML
				case "csv"
					opcoes.formatoDeSaida = FT_CSV
				case "xlsx"
					opcoes.formatoDeSaida = FT_XLSX
				case "null"
					opcoes.formatoDeSaida = FT_NULL
				case else
					print wstr("Erro: formato de saída inválido")
					exit sub
				end select
				nroOpcoes += 2
			case "-somenteressarcimentost"
				opcoes.somenteRessarcimentoST = true
				nroOpcoes += 1
			case "-dbemdisco"
				opcoes.dbEmDisco = true
				nroOpcoes += 1
			case "-manterdb"
				opcoes.manterDb = true
				opcoes.dbEmDisco = true
				nroOpcoes += 1
			case else
				print wstr("Erro: opção inválida: " + arg)
				exit sub
			end select
		end if
		
		i += 1
	loop
	
	dim as Efd e
	
	'' 
	var arquivoSaida = iif( len(command(nroOpcoes+2)) > 0, "__efd__", command(nroOpcoes+1))
   
	e.iniciarExtracao(arquivoSaida, opcoes)
   
	'' mais de um arquivo informado?
	if len(command(nroOpcoes+2)) > 0 then
	   '' carregar arquivos .csv primeiro com dados de NF-e e CT-e 
	   var i = nroOpcoes+1
	   var arquivoEntrada = command(i)
	   do while len(arquivoEntrada) > 0
			if lcase(right(arquivoEntrada,3)) = "csv" then
				if not e.carregarCsv( arquivoEntrada, @mostrarProgresso ) then
					print !"\r\nErro ao carregar arquivo: "; arquivoEntrada
					end -1
				end if
			elseif lcase(right(arquivoEntrada,4)) = "xlsx" then
				if not e.carregarXlsx( arquivoEntrada, @mostrarProgresso ) then
					print !"\r\nErro ao carregar arquivo: "; arquivoEntrada
					end -1
				end if
			end if 

			i += 1
			arquivoEntrada = command(i)
	   loop
   
	   '' carregar arquivos .txt com EFD ou Sintegra
	   i = nroOpcoes+1
	   arquivoEntrada = command(i)
	   do while len(arquivoEntrada) > 0
			if lcase(right(arquivoEntrada,3)) = "txt" then
				if not e.carregarTxt( arquivoEntrada, @mostrarProgresso ) then
					print !"\r\nErro ao carregar arquivo: "; arquivoEntrada
					end -1
				end if
				
				print "Processando:"
				if not e.processar( arquivoEntrada, @mostrarProgresso ) then
					print !"\r\nErro ao extrair arquivo: "; arquivoEntrada
					end -1
				end if
			end if 
			 
			i += 1
			arquivoEntrada = command(i)
	   loop
	   
	'' só um arquivo .txt informado..
	else
		var arquivoEntrada = command(nroOpcoes+1)
		if not e.carregarTxt( arquivoEntrada, @mostrarProgresso ) then
			print !"\r\nErro ao carregar arquivo: "; arquivoEntrada
			end -1
		end if
	
		print "Processando:"
		if not e.processar( arquivoEntrada, @mostrarProgresso ) then
			print !"\r\nErro ao extrair arquivo: "; arquivoEntrada
			end -1
		end if
	end if
	
	''
	if opcoes.formatoDeSaida <> FT_NULL then
		print "Analisando:"
		e.analisar(@mostrarProgresso)
	end if
   
	''
	e.finalizarExtracao( @mostrarProgresso )
	
end sub

  
'''''''''''
sub importarGia()   

	const SEP = asc("|")
	
	var db = new TDb
	
	db->open(ExePath + "\db\GIA.db")
	db->execNonQuery("PRAGMA JOURNAL_MODE=OFF")
	db->execNonQuery("PRAGMA SYNCHRONOUS=0")
	db->execNonQuery("PRAGMA LOCKING_MODE=EXCLUSIVE")
	
	var stmt = db->prepare("insert into GIA (ie, mes, ano, totCreditos, totDebitos) values (?,?,?,?,?)")
	var updStmt = db->prepare("update GIA set totDevolucoes = ?, totRetencoes = ? where ie = ? and mes = ? and ano = ?")
	
	var i = 2
	do
		var arquivo = command(i)
		if len(arquivo) = 0 then
			exit do
		end if
		
		dim as bfile inf
		if not inf.abrir(arquivo) then
			print wstr("Erro: ao carregar arquivo: " + arquivo)
			exit do
		end if
		
		'' encontrar ano na 1a linha
		inf.varint(SEP)
		inf.varint(SEP)
		inf.varint(SEP)
		inf.varint(SEP)
		var ano = inf.int4
		inf.char2			'' skip \r\n
		
		mostrarProgresso("Carregando GIA(" & arquivo & ")", 0)
		
		'' remover todos os registros desse ano
		db->execNonQuery("delete from GIA where ano = " & ano)
		
		var arqTamanho = inf.tamanho
		var l = 0
		do while inf.temProximo()
			
			if l = 0 then
				db->execNonQuery("begin")
			end if
			
			'' carregar cada registro
			'' formato: IE¨mês¨indICMS¨(totDebitos¨totCreditos|totDevolucoes|totRetencoes)\r\n
			var ie = inf.varint(SEP)
			var mes = inf.varint(SEP)
			var icmsSt = inf.varint(SEP)
			'' icms próprio?
			if icmsSt = 0 then
				var totDebitos = inf.vardbl(SEP, asc("."))
				var totCreditos = inf.vardbl(13, asc("."))
				stmt->reset()
				stmt->bind(1, ie)
				stmt->bind(2, mes)
				stmt->bind(3, ano)
				stmt->bind(4, totCreditos)
				stmt->bind(5, totDebitos)
				db->execNonQuery(stmt)
			'' st..
			else
				var totDevolucoes = inf.vardbl(SEP, asc("."))
				var totRetencoes = inf.vardbl(13, asc("."))
				updStmt->reset()
				updStmt->bind(1, totDevolucoes)
				updStmt->bind(2, totRetencoes)
				updStmt->bind(3, ie)
				updStmt->bind(4, mes)
				updStmt->bind(5, ano)
				db->execNonQuery(updStmt)
			end if
			
			inf.char1			'' skip \n
			
			mostrarProgresso(0, inf.posicao / arqTamanho)
			
			if l = 100000 then
				db->execNonQuery("end")
				l = -1
			end if
			
			l += 1
		loop

		if l > 0 then
			mostrarProgresso(0, 1)
			db->execNonQuery("end")
		end if
		
		inf.fechar()
		
		i += 1
	loop
	
	db->close()
	
end sub

private function brdata2yyyymmdd(s as const zstring ptr) as string
	dim as string res = "yyyymmdd"
	
	var i = 0
	if s[i+1] = asc("/") then
		res[6] = asc("0")
		res[7] = s[i]
		i += 2
	else
		res[6] = s[i]
		res[7] = s[i+1]
		i += 3
	end if
	
	if s[i+1] = asc("/") then
		res[4] = asc("0")
		res[5] = s[i]
		i += 2
	else
		res[4] = s[i]
		res[5] = s[i+1]
		i += 3
	end if
	
	res[0] = s[i]
	res[1] = s[i+1]
	res[2] = s[i+2]
	res[3] = s[i+3]
	
	function = res

end function

'''''''''''
sub importarCadContribuinte()   

	const SEP = asc("|")
	
	var arquivo = command(2)
	if len(arquivo) = 0 then
		return
	end if
		
	dim as bfile inf
	if not inf.abrir(arquivo) then
		print wstr("Erro: ao carregar arquivo: " + arquivo)
		return
	end if
	
	var db = new TDb
	
	db->open(ExePath + "\db\CadContribuinte.db")
	db->execNonQuery("PRAGMA JOURNAL_MODE=OFF")
	db->execNonQuery("PRAGMA SYNCHRONOUS=0")
	db->execNonQuery("PRAGMA LOCKING_MODE=EXCLUSIVE")
	
	'' pular as 2 primeiras linhas
	inf.varchar(10)
	inf.varchar(10)
	
	mostrarProgresso("Carregando Cadastro Contribuinte (" & arquivo & ")", 0)
	
	'' remover todos os registros
	db->execNonQuery("delete from Contribuinte")
	
	var stmt = db->prepare("insert into Contribuinte (cnpj, ie, dataIni, dataFim, codBaixa, cnae) values (?,?,?,?,?,?)")
	
	var arqTamanho = inf.tamanho
	var l = 0
	do while inf.temProximo()
		
		if l = 0 then
			db->execNonQuery("begin")
		end if
		
		'' carregar cada registro
		'' formato: CNPJ¨IE¨Nome¨DataIni¨DataFim¨CodBaixa¨Cnae\r\n
		var cnpj = inf.varint(SEP)
		var ie = inf.varint(SEP)
		var nome = inf.varchar(SEP)
		var dataIni = inf.varchar(SEP)
		var dataFim = inf.varchar(SEP)
		var codBaixa = inf.varint(SEP)
		var cnae = inf.varint(13)
		inf.char1			'' skip \n
			
		dataIni = brdata2yyyymmdd(dataIni)
		
		if dataFim = "31/12/1899" then
			dataFim = "99999999"
		else
			dataFim = brdata2yyyymmdd(dataFim)
		end if
		
		stmt->reset()
		stmt->bind(1, cnpj)
		stmt->bind(2, ie)
		stmt->bind(3, dataIni)
		stmt->bind(4, dataFim)
		stmt->bind(5, codBaixa)
		stmt->bind(6, cnae)
		db->execNonQuery(stmt)
		
		if l = 100000 then
			mostrarProgresso(0, inf.posicao / arqTamanho)
			db->execNonQuery("end")
			l = -1
		end if
		
		l += 1
	loop
	
	if l > 0 then
		mostrarProgresso(0, 1)
		db->execNonQuery("end")
	end if
	
	inf.fechar()
	
	db->close()
	
end sub

'''''''''''
sub importarCadContribuinteRegime()   

	''dim lastDayOfMonth(1 to 12) as string = {"31", "28", "31", "30", "31", "30", "31", "31", "30", "31", "30", "31"}

	const SEP = asc("|")
	
	var arquivo = command(2)
	if len(arquivo) = 0 then
		return
	end if
		
	dim as bfile inf
	if not inf.abrir(arquivo) then
		print wstr("Erro: ao carregar arquivo: " + arquivo)
		return
	end if
	
	var db = new TDb
	
	db->open(ExePath + "\db\CadContribuinte.db")
	db->execNonQuery("PRAGMA JOURNAL_MODE=OFF")
	db->execNonQuery("PRAGMA SYNCHRONOUS=0")
	db->execNonQuery("PRAGMA LOCKING_MODE=EXCLUSIVE")
	
	'' pular a primeira linha
	inf.varchar(10)
	
	mostrarProgresso("Carregando Cadastro Regimes (" & arquivo & ")", 0)
	
	'' remover todos os registros
	db->execNonQuery("delete from Regimes")
	
	var stmt = db->prepare("insert into Regimes (ie, tipo, dataIni, dataFim) values (?,?,?,?)")
	
	var arqTamanho = inf.tamanho
	var l = 0
	do while inf.temProximo()
		
		if l = 0 then
			db->execNonQuery("begin")
		end if
		
		'' carregar cada registro
		'' formato: IE¨Regime¨DataIni¨DataFim\r\n
		var ie = inf.varint(SEP)
		var regime = inf.varchar(SEP)
		var dataIni = inf.varchar(SEP)
		var dataFim = inf.varchar(13)
		inf.char1			'' skip \n
			
		stmt->reset()
		stmt->bind(1, ie)
		stmt->bind(2, regime)
		stmt->bind(3, dataIni) '' yyyymm
		stmt->bind(4, dataFim) '' yyyymm
		db->execNonQuery(stmt)
		
		if l = 100000 then
			mostrarProgresso(0, inf.posicao / arqTamanho)
			db->execNonQuery("end")
			l = -1
		end if
		
		l += 1
	loop
	
	if l > 0 then
		mostrarProgresso(0, 1)
		db->execNonQuery("end")
	end if
	
	inf.fechar()
	
	db->close()
	
end sub

'''''''''''
sub importarCadInidoneo()   

	const SEP = asc(";")
	
	var arquivo = command(2)
	if len(arquivo) = 0 then
		return
	end if
		
	dim as bfile inf
	if not inf.abrir(arquivo) then
		print wstr("Erro: ao carregar arquivo: " + arquivo)
		return 
	end if
	
	var db = new TDb
	
	db->open(ExePath + "\db\Inidoneos.db")
	db->execNonQuery("PRAGMA JOURNAL_MODE=OFF")
	db->execNonQuery("PRAGMA SYNCHRONOUS=0")
	db->execNonQuery("PRAGMA LOCKING_MODE=EXCLUSIVE")
	
	'' pular a primeira linha
	inf.varchar(10)
	
	mostrarProgresso("Carregando Cadastro de Inidôneos (" & arquivo & ")", 0)
	
	'' remover todos os registros
	db->execNonQuery("delete from Inidoneos")
	
	var stmt = db->prepare("insert into Inidoneos (ie, cnpj, nome, dataof, dtinidon, uf, cod_inid) values (?,?,?,?,?,?,?)")
	
	var arqTamanho = inf.tamanho
	var l = 0
	do while inf.temProximo()
		
		if l = 0 then
			db->execNonQuery("begin")
		end if
		
		'' carregar cada registro
		'' formato: "OFICIO";"NUMORD";"IE";"CGC";"NOME";"dataof";"dtinidon";"UF";"OCORRENCIAS";"ENDERECO";"MUNICIPIO";"PROCESSO"
		inf.charCsv(SEP)
		inf.intCsv(SEP, 0)
		var ie = vallng(inf.charCsv(SEP))
		var cnpj = vallng(inf.charCsv(SEP))
		var nome = inf.charCsv(SEP)
		var dataoficio = valint(left(csvDate2YYYYMMDD(inf.charCsv(SEP, 0)), 4+2+2))
		var datainicio = valint(left(csvDate2YYYYMMDD(inf.charCsv(SEP, 0)), 4+2+2))
		var uf = inf.charCsv(SEP)
		inf.varchar(13)
		inf.char1 '' skip \n
		
		if ie <> 0 andalso cnpj <> 0 then
			stmt->reset()
			stmt->bind(1, ie)
			stmt->bind(2, cnpj)
			stmt->bind(3, nome)
			stmt->bind(4, dataoficio) '' yyyymm
			stmt->bind(5, datainicio) '' yyyymm
			stmt->bind(6, uf)
			stmt->bind(7, 0)
			db->execNonQuery(stmt)
		end if
		
		if l = 100000 then
			mostrarProgresso(0, inf.posicao / arqTamanho)
			db->execNonQuery("end")
			l = -1
		end if
		
		l += 1
	loop
	
	if l > 0 then
		mostrarProgresso(0, 1)
		db->execNonQuery("end")
	end if
	
	inf.fechar()
	
	db->close()
	
end sub

main()
end 0

exceptionReport:
	print wstr(!"\r\nErro não tratado (" & Err & ") no módulo(" & *Ermn & ") na função(" & *Erfn & ") na linha (" & erl & !")\r\n")
	end 1
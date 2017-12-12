'' para compilar: fbc.exe EfdExtrator.bas Efd.bas bfile.bas ExcelWriter.bas list.bas Dict.bas DocxFactoryDyn.bas DB.bas

#include once "EFD.bi"

declare sub main()

'''''''''''
sub mostrarUso()
	print wstr("Modo de usar:")
	print wstr("EfdExtrator.exe [-gerarRelatorios] arquivo.txt [arquivo.csv]")
	print wstr("Notas:")
	print wstr(!"\t1. No lugar do nome dos arquivos, podem ser usadas máscaras,")
	print wstr(!"\t   como por exemplo: *.txt e *.csv")
	print wstr(!"\t2. O(s) arquivo(s) .txt pode(m) ser em formato Sintegra ou EFD")
	print wstr(!"\t3. Os arquivos .csv do SAFI devem manter o padrão de nome dado pelo")
	print wstr(!"\t   Infoview BO, mas podem ser usados prefixos e sufixos no nome,")
	print wstr(!"\t   como por exemplo: \"2017 SAFI_NFe_Emitente_Itens parte 1.csv\"")
	print wstr(!"\t4. No final da extração será gerado um arquivo .xml que deve ser")
	print wstr(!"\t   aberto no Excel 2003 ou superior")
	print wstr(!"\t5. A opção -gerarRelatorios gera os relatórios do EFD-ICMS-IPI")
	print wstr(!"\t   no formato Word/docx")
	print 
end sub

'''''''''''   
sub mostrarCopyright()
	print "Extrator de EFD/Sintegra para Excel"
	print wstr("Copyleft 2017 by André Vicentini (avtvicentini)")
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
	var gerarRelatorios = false
	
	mostrarCopyright()
   
	if len(command(1)) = 0 then
		mostrarUso()
		exit sub
	end if
   
	dim as Efd e
	
	'' verificar opções
	var nroOpcoes = 0
	var i = 1
	do 
		var arg = command(i)
		if len(arg) = 0 then
			exit do
		end if
		
		if arg[0] = asc("-") then
			if arg = "-gerarRelatorios" then
				gerarRelatorios = true
				nroOpcoes += 1
			else
				mostrarUso()
				exit sub
			end if
		end if
		
		i += 1
	loop
	
	'' 
	var arquivoSaida = iif( len(command(nroOpcoes+2)) > 0, "__efd__", command(nroOpcoes+1))
   
	e.iniciarExtracao(arquivoSaida + ".xml")
   
	'' mais de um arquivo informado?
	if len(command(nroOpcoes+2)) > 0 then
	   '' carregar arquivos .csv primeiro com dados de NF-e e CT-e 
	   var i = nroOpcoes+1
	   var arquivoEntrada = command(i)
	   do while len(arquivoEntrada) > 0
			if lcase(right(arquivoEntrada,3)) = "csv" then
				if not e.carregarCsv( arquivoEntrada, @mostrarProgresso ) then
					print !"\r\nErro ao carregar arquivo!"
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
					print !"\r\nErro ao carregar arquivo!"
					end -1
				end if
				
				print "Processando:"
				if not e.processar( arquivoEntrada, @mostrarProgresso, gerarRelatorios ) then
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
		if not e.processar( arquivoEntrada, @mostrarProgresso, gerarRelatorios ) then
			print !"\r\nErro ao extrair arquivo: "; arquivoEntrada
			end -1
		end if
	end if
	
	''
	print "Analisando:"
	e.analisar(@mostrarProgresso)
   
	''
	e.finalizarExtracao( @mostrarProgresso )
	
end sub
   

main()
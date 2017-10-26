'' para compilar: fbc.exe EfdExtrator.bas Efd.bas bfile.bas ExcelWriter.bas list.bas hash.bas

#include once "EFD.bi"

declare sub main()

'''''''''''
sub mostrarUso()
	print wstr("Modo de usar:")
	print wstr("EfdExtrator.exe [-gerarPDF] arquivo.txt [arquivo.csv]")
	print wstr("Notas:")
	print wstr(!"\t1. No lugar do nome dos arquivos, podem ser usadas m�scaras,")
	print wstr(!"\t   como por exemplo: *.txt e *.csv")
	print wstr(!"\t2. O(s) arquivo(s) .txt pode(m) ser em formato Sintegra ou EFD")
	print wstr(!"\t3. Os arquivos .csv do SAFI devem manter o padr�o de nome dado pelo")
	print wstr(!"\t   Infoview BO, mas podem ser usados prefixos e sufixos no nome,")
	print wstr(!"\t   como por exemplo: \"2017 SAFI_NFe_Emitente_Itens parte 1.csv\"")
	print wstr(!"\t4. No final da extra��o ser� gerado um arquivo .xml que deve ser")
	print wstr(!"\t   aberto no Excel 2003 ou superior")
	print wstr(!"\t5. A op��o -gerarPDF ir gerar os relat�rios no formato do EFD-ICMS-IPI")
	print 
end sub

'''''''''''   
sub mostrarCopyright()
	print "Extrator de EFD/Sintegra para Excel"
	print wstr("Copyleft 2017 by Andr� Vicentini (avtvicentini)")
	print
end sub

	dim shared ultPorCompleto as double = 0

sub mostrarProgresso(porCompleto as double)
	do while porCompleto >= ultPorCompleto + 0.05
		print ".";
		ultPorCompleto += 0.05
	loop
end sub

'''''''''''
sub main()
	var gerarPDF = false
	
	mostrarCopyright()
   
	if len(command(1)) = 0 then
		mostrarUso()
		exit sub
	end if
   
	dim e as Efd
	
	'' verificar op��es
	var nroOpcoes = 0
	var i = 1
	do 
		var arg = command(i)
		if len(arg) = 0 then
			exit do
		end if
		
		if arg[0] = asc("-") then
			if arg = "-gerarPDF" then
				gerarPDF = true
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
				print "Carregando arquivo " + arquivoEntrada;
				
				ultPorCompleto = 0
				if not e.carregarCsv( arquivoEntrada, @mostrarProgresso ) then
					print !"\r\nErro ao carregar arquivo!"
					end -1
				end if
				print "OK!"
			end if 

			i += 1
			arquivoEntrada = command(i)
	   loop
   
	   '' carregar arquivos .txt com EFD ou Sintegra
	   i = nroOpcoes+1
	   arquivoEntrada = command(i)
	   do while len(arquivoEntrada) > 0
			if lcase(right(arquivoEntrada,3)) = "txt" then
				print "Carregando arquivo " + arquivoEntrada;

				ultPorCompleto = 0
				if not e.carregarTxt( arquivoEntrada, @mostrarProgresso ) then
					print !"\r\nErro ao carregar arquivo!"
					end -1
				end if
				print "OK!"
				
				print "Processando";
				ultPorCompleto = 0
				if not e.processar( @mostrarProgresso, gerarPDF ) then
					print !"\r\nErro ao extrair arquivo: "; arquivoEntrada
					end -1
				end if
				print "OK!"
			end if 
			 
			i += 1
			arquivoEntrada = command(i)
	   loop
	   
	'' s� um arquivo .txt informado..
	else
		print "Carregando arquivo";
		ultPorCompleto = 0
		var arquivoEntrada = command(nroOpcoes+1)
		if not e.carregarTxt( arquivoEntrada, @mostrarProgresso ) then
			print !"\r\nErro ao carregar arquivo: "; arquivoEntrada
			end -1
		end if
		print "OK!"
	
		print "Processando";
		ultPorCompleto = 0
		if not e.processar( @mostrarProgresso, gerarPDF ) then
			print !"\r\nErro ao extrair arquivo: "; arquivoEntrada
			end -1
		end if
		print "OK!"
	end if
	
	''
	print wstr("Realizando cruzamentos e an�lises");
	ultPorCompleto = 0
	e.analisar(@mostrarProgresso)
	print "OK!"
   
	''
	print wstr("Gravando arquivo de sa�da: "); arquivoSaida + ".xml";
	ultPorCompleto = 0
	e.finalizarExtracao( @mostrarProgresso )
	print "OK!"
	
end sub
   

main()
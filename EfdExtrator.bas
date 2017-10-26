'' para compilar: fbc.exe EfdExtrator.bas Efd.bas bfile.bas ExcelWriter.bas list.bas hash.bas

#include once "EFD.bi"
#define null 0

declare sub main()

'''''''''''
sub mostrarUso()
	print wstr("Modo de usar:")
	print wstr("EfdExtrator.exe arquivo.txt [arquivo.csv]")
	print wstr("Notas:")
	print wstr(!"\t1. No lugar do nome dos arquivos, podem ser usadas máscaras,")
	print wstr(!"\t   como por exemplo: *.txt e *.csv")
	print wstr(!"\t2. O(s) arquivo(s) .txt pode(m) ser em formato Sintegra ou EFD")
	print wstr(!"\t3. Os arquivos .csv do SAFI devem manter o padrão de nome dado pelo")
	print wstr(!"\t   Infoview BO, mas podem ser usados prefixos e sufixos no nome,")
	print wstr(!"\t   como por exemplo: \"2017 SAFI_NFe_Emitente_Itens parte 1.csv\"")
	print wstr(!"\t4. No final da extração será gerado um arquivo .xml que deve ser")
	print wstr(!"\t   aberto no Excel 2003 ou superior")
end sub

'''''''''''   
sub mostrarCopyright()
	print "Extrator de EFD/Sintegra para Excel"
	print wstr("Copyleft 2017 by André Vicentini (avtvicentini)")
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
	mostrarCopyright()
   
	if len(command(1)) = 0 then
		mostrarUso()
		exit sub
	end if
   
	dim e as Efd
   
	var arquivoSaida = iif( len(command(2)) > 0, "__efd__", command(1))
   
	e.iniciarExtracao(arquivoSaida + ".xml")
   
	if len(command(2)) > 0 then
	   '' carregar .csv primeiro com dados de NF-e e CT-e 
	   var i = 1
	   var arquivoEntrada = command(1)
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
   
	   i = 1
	   arquivoEntrada = command(1)
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
				if not e.processar( @mostrarProgresso ) then
					print !"\r\nErro ao extrair arquivo: "; arquivoEntrada
					end -1
				end if
				print "OK!"
			end if 
			 
			i += 1
			arquivoEntrada = command(i)
	   loop
	   
	else
		print "Carregando arquivo";
		ultPorCompleto = 0
		var arquivoEntrada = command(1)
		if not e.carregarTxt( arquivoEntrada, @mostrarProgresso ) then
			print !"\r\nErro ao carregar arquivo: "; arquivoEntrada
			end -1
		end if
		print "OK!"
	
		print "Processando";
		ultPorCompleto = 0
		if not e.processar( @mostrarProgresso ) then
			print !"\r\nErro ao extrair arquivo: "; arquivoEntrada
			end -1
		end if
		print "OK!"
	end if
	
	print wstr("Realizando cruzamentos e análises");
	ultPorCompleto = 0
	e.analisar(@mostrarProgresso)
	print "OK!"
   
	print wstr("Gravando arquivo de saída: "); arquivoSaida + ".xml";
	ultPorCompleto = 0
	e.finalizarExtracao( @mostrarProgresso )
	print "OK!"
   
end sub
   

main()
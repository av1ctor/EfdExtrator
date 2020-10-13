#include once "Efd-GUI.bi"

#include "icons.bas"

dim shared curFileGrid as FileGridData ptr
dim shared curFile as TFile ptr
dim shared statusBar as Ihandle ptr
dim shared outPathEdit as Ihandle ptr

''
private function getFiles(filter as string, filterInfo as string, files as TList ptr) as integer

	var dlg = IupFileDlg()

	IupSetAttribute(dlg, "DIALOGTYPE", "OPEN")
	IupSetAttribute(dlg, "TITLE", "Selecione os arquivos")
	IupSetAttribute(dlg, "MULTIPLEFILES", "YES")
	IupSetAttribute(dlg, "FILTER", filter)
	IupSetAttribute(dlg, "FILTERINFO", filterInfo)

	IupPopup(dlg, IUP_CURRENT, IUP_CURRENT)

	if IupGetInt(dlg, "STATUS") <> -1 then
		var v = *IupGetAttribute(dlg, "VALUE")
		if instr(v, "|") > 0 then
			dim parts() as string
			splitstr(v, "|", parts())
			dim path as string = parts(0) + "\"
			for i as integer = 0 to ubound(parts)-2
				var file = cast(TFile ptr, files->add())
				file->path = path
				file->name = parts(1+i)
				file->num = i
			next
			function = ubound(parts)-1
		
		else
			var p = instrrev(v, "\")
			var file = cast(TFile ptr, files->add())
			file->path = left(v, p)
			file->name = mid(v, p+1)
			file->num = 0
			function = 1
		end if
		
	else
		function = 0
	end if

	IupDestroy(dlg)

end function

#define ROWCOL(r, c) (r) & ":" & (c)

private function edition_cb cdecl(self as Ihandle ptr, lin as long, col as long, update as long) as long
	return IUP_IGNORE
end function

private function dropcheck_cb cdecl(self as Ihandle ptr, lin as long, col as long) as long
	return IUP_IGNORE
end function

private function togglevalue_cb(self as Ihandle ptr, lin as long, col as long, value as long) as long
  return IUP_DEFAULT
end function

private function editaction_cb(self as Ihandle ptr, c as long, value as zstring ptr) as long
	return IUP_IGNORE
end function

sub addFileToGrid(file as TFile ptr, at as integer, mat as Ihandle ptr)
	IupSetInt(mat, ROWCOL(1+at, 0), 1+at)
	IupSetAttribute(mat, ROWCOL(1+at, 1), file->name)
	IupSetAttribute(mat, ROWCOL(1+at, 2), "Selecionado")
	IupSetInt(mat, ROWCOL(1+at, 3), 0)
end sub

private function dropfiles_cb(self as Ihandle ptr, fname as zstring ptr, num as long, x as long, y as long) as long
	var dat = cast(FileGridData ptr, IupGetAttribute(self, "FGDATA"))
	
	if num >= dat->num then
		if dat->files <> null then
			delete dat->files
		end if
		dat->files = new TList(10, len(TFile))
		dat->num = num
		IupSetInt(dat->mat, "NUMLIN", num+1)
	
	elseif num = 0 then
		dat->num = 0
	end if

	var p = instrrev(*fname, "\")
	var file = cast(TFile ptr, dat->files->add())
	var path = left(*fname, p)
	file->path = path
	file->name = mid(*fname, p+1)
	file->num = num
	
	addFileToGrid(file, num, dat->mat)
	
	if num = 0 then
		if len(*IupGetAttribute(outPathEdit, "VALUE")) = 0 then
			IupSetStrAttribute(outPathEdit, "VALUE", path)
		end if
	end if
	
	IupSetAttribute(dat->mat, "REDRAW", "L" & (1+num))
	
	return IUP_DEFAULT
end function

private sub showSelectFilesAndUpdateMatrix(dat as FileGridData ptr)
	
	var files = new TList(10, len(TFile))
	
	var num = getFiles(dat->filter, dat->filterInfo, files)
	if num > 0 then
		if dat->files <> null then
			delete dat->files
		end if
		dat->files = files
		
		IupSetInt(dat->mat, "NUMLIN", num)
	
		var file = cast(TFile ptr, files->head)
		if len(*IupGetAttribute(outPathEdit, "VALUE")) = 0 then
			IupSetStrAttribute(outPathEdit, "VALUE", file->path)
		end if
		
		var i = 0
		do while file <> null
			addFileToGrid(file, i, dat->mat)
			file = files->next_(file)
			i += 1
		loop
		
		IupSetAttribute(dat->mat, "REDRAW", "L1-" & i)
	
	else
		delete files
	end if
end sub

private function selfiles_cb(self as Ihandle ptr) as long
	var dat = cast(FileGridData ptr, IupGetAttribute(self, "FGDATA"))
	showSelectFilesAndUpdateMatrix(dat)
	return IUP_DEFAULT
end function

private function selfiles_click_cb(self as Ihandle ptr, button as long, pressed as long, x as long, y as long, status as zstring ptr) as long
	if button = IUP_BUTTON1 and pressed  then
		var dat = cast(FileGridData ptr, IupGetAttribute(self, "FGDATA"))
		showSelectFilesAndUpdateMatrix(dat)
		return IUP_IGNORE
	else
		return IUP_DEFAULT
	end if
end function

function EfdGUI.buildFileGrid(grid as FILE_GRID, title as string, filter as string, filterInfo as string) as IHandle ptr

	''
	var mat = IupMatrix(NULL)

	var dat = @fileGrids(grid)
	dat->filter = filter
	dat->filterInfo = filterInfo
	dat->files = null
	dat->mat = mat
	
	IupSetAttribute(mat, "EXPAND", "YES")

	IupSetInt(mat, "NUMLIN", 0)
	IupSetInt(mat, "NUMCOL", 3)
	IupSetInt(mat, "NUMCOL_VISIBLE", 3)
	IupSetInt(mat, "NUMLIN_VISIBLE", 5)
	
	IupSetAttribute(mat, "SHOWFILLVALUE", "YES")
	'IupSetAttribute(mat, "TOGGLECENTERED", "YES")
	IupSetInt(mat, "WIDTHDEF", 40)
	'IupSetInt(mat, "HEIGHTDEF", 12)
	'IupSetAttribute(mat, "RESIZEMATRIX", "Yes")

	IupSetAttribute(mat, ROWCOL(0, 0), "#")
	IupSetAttribute(mat, ROWCOL(0, 1), "Nome")
	IupSetAttribute(mat, ROWCOL(0, 2), "Estado")
	IupSetAttribute(mat, ROWCOL(0, 3), "Progresso")

	IupSetAttribute(mat, "TYPE*:3", "FILL")
	IupSetAttribute(mat, "FGCOLOR*:3", "#008080")
	
	IupSetInt(mat, "WIDTH1", 400)
	IupSetInt(mat, "WIDTH2", 150)
	IupSetInt(mat, "WIDTH3", 100)
	
	IupSetAttribute(mat, "ALIGNMENT1", "ALEFT")
	IupSetAttribute(mat, "ALIGNMENT2", "ALEFT")
	
	'IupSetCallback(mat, "DROPCHECK_CB", cast(Icallback, @dropcheck_cb))
	'IupSetCallback(mat, "TOGGLEVALUE_CB", cast(Icallback, @togglevalue_cb))
	IupSetCallback(mat, "EDITION_CB", cast(Icallback, @edition_cb))
	
	''
	var edit = IupText(null)
	IupSetAttribute(edit, "VALUE", "Clique para selecionar os arquivos, ou arraste e solte-os aqui...")
	IupSetAttribute(edit, "FGDATA", cast(zstring ptr, dat))
	IupSetAttribute(edit, "EXPAND", "HORIZONTAL")
	IupSetCallback(edit, "ACTION", cast(Icallback, @editaction_cb))
	IupSetCallback(edit, "DROPFILES_CB", cast(Icallback, @dropfiles_cb))
	IupSetCallback(edit, "BUTTON_CB", cast(Icallback, @selfiles_click_cb))
	
	'var but = IupButton("Selecionar...", NULL)
	'IupSetAttribute(but, "FGDATA", cast(zstring ptr, dat))
	'IupSetCallback(but, "ACTION", cast(Icallback, @selfiles_cb))
	
	var hbox = IupHbox _
		( _
			edit, _ 'but, _
			NULL _
		) _

	IupSetAttribute(hbox, "GAP", "10")
	IupSetAttribute(hbox, "ALIGNMENT", "ACENTER")

	''
	var vbox = IupVbox _
		( _
			hbox, _
			IupFill(), _
			mat, _
			NULL _
		) _
		
	''
	var frm = IupFrame(vbox)
	IupSetStrAttribute(frm, "TITLE", title)

	return frm

end function

private function item_efd_action_cb(item as Ihandle ptr) as long
	var dat = cast(FileGridData ptr, IupGetAttribute(IupGetDialog(item), "FG_EFD"))
	showSelectFilesAndUpdateMatrix(dat)
	return IUP_DEFAULT
end function

private function item_dfe_action_cb(item as Ihandle ptr) as long
	var dat = cast(FileGridData ptr, IupGetAttribute(IupGetDialog(item), "FG_DFE"))
	showSelectFilesAndUpdateMatrix(dat)
	return IUP_DEFAULT
end function

private function item_exit_action_cb(item as Ihandle ptr) as long
	return IUP_CLOSE
end function

private function item_about_action_cb(item as Ihandle ptr) as long
	IupMessage("Sobre", !"Extrator de EFD/Sintegra para Excel, versão 0.9.1 beta\nCopyleft 2017-2020 by André Vicentini (avtvicentini)")
	return IUP_DEFAULT
end function

private sub onProgress(estagio as const zstring ptr, completado as double = 0)
	static ultCompletado as double = 0
	
	dim msg as string = ""
	if estagio <> null then
		msg += *estagio
	end if
	
	var useStatusBar = curFileGrid = null orelse curFile = null
	
	if completado = 0 then
		ultCompletado = 0
	else
		if useStatusBar then
			do while completado >= ultCompletado + 0.05
				msg += "."
				ultCompletado += 0.05
			loop
		
			if completado = 1 then
				msg += "OK!"
			end if
		end if
	end if
	
	if not useStatusBar then
		var l = 1+curFile->num
		if len(msg) > 0 then
			IupSetStrAttribute(curFileGrid->mat, ROWCOL(l, 2), msg)
		end if

		IupSetInt(curFileGrid->mat, ROWCOL(l, 3), cint(completado * 100))
		
		if len(msg) > 0 orelse completado - ultCompletado >= 0.05 then
			IupSetAttribute(curFileGrid->mat, "REDRAW", "L" & l)
			IupSetAttribute(curFileGrid->mat, "SHOW", l & ":*")
			ultCompletado = completado
		end if
		
	else
		if len(msg) > 0 then
			IupSetStrAttribute(statusBar, "TITLE", msg)
		end if
	end if
	
	IupFlush()
end sub

private sub onError(msg as const zstring ptr)
	if len(msg) > 0 then
		if curFileGrid <> null andalso curFile <> null then
			var l = (1+curFile->num)
			IupSetStrAttribute(curFileGrid->mat, ROWCOL(l, 2), msg)
			IupSetAttribute(curFileGrid->mat, "REDRAW", "L" & l)
			IupSetAttribute(curFileGrid->mat, "SHOW", l & ":*")
		else
			IupSetStrAttribute(statusBar, "TITLE", msg)
		end if
		IupFlush()
	end if
end sub

private sub toggleActionButton(to_ as string)
	var btn = IupGetHandle("EFD_BTN_EXEC")
	IupSetAttribute(btn, "ACTIVE", to_)
end sub

private function item_exec_action_cb(item as Ihandle ptr) as long

	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "EFD_SELF"))

	toggleActionButton("NO")
	
	onProgress("Iniciando...")
	
	var ext = new Efd(@onProgress, @onError)
	
	var path = *IupGetAttribute(outPathEdit, "VALUE")
	if len(path) > 0 then
		chdir path
	else
		chdir exepath()
	end if
	
	dim arquivoSaida as string =  "__efd__"
	ext->iniciarExtracao(arquivoSaida, gui->opcoes)

	var errCnt = 0
	
	onProgress("Processando...")
	
	curFileGrid = @gui->fileGrids(FG_DFE)
	if curFileGrid->files <> null then
		curFile = cast(TFile ptr, curFileGrid->files->head)
		do while curFile <> null
			onProgress("Carregando")
			
			var arquivoEntrada = curFile->path + curFile->name
			if lcase(right(arquivoEntrada,3)) = "csv" then
				if not ext->carregarCsv( arquivoEntrada ) then
					onError(!"\r\nErro ao carregar arquivo: " & arquivoEntrada)
					errCnt += 1
				end if
				
			elseif lcase(right(arquivoEntrada,4)) = "xlsx" then
				if not ext->carregarXlsx( arquivoEntrada ) then
					onError(!"\r\nErro ao carregar arquivo: " & arquivoEntrada)
					errCnt += 1
				end if
			end if 
			
			IupFlush()

			curFile = curFileGrid->files->next_(curFile)
		loop	
	end if
	
	var efdCnt = 0
	curFileGrid = @gui->fileGrids(FG_EFD)
	if curFileGrid->files <> null then
		curFile = cast(TFile ptr, curFileGrid->files->head)
		do while curFile <> null
			var arquivoEntrada = curFile->path + curFile->name
			if lcase(right(arquivoEntrada,3)) = "txt" then
				onProgress("Carregando")
				if not ext->carregarTxt( arquivoEntrada ) then
					onError(!"\r\nErro ao carregar arquivo: " & arquivoEntrada)
					errCnt += 1
				end if
				
				efdCnt += 1
				
				if errCnt = 0 then
					onProgress("Processando")
					if not ext->processar( arquivoEntrada ) then
						onError(!"\r\nErro ao extrair arquivo: " & arquivoEntrada)
						errCnt += 1
					end if
				end if
			end if 
			
			IupFlush()
			 
			curFile = curFileGrid->files->next_(curFile)
		loop
	end if
	
	curFileGrid = null
	curFile = null
	
	if errCnt = 0 andalso efdCnt > 0 then
		if gui->opcoes.formatoDeSaida <> FT_NULL then
			onProgress("Analisando")
			IupFlush()
			ext->analisar()

			onProgress("Resumindo")
			IupFlush()
			ext->criarResumos()
		end if
	end if
   
	''
	if errCnt = 0 andalso efdCnt > 0 then
		IupFlush()
		ext->finalizarExtracao()
	end if
	
	onProgress("Finalizado!")

	toggleActionButton("YES")
	
	delete ext

	return IUP_DEFAULT
end function

function EfdGUI.buildMenu() as IHandle ptr
	
	'' Arquivo
	var item_efd = IupItem("Selecionar &EFD's...", NULL)
	IupSetAttribute(item_efd, "IMAGE", "EFD_OPEN_EFD_ICON")
	IupSetCallback(item_efd, "ACTION", cast(Icallback, @item_efd_action_cb))

	var item_dfe = IupItem("Selecionar &DFe's...", NULL)
	IupSetAttribute(item_dfe, "IMAGE", "EFD_OPEN_DFE_ICON")
	IupSetCallback(item_dfe, "ACTION", cast(Icallback, @item_dfe_action_cb))

	var item_exec = IupItem("Executar", NULL)
	IupSetAttribute(item_exec, "IMAGE", "EFD_EXEC_ICON")
	IupSetCallback(item_exec, "ACTION", cast(Icallback, @item_exec_action_cb))

	var item_exit = IupItem("&Sair", NULL)
	IupSetAttribute(item_exit, "IMAGE", "EFD_EXIT_ICON")
	IupSetCallback(item_exit, "ACTION", cast(Icallback, @item_exit_action_cb))
	
	var file_menu = IupMenu _
	( _
		item_efd, _
		item_dfe, _
		item_exec, _
		item_exit, _
		NULL _
	)
	
	var sub_menu_file = IupSubmenu("&Arquivo", file_menu)

	'' Ajuda
	var item_about = IupItem("&Sobre...", NULL)
	IupSetAttribute(item_about, "IMAGE", "EFD_HELP_ICON")
	IupSetCallback(item_about, "ACTION", cast(Icallback, @item_about_action_cb))

	var help_menu = IupMenu _
	( _
		item_about, _
		NULL _
	)
	
	var sub_menu_help = IupSubmenu("A&juda", help_menu)
	
	''
	var menu = IupMenu _
	( _
		sub_menu_file, _
		sub_menu_help,_
		NULL _
	)
	
	return menu
end function

function EfdGUI.buildToolBar() as Ihandle ptr

	var btn_open_efd = IupButton(NULL, NULL)
	IupSetAttribute(btn_open_efd, "IMAGE", "EFD_OPEN_EFD_ICON")
	IupSetAttribute(btn_open_efd, "FLAT", "YES")
	IupSetCallback(btn_open_efd, "ACTION", cast(Icallback, @item_efd_action_cb))
	IupSetAttribute(btn_open_efd, "TIP", "Selecionar EFD's")
	IupSetAttribute(btn_open_efd, "CANFOCUS", "No")
	
	var btn_open_dfe = IupButton(NULL, NULL)
	IupSetAttribute(btn_open_dfe, "IMAGE", "EFD_OPEN_DFE_ICON")
	IupSetAttribute(btn_open_dfe, "FLAT", "YES")
	IupSetCallback(btn_open_dfe, "ACTION", cast(Icallback, @item_dfe_action_cb))
	IupSetAttribute(btn_open_dfe, "TIP", "Selecionar DFe's")
	IupSetAttribute(btn_open_dfe, "CANFOCUS", "No")
	
	var toolbar = IupHbox _
	( _
		btn_open_efd, _
		btn_open_dfe, _
		IupSetAttributes(IupLabel(NULL), "SEPARATOR=VERTICAL"), _
		NULL _
	)
	
	IupSetAttribute(toolbar, "MARGIN", "0x0")
	IupSetAttribute(toolbar, "GAP", "2")
	
	return toolbar
end function

function EfdGUI.buildStatusBar() as Ihandle ptr
	var bar = IupLabel("Selecione os arquivos e clique no botão Executar")
	IupSetAttribute(bar, "NAME", "STATUSBAR")
	IupSetAttribute(bar, "EXPAND", "HORIZONTAL")
	IupSetAttribute(bar, "PADDING", "10x5")
	return bar
end function

type TOption
	name as zstring * 32
	label as zstring * 128
end type

private function opcao_action_cb(item as Ihandle ptr, state as long) as long
	
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "EFD_SELF"))
	
	var name_ = *IupGetAttribute(item, "OPTIONNAME")
	select case name_
	case "gerarrelatorios"
		gui->opcoes.gerarRelatorios = state
	case "naogerarlre"
		gui->opcoes.pularLre = state
	case "naogerarlrs"
		gui->opcoes.pularLrs = state
	case "naogerarlraicms"
		gui->opcoes.pularLraicms = state
	case "realcar"
		gui->opcoes.highlight = state
	case "dbemdisco"
		gui->opcoes.dbEmDisco = state
	case "manterdb"
		gui->opcoes.manterDb = state
		gui->opcoes.dbEmDisco = state
	end select
	
	return IUP_DEFAULT
end function

private function cnpjs_set_cb(item as Ihandle ptr) as long
	
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "EFD_SELF"))
	var list = cast(IHandle ptr, IupGetAttribute(item, "CNPJS_LIST"))
	var edit = cast(IHandle ptr, IupGetAttribute(item, "CNPJS_EDIT"))
	
	var value = *IupGetAttribute(edit, "VALUE")
	IupSetStrAttribute(edit, "VALUE", "")
	
	var cnt = iif(len(value) > 0, splitstr(value, ",", gui->opcoes.listaCnpj()), 0)
	IupSetStrAttribute(list, "1", null)
	if cnt > 0 then
		for i as integer = 0 to cnt-1
			IupSetStrAttribute(list, str(1+i), gui->opcoes.listaCnpj(i))
		next
		gui->opcoes.filtrarCnpj = true
	else
		gui->opcoes.filtrarCnpj = false
	end if
	
	return IUP_DEFAULT
end function

function EfdGUI.buildCnpjFilterBox() as IHandle ptr
	var list = IupList(NULL)
	IupSetAttribute(list, "SIZE", "150x60")
	IupSetAttribute(list, "EXPAND", "HORIZONTAL")
	
	var edit = IupText(NULL)
	IupSetAttribute(edit, "EXPAND", "HORIZONTAL")
	
	var btn = IupButton("Filtrar", NULL)
	IupSetAttribute(btn, "CNPJS_EDIT", cast(zstring ptr, edit))
	IupSetAttribute(btn, "CNPJS_LIST", cast(zstring ptr, list))
	IupSetCallback(btn, "ACTION", cast(Icallback, @cnpjs_set_cb))
	
	var hbox = IupHBox _
	( _
		edit, _
		btn, _
		NULL _
	)
	IupSetAttribute(hbox, "MARGIN", "0x0")
	
	var vbox = IupVbox _
	( _
		IupLabel("Filtrar por CNPJ (sep. por vírgula)"), _
		hbox, _
		list, _
		NULL _
	)

	IupSetAttribute(vbox, "MARGIN", "5x0")
	
	return vbox
end function

private function chaves_set_cb(item as Ihandle ptr) as long
	
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "EFD_SELF"))
	var list = cast(IHandle ptr, IupGetAttribute(item, "CHAVES_LIST"))
	var edit = cast(IHandle ptr, IupGetAttribute(item, "CHAVES_EDIT"))
	
	var value = *IupGetAttribute(edit, "VALUE")
	IupSetStrAttribute(edit, "VALUE", "")
	
	var cnt = iif(len(value) > 0, splitstr(value, ",", gui->opcoes.listaChaves()), 0)
	IupSetStrAttribute(list, "1", null)
	if cnt > 0 then
		for i as integer = 0 to cnt-1
			IupSetStrAttribute(list, str(1+i), gui->opcoes.listaChaves(i))
		next
		gui->opcoes.filtrarChaves = true
	else
		gui->opcoes.filtrarChaves = false
	end if
	
	return IUP_DEFAULT
end function

function EfdGUI.buildChavesFilterBox() as IHandle ptr
	var list = IupList(NULL)
	IupSetAttribute(list, "SIZE", "150x60")
	IupSetAttribute(list, "EXPAND", "HORIZONTAL")
	
	var edit = IupText(NULL)
	IupSetAttribute(edit, "EXPAND", "HORIZONTAL")
	
	var btn = IupButton("Filtrar", NULL)
	IupSetAttribute(btn, "CHAVES_EDIT", cast(zstring ptr, edit))
	IupSetAttribute(btn, "CHAVES_LIST", cast(zstring ptr, list))
	IupSetCallback(btn, "ACTION", cast(Icallback, @chaves_set_cb))
	
	var hbox = IupHBox _
	( _
		edit, _
		btn, _
		NULL _
	)
	IupSetAttribute(hbox, "MARGIN", "0x0")
	IupSetAttribute(hbox, "GAP", "5")
	
	var vbox = IupVbox _
	( _
		IupLabel("Filtrar por chave (sep. por vírgula)"), _
		hbox, _
		list, _
		NULL _
	)
	
	IupSetAttribute(vbox, "MARGIN", "5x0")
	
	return vbox
end function

private function format_action_cb(self as Ihandle ptr, text as zstring ptr, item as long, state as long) as long
	
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(self), "EFD_SELF"))
	if state = 1 then
		select case *text
		case "xml" 
			gui->opcoes.formatoDeSaida = FT_XML
		case "csv"
			gui->opcoes.formatoDeSaida = FT_CSV
		case "xlsx"
			gui->opcoes.formatoDeSaida = FT_XLSX
		case "null"
			gui->opcoes.formatoDeSaida = FT_NULL
		end select
	end if

	return IUP_DEFAULT
end function

function EfdGUI.buildOutFormatBox() as Ihandle ptr
	var list = IupList(NULL)
	IupSetAttribute(list, "EXPAND", "HORIZONTAL")
	IupSetAttribute(list, "DROPDOWN", "YES")
	IupSetAttribute(list, "1", "xlsx")
	IupSetAttribute(list, "2", "csv")
	IupSetAttribute(list, "3", "xml")
	IupSetAttribute(list, "4", "null")
	IupSetAttribute(list, "VALUE", "1")
	IupSetCallback(list, "ACTION", cast(Icallback, @format_action_cb))
	
	var box = IupVBox _
	( _
		IupLabel("Formato de saída"), _
		list, _
		NULL _
	)
	
	IupSetAttribute(box, "MARGIN", "0x5")
	
	return box
end function

function EfdGUI.buildOptionsFrame() as Ihandle ptr
	
	dim opcoes(0 to ...) as TOption = { _
		("gerarrelatorios", "Gerar relatórios"), _
		("naogerarlre", "Não gerar LRE"), _
		("naogerarlrs", "Não gerar LRS"), _
		("naogerarlraicms", "Não gerar LRAICMS"), _
		("realcar", "Realçar registros filtrados no relatório"), _
		("dbemdisco", "Criar DB em disco"), _
		("manterdb", "Manter DB em disco") _
	}
	
	var optionsBox = IupVbox _
	( _
		NULL _
	)
	
	IupSetAttribute(optionsBox, "MARGIN", "5x0")
	
	for i as integer = 0 to ubound(opcoes)
		var opcao = @opcoes(i)
		var toggle = IupToggle(opcao->label, NULL)
		IupSetStrAttribute(toggle, "OPTIONNAME", opcao->name)
		IupSetCallback(toggle, "ACTION", cast(Icallback, @opcao_action_cb))
		IupAppend(optionsBox, toggle)
	next
	
	IupAppend(optionsBox, buildOutFormatBox())
	
	var hbox = IupHBox _
	( _
		optionsBox, _
		buildCnpjFilterBox(), _
		buildChavesFilterBox(), _
		NULL _
	)

	IupSetAttribute(hbox, "MARGIN", "0x5")

	outPathEdit = IupText(NULL)
	IupSetAttribute(outPathEdit, "EXPAND", "HORIZONTAL")
	
	var vbox = IupVBox _
	( _
		IupLabel("Pasta de saída (deve existir)"), _
		outPathEdit, _
		hbox, _
		NULL _
	)
	
	IupSetAttribute(vbox, "MARGIN", "5x5")

	var frm = IupFrame(vbox)
	IupSetStrAttribute(frm, "TITLE", "Opções")
	
	return frm
end function

function EfdGUI.buildActionsFrame() as IHandle ptr
	var btn_exec = IupButton("Executar", NULL)
	IupSetHandle("EFD_BTN_EXEC", btn_exec)
	IupSetAttribute(btn_exec, "IMAGE", "EFD_EXEC_ICON")
	IupSetAttribute(btn_exec, "FLAT", "NO")
	IupSetCallback(btn_exec, "ACTION", cast(Icallback, @item_exec_action_cb))
	IupSetAttribute(btn_exec, "TIP", "Executar")
	IupSetAttribute(btn_exec, "CANFOCUS", "NO")

	var hbox = IupHBox _
	( _
		IupFill(), _
		btn_exec, _
		IupFill(), _
		NULL _
	)

	var frm = IupFrame(hbox)
	IupSetStrAttribute(frm, "TITLE", "Ações")
	
	return frm
	
end function

function EfdGUI.buildDlg(efdFrm as IHandle ptr, dfeFrm as IHandle ptr) as IHandle ptr
	
	statusBar = buildStatusBar()
	
	var dlg = IupDialog _
	( _
		IupVbox _
		( _
			buildToolBar(), _
			efdFrm, _
			dfeFrm, _
			buildOptionsFrame(), _
			buildActionsFrame(), _
			statusBar, _
			NULL _
		) _
	)
	
	IupSetAttributeHandle(dlg, "MENU", buildMenu())
	
	IupSetAttribute(dlg, "TITLE", "EfdExtrator")
	IupSetAttribute(dlg, "MARGIN", "10x10")
	IupSetAttribute(dlg, "MINSIZE", "1080x600")
	IupSetAttribute(dlg, "MAXSIZE", "1080x65535")
	IupSetAttribute(efdFrm, "MARGIN", "0x5")
	IupSetAttribute(dfeFrm, "MARGIN", "0x5")
	
	IupSetCallback(dlg, "K_cW", cast(Icallback, @item_exit_action_cb))

	IupShowXY(dlg, IUP_CENTER, IUP_CENTER)

	return dlg
end function

constructor EfdGUI()
	if IupOpen( NULL, NULL ) = IUP_ERROR then
		return
	end if
	
	IupControlsOpen()
	
	IupSetHandle("EFD_OPEN_EFD_ICON", IupImageRGBA(32, 32, @open_efd_icon(0)))
	IupSetHandle("EFD_OPEN_DFE_ICON", IupImageRGBA(32, 32, @open_dfe_icon(0)))
	IupSetHandle("EFD_EXIT_ICON", IupImageRGBA(32, 32, @exit_icon(0)))
	IupSetHandle("EFD_HELP_ICON", IupImageRGBA(32, 32, @help_icon(0)))
	IupSetHandle("EFD_EXEC_ICON", IupImageRGBA(32, 32, @exec_icon(0)))
end constructor

function EfdGUI.build() as boolean
	var dlg = buildDlg( _
		buildFileGrid(FG_EFD, "EFD's", "SPED*.txt", "Arquivos do SPED|SPED*.txt"), _
		buildFileGrid(FG_DFE, "DFe's", "*.xlsx;*.csv", "Arquivos Excel do BO Launch PAD|*.xlsx;Arquivos CSV do SAFI|*.csv"))
		
	IupSetAttribute(dlg, "EFD_SELF", cast(zstring ptr, @this))
	IupSetAttribute(dlg, "FG_EFD", cast(zstring ptr, @fileGrids(FG_EFD)))
	IupSetAttribute(dlg, "FG_DFE", cast(zstring ptr, @fileGrids(FG_DFE)))
		
	return true
end function

sub EfdGUI.run()
	IupMainLoop()
end sub

destructor EfdGUI
	IupClose()
end destructor
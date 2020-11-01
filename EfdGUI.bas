#include once "EfdGUI.bi"

#include "icons.bas"

dim shared running as boolean
dim shared curFileGrid as FileGridData ptr
dim shared curFile as TFile ptr
dim shared statusBar as Ihandle ptr

private function item_about_action_cb(item as Ihandle ptr) as long
	IupMessage("Sobre", !"Extrator de EFD/Sintegra para Excel/Csv/SQLite/Pdf, versão 0.9.3 beta\nCopyleft 2017-2020 by André Vicentini (avtvicentini)")
	return IUP_DEFAULT
end function

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

private sub toggleExecButton(to_ as string)
	var btn = IupGetHandle("EFD_BTN_EXEC")
	IupSetAttribute(btn, "ACTIVE", to_)
end sub

private sub toggleClearButton(to_ as string)
	var btn = IupGetHandle("EFD_BTN_CLEAR")
	IupSetAttribute(btn, "ACTIVE", to_)
end sub

private sub toggleStopButton(to_ as string)
	var btn = IupGetHandle("EFD_BTN_STOP")
	IupSetAttribute(btn, "ACTIVE", to_)
end sub

private function cmp_tfile(num as long, node as any ptr) as boolean
	return num < cast(TFile ptr, node)->num
end function

private function dropfiles_cb(item as Ihandle ptr, fname as zstring ptr, num as long, x as long, y as long) as long
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	var dat = cast(FileGridData ptr, IupGetAttribute(item, "FGDATA"))
	
	if num >= dat->num then
		if dat->files <> null then
			delete dat->files
		end if
		dat->files = new TList(10, len(TFile))
		dat->num = num
		IupSetInt(dat->mat, "NUMLIN", num+1)
	end if

	var at = dat->num - num
	
	var p = instrrev(*fname, "\")
	var file = cast(TFile ptr, dat->files->addOrdAsc(at, @cmp_tfile))
	var path = left(*fname, p)
	file->path = path
	file->name = mid(*fname, p+1)
	file->num = at
	
	addFileToGrid(file, at, dat->mat)
	
	if num = 0 then
		if len(*IupGetAttribute(gui->outPathEdit, "VALUE")) = 0 then
			IupSetStrAttribute(gui->outPathEdit, "VALUE", path)
		end if
		dat->num = 0
		
		toggleExecButton("YES")
	end if
	
	IupSetAttribute(dat->mat, "REDRAW", "L" & (1+at))
	
	return IUP_DEFAULT
end function

private sub showSelectFilesAndUpdateMatrix(gui as EfdGUI ptr, dat as FileGridData ptr)
	
	var files = new TList(10, len(TFile))
	
	var num = getFiles(dat->filter, dat->filterInfo, files)
	if num > 0 then
		if dat->files <> null then
			delete dat->files
		end if
		dat->files = files
		
		IupSetInt(dat->mat, "NUMLIN", num)
	
		var file = cast(TFile ptr, files->head)
		if len(*IupGetAttribute(gui->outPathEdit, "VALUE")) = 0 then
			IupSetStrAttribute(gui->outPathEdit, "VALUE", file->path)
		end if
		
		var i = 0
		do while file <> null
			addFileToGrid(file, i, dat->mat)
			file = files->next_(file)
			i += 1
		loop
		
		IupSetAttribute(dat->mat, "REDRAW", "L1-" & i)
		toggleExecButton("YES")
	
	else
		delete files
	end if
end sub

private function selfiles_cb(item as Ihandle ptr) as long
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	var dat = cast(FileGridData ptr, IupGetAttribute(item, "FGDATA"))
	showSelectFilesAndUpdateMatrix(gui, dat)
	return IUP_DEFAULT
end function

private function selfiles_click_cb(item as Ihandle ptr, button as long, pressed as long, x as long, y as long, status as zstring ptr) as long
	if button = IUP_BUTTON1 and pressed  then
		var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
		var dat = cast(FileGridData ptr, IupGetAttribute(item, "FGDATA"))
		showSelectFilesAndUpdateMatrix(gui, dat)
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
	IupSetAttribute(edit, "CUEBANNER", "Clique para selecionar os arquivos, ou arraste os arquivos e solte-os aqui...")
	IupSetAttribute(edit, "CANFOCUS", "NO")
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
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	var dat = cast(FileGridData ptr, IupGetAttribute(IupGetDialog(item), "FG_EFD"))
	showSelectFilesAndUpdateMatrix(gui, dat)
	return IUP_DEFAULT
end function

private function item_dfe_action_cb(item as Ihandle ptr) as long
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	var dat = cast(FileGridData ptr, IupGetAttribute(IupGetDialog(item), "FG_DFE"))
	showSelectFilesAndUpdateMatrix(gui, dat)
	return IUP_DEFAULT
end function

private function item_exit_action_cb(item as Ihandle ptr) as long
	return IUP_CLOSE
end function

private function onProgress(estagio as const zstring ptr, completado as double = 0) as boolean
	static ultCompletado as double = 0
	
	dim msg as string = ""
	if estagio <> null then
		msg += *estagio
	end if
	
	var useStatusBar = curFileGrid = null orelse curFile = null
	
	if completado = 0 then
		ultCompletado = 0
	end if
	
	if not useStatusBar then
		var l = 1+curFile->num
		if len(msg) > 0 then
			IupSetStrAttribute(curFileGrid->mat, ROWCOL(l, 2), msg)
		end if

		if len(msg) > 0 orelse completado = 0 orelse completado = 1 orelse completado - ultCompletado >= 0.01 then
			IupSetInt(curFileGrid->mat, ROWCOL(l, 3), cint(completado * 100))
			IupSetAttribute(curFileGrid->mat, "REDRAW", "L" & l)
			IupSetAttribute(curFileGrid->mat, "SHOW", l & ":*")
			IupFlush()
			ultCompletado = completado
		end if
		
	else
		if len(msg) > 0 then
			IupSetStrAttribute(cast(IHandle ptr, IupGetAttribute(statusBar, "_LABEL")), "TITLE", msg)
		end if

		if len(msg) > 0 orelse completado = 0 orelse completado = 1 orelse completado - ultCompletado >= 0.01 then
			IupSetDouble(cast(IHandle ptr, IupGetAttribute(statusBar, "_PROGRESS")), "VALUE", completado)
			IupFlush()
			ultCompletado = completado
		end if
	end if
	
	return running
	
end function

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

private function item_exec_action_cb(item as Ihandle ptr) as long

	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))

	running = true
	
	toggleExecButton("NO")
	toggleClearButton("NO")
	toggleStopButton("YES")
	
	onProgress("Iniciando...")
	
	var ext = new Efd(@onProgress, @onError)
	
	var path = *IupGetAttribute(gui->outPathEdit, "VALUE")
	if len(path) > 0 then
		chdir path
	else
		chdir exepath()
	end if
	
	dim arquivoSaida as string =  "__efd__"
	ext->iniciar(arquivoSaida, gui->opcoes)

	var errCnt = 0
	
	onProgress("Processando...")

	for i as integer = 0 to ubound(gui->fileGrids)
		var fg = @gui->fileGrids(i)
		if fg->files <> null then
			var file = cast(TFile ptr, fg->files->head)
			do while file <> null
				var at = file->num
				IupSetAttribute(fg->mat, ROWCOL(1+at, 2), "Selecionado")
				IupSetInt(fg->mat, ROWCOL(1+at, 3), 0)
				file = fg->files->next_(file)
			loop
			IupSetAttribute(fg->mat, "REDRAW", "ALL")
		end if
	next
	
	'' DFe's
	curFileGrid = @gui->fileGrids(FG_DFE)
	if curFileGrid->files <> null then
		curFile = cast(TFile ptr, curFileGrid->files->head)
		do while running andalso curFile <> null
			
			var arquivoEntrada = curFile->path + curFile->name
			if lcase(right(arquivoEntrada,3)) = "csv" then
				if not gui->opcoes.pularAnalises orelse gui->opcoes.acrescentarDados orelse gui->opcoes.manterDb then
					onProgress("Carregando")
					if not ext->carregarCsv( arquivoEntrada ) then
						onError(!"\r\nErro ao carregar arquivo: " & arquivoEntrada)
						errCnt += 1
					end if
				else
					onProgress("Ignorado", 1.0)
				end if
				
			elseif lcase(right(arquivoEntrada,4)) = "xlsx" then
				if not gui->opcoes.pularAnalises orelse gui->opcoes.acrescentarDados orelse gui->opcoes.manterDb then
					onProgress("Carregando")
					if not ext->carregarXlsx( arquivoEntrada ) then
						onError(!"\r\nErro ao carregar arquivo: " & arquivoEntrada)
						errCnt += 1
					end if
				else
					onProgress("Ignorado", 1.0)
				end if
			end if 
			
			IupFlush()

			curFile = curFileGrid->files->next_(curFile)
		loop	
	end if
	
	if not gui->opcoes.acrescentarDados then
		ext->descarregarDFe()
	end if
	
	'' EFD's
	var efdCnt = 0
	curFileGrid = @gui->fileGrids(FG_EFD)
	if curFileGrid->files <> null then
		curFile = cast(TFile ptr, curFileGrid->files->head)
		do while running andalso curFile <> null
			var arquivoEntrada = curFile->path + curFile->name
			if lcase(right(arquivoEntrada,3)) = "txt" then
				onProgress("Carregando")
				var txt = ext->carregarTxt( arquivoEntrada )
				if txt = null  then
					onError(!"\r\nErro ao carregar arquivo: " & arquivoEntrada)
					errCnt += 1
				end if
				
				efdCnt += 1
				
				if errCnt = 0 then
					onProgress("Processando")
					if not ext->processar( txt, arquivoEntrada ) then
						onError(!"\r\nErro ao extrair arquivo: " & arquivoEntrada)
						errCnt += 1
					end if
				end if
				
				if txt <> null then
					delete txt
				end if
			end if 
			
			IupFlush()
			 
			curFile = curFileGrid->files->next_(curFile)
		loop
	end if
	
	curFileGrid = null
	curFile = null
	
	if running then
		if errCnt = 0 andalso efdCnt > 0 then
			if gui->opcoes.formatoDeSaida <> FT_NULL then
				if not gui->opcoes.pularAnalises then
					onProgress("Analisando")
					IupFlush()
					ext->analisar()
				end if

				if not gui->opcoes.pularResumos then
					onProgress("Resumindo")
					IupFlush()
					ext->resumir()
				end if
			end if
		end if
	end if

	IupFlush()
	ext->finalizar()
	
	onProgress("Finalizado!")

	toggleExecButton("YES")
	toggleClearButton("YES")
	toggleStopButton("NO")
	
	delete ext
	running = false

	return IUP_DEFAULT
end function

private function item_clear_action_cb(item as Ihandle ptr) as long

	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	
	for i as integer = 0 to ubound(gui->fileGrids)
		var fg = @gui->fileGrids(i)
		IupSetInt(fg->mat, "NUMLIN", 0)
		if fg->files <> null then
			delete fg->files
			fg->files = null
			fg->num = 0
		end if
	next
	
	IupSetStrAttribute(gui->outPathEdit, "VALUE", "")
	
	IupSetStrAttribute(gui->cnpjsList, "1", null)
	erase gui->opcoes.listaCnpj
	gui->opcoes.filtrarCnpj = false
	
	IupSetStrAttribute(gui->chavesList, "1", null)
	erase gui->opcoes.listaChaves
	gui->opcoes.filtrarChaves = false
	
	toggleExecButton("NO")

	return IUP_DEFAULT
end function

private function item_stop_action_cb(item as Ihandle ptr) as long

	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	
	running = false

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

	var item_exit = IupItem("&Sair", NULL)
	IupSetAttribute(item_exit, "IMAGE", "EFD_EXIT_ICON")
	IupSetCallback(item_exit, "ACTION", cast(Icallback, @item_exit_action_cb))
	
	var file_menu = IupMenu _
	( _
		item_efd, _
		item_dfe, _
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
	IupSetAttribute(btn_open_efd, "TIP", "Selecionar EFD's...")
	IupSetAttribute(btn_open_efd, "CANFOCUS", "No")
	
	var btn_open_dfe = IupButton(NULL, NULL)
	IupSetAttribute(btn_open_dfe, "IMAGE", "EFD_OPEN_DFE_ICON")
	IupSetAttribute(btn_open_dfe, "FLAT", "YES")
	IupSetCallback(btn_open_dfe, "ACTION", cast(Icallback, @item_dfe_action_cb))
	IupSetAttribute(btn_open_dfe, "TIP", "Selecionar DFe's...")
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
	var label = IupLabel("Selecione os arquivos e clique no botão Iniciar")
	IupSetAttribute(label, "EXPAND", "HORIZONTAL")
	
	var prog = IupProgressBar()
	IupSetAttribute(prog, "EXPAND", "HORIZONTAL")
	IupSetAttribute(prog, "DASHED", "YES")

	var bar = IupHBox _
	( _
		label, _
		prog, _
		NULL _
	)

	IupSetAttribute(bar, "PADDING", "0x0")
	IupSetAttribute(bar, "_LABEL", cast(zstring ptr, label))
	IupSetAttribute(bar, "_PROGRESS", cast(zstring ptr, prog))
	
	return bar
end function

type TOption
	name as zstring * 32
	label as zstring * 128
	tip as zstring * 512
end type

private function opcao_action_cb(item as Ihandle ptr, state as long) as long
	
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	
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
	case "naogerarciap"
		gui->opcoes.pularCiap = state
	case "naoanalisar"
		gui->opcoes.pularAnalises = state
	case "naoresumir"
		gui->opcoes.pularResumos = state
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
	
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	var list = gui->cnpjsList
	var edit = cast(IHandle ptr, IupGetAttribute(item, "CNPJS_EDIT"))
	
	var value = *IupGetAttribute(edit, "VALUE")
	IupSetStrAttribute(edit, "VALUE", "")
	
	var cnt = iif(len(value) > 0, splitstr(value, ",", gui->opcoes.listaCnpj()), 0)
	IupSetStrAttribute(list, "1", null)
	if cnt > 0 then
		for i as integer = 0 to cnt-1
			gui->opcoes.listaCnpj(i) = right("00000000000000" & gui->opcoes.listaCnpj(i), 14)
			IupSetStrAttribute(list, str(1+i), gui->opcoes.listaCnpj(i))
		next
		gui->opcoes.filtrarCnpj = true
	else
		gui->opcoes.filtrarCnpj = false
	end if
	
	return IUP_DEFAULT
end function

private function cnpjs_dropfile_cb(item as Ihandle ptr, fname as zstring ptr, num as long, x as long, y as long) as long
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	var list = gui->cnpjsList
	
	IupSetStrAttribute(list, "1", null)
	var cnt = loadstrings(*fname, gui->opcoes.listaCnpj())
	if cnt > 0 then
		for i as integer = 0 to cnt-1
			gui->opcoes.listaCnpj(i) = right("00000000000000" & gui->opcoes.listaCnpj(i), 14)
			IupSetStrAttribute(list, str(1+i), gui->opcoes.listaCnpj(i))
		next
		gui->opcoes.filtrarCnpj = true
	else
		gui->opcoes.filtrarCnpj = false
	end if

	return IUP_IGNORE
end function

function EfdGUI.buildCnpjFilterBox() as IHandle ptr
	cnpjsList = IupList(NULL)
	IupSetAttribute(cnpjsList, "SIZE", "130x85")
	IupSetAttribute(cnpjsList, "EXPAND", "YES")
	IupSetAttribute(cnpjsList, "CANFOCUS", "NO")
	
	var edit = IupText(NULL)
	IupSetAttribute(edit, "CUEBANNER", "Digite a lista, ou arraste e solte o arquivo aqui...")
	IupSetAttribute(edit, "TIP", "A lista deve ser separada por vírgula. O arquivo deve conter apenas um CNPJ por linha, com zeros à esquerda, sem espaços ou linhas em branco.")
	IupSetAttribute(edit, "EXPAND", "HORIZONTAL")
	IupSetCallback(edit, "DROPFILES_CB", cast(Icallback, @cnpjs_dropfile_cb))
	
	var btn = IupButton("Filtrar", NULL)
	IupSetAttribute(btn, "CNPJS_EDIT", cast(zstring ptr, edit))
	IupSetCallback(btn, "ACTION", cast(Icallback, @cnpjs_set_cb))
	
	var hbox = IupHBox _
	( _
		edit, _
		btn, _
		NULL _
	)
	IupSetAttribute(hbox, "MARGIN", "0x0")
	'IupSetAttribute(hbox, "GAP", "5")
	
	var vbox = IupVbox _
	( _
		IupLabel("Filtrar por CNPJ (sep. por vírgula)"), _
		hbox, _
		cnpjsList, _
		NULL _
	)

	IupSetAttribute(vbox, "MARGIN", "5x0")
	
	return vbox
end function

private function chaves_set_cb(item as Ihandle ptr) as long
	
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	var list = gui->chavesList
	var edit = cast(IHandle ptr, IupGetAttribute(item, "CHAVES_EDIT"))
	
	var value = *IupGetAttribute(edit, "VALUE")
	IupSetStrAttribute(edit, "VALUE", "")
	
	IupSetStrAttribute(list, "1", null)
	var cnt = iif(len(value) > 0, splitstr(value, ",", gui->opcoes.listaChaves()), 0)
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

private function chaves_dropfile_cb(item as Ihandle ptr, fname as zstring ptr, num as long, x as long, y as long) as long
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(item), "_SELF"))
	var list = gui->chavesList
	
	IupSetStrAttribute(list, "1", null)
	var cnt = loadstrings(*fname, gui->opcoes.listaChaves())
	if cnt > 0 then
		for i as integer = 0 to cnt-1
			IupSetStrAttribute(list, str(1+i), gui->opcoes.listaChaves(i))
		next
		gui->opcoes.filtrarChaves = true
	else
		gui->opcoes.filtrarChaves = false
	end if

	return IUP_IGNORE
end function

function EfdGUI.buildChavesFilterBox() as IHandle ptr
	chavesList = IupList(NULL)
	IupSetAttribute(chavesList, "SIZE", "130x85")
	IupSetAttribute(chavesList, "EXPAND", "YES")
	IupSetAttribute(chavesList, "CANFOCUS", "NO")
	
	var edit = IupText(NULL)
	IupSetAttribute(edit, "CUEBANNER", "Digite a lista, ou arraste e solte o arquivo aqui...")
	IupSetAttribute(edit, "TIP", "A lista de chaves deve ser separada por vírgula. O arquivo deve conter apenas uma chave por linha, sem espaços ou linhas em branco.")
	IupSetAttribute(edit, "EXPAND", "HORIZONTAL")
	IupSetCallback(edit, "DROPFILES_CB", cast(Icallback, @chaves_dropfile_cb))
	
	var btn = IupButton("Filtrar", NULL)
	IupSetAttribute(btn, "CHAVES_EDIT", cast(zstring ptr, edit))
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
		chavesList, _
		NULL _
	)
	
	IupSetAttribute(vbox, "MARGIN", "5x0")
	
	return vbox
end function

private function format_action_cb(self as Ihandle ptr, text as zstring ptr, item as long, state as long) as long
	
	var gui = cast(EfdGUI ptr, IupGetAttribute(IupGetDialog(self), "_SELF"))
	if state = 1 then
		select case *text
		case "xml" 
			gui->opcoes.formatoDeSaida = FT_XML
		case "csv"
			gui->opcoes.formatoDeSaida = FT_CSV
		case "xlsx"
			gui->opcoes.formatoDeSaida = FT_XLSX
		case "sqlite"
			gui->opcoes.formatoDeSaida = FT_SQLITE
		case "access"
			gui->opcoes.formatoDeSaida = FT_ACCESS
		case "null"
			gui->opcoes.formatoDeSaida = FT_NULL
		end select
		
		if gui->opcoes.formatoDeSaida = FT_NULL then
			gui->opcoes.pularAnalises = true
			gui->opcoes.pularResumos = true
		else
			gui->opcoes.pularAnalises = false
			gui->opcoes.pularResumos = false
		end if
	end if

	return IUP_DEFAULT
end function

function EfdGUI.buildOutFormatBox() as Ihandle ptr
	dim formatos(0 to ...) as string = { _
		"xlsx", _
		"csv", _
		"xml", _
		"sqlite", _
		"access", _
		"null" _
	}
	
	var list = IupList(NULL)
	IupSetAttribute(list, "EXPAND", "HORIZONTAL")
	IupSetAttribute(list, "TIP", "Formato da planilha que será gerada contendo os registros extraídos. Selecione 'null' para não gerar a planilha.")
	IupSetAttribute(list, "DROPDOWN", "YES")
	for i as integer = 0 to ubound(formatos)
		IupSetStrAttribute(list, str(1+i), formatos(i))
	next
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
		("", "Relatórios", ""), _
		("gerarrelatorios", "Gerar EFD-ICMS-IPI", "Gerar relatórios do PVA EFD-ICMS-IPI em formato PDF."), _
		("realcar", "Realçar registros", "Realçar, nos relatórios do EFD-ICMS-IPI, os registros filtrados por CNPJ e/ou chave."), _
		("", "Não extrair", ""), _
		("naogerarlre", "LRE", "Não extrair os Livros Registro de Entradas."), _
		("naogerarlrs", "LRS", "Não extrair os Livros Registro de Saídas."), _
		("naogerarlraicms", "LRAICMS", "Não extrair os Livros Registro de Apuração."), _
		("naogerarciap", "CIAP", "Não extrair os registros do CIAP."), _
		("", "Não realizar", ""), _
		("naoanalisar", "Análises", "Não gerar a planilha de inconsistências."), _
		("naoresumir", "Resumos", "Não gerar a planilha de resumos."), _
		("", "DB", ""), _
		("dbemdisco", "Em disco", "Gravar o banco de dados intermediário em disco, poupando memória."), _
		("manterdb", "Manter", "Manter o banco de dados intermediário em disco.") _
	}
	
	var optionsBox = IupVbox _
	( _
		NULL _
	)
	
	IupSetAttribute(optionsBox, "MARGIN", "5x0")
	
	var box = cast(IHandle ptr, null)
	
	for i as integer = 0 to ubound(opcoes)
		var opcao = @opcoes(i)
		if len(opcao->name) > 0 then
			var toggle = IupToggle(opcao->label, NULL)
			IupSetStrAttribute(toggle, "OPTIONNAME", opcao->name)
			IupSetStrAttribute(toggle, "TIP", opcao->tip)
			IupSetCallback(toggle, "ACTION", cast(Icallback, @opcao_action_cb))
			IupSetAttribute(toggle, "EXPAND", "HORIZONTAL")
			IupAppend(box, toggle)
		else
			box = IupVBox(NULL)
			var frm = IupFrame(box)
			IupSetStrAttribute(frm, "TITLE", opcao->label)
			IupAppend(optionsBox, frm)
		end if
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
	IupSetAttribute(outPathEdit, "TIP", "Caminho da pasta de destino, onde serão gravados todos os arquivos. A pasta deve existir.")
	
	var vbox = IupVBox _
	( _
		IupLabel("Pasta de destino (deve existir)"), _
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
	var btn_exec = IupButton("Iniciar", NULL)
	IupSetHandle("EFD_BTN_EXEC", btn_exec)
	IupSetAttribute(btn_exec, "IMAGE", "EFD_EXEC_ICON")
	IupSetAttribute(btn_exec, "FLAT", "NO")
	IupSetAttribute(btn_exec, "PADDING", "10x2")
	IupSetAttribute(btn_exec, "TIP", "Iniciar")
	IupSetAttribute(btn_exec, "ACTIVE", "NO")
	IupSetCallback(btn_exec, "ACTION", cast(Icallback, @item_exec_action_cb))

	var btn_clear = IupButton("Limpar", NULL)
	IupSetHandle("EFD_BTN_CLEAR", btn_clear)
	IupSetAttribute(btn_clear, "IMAGE", "EFD_CLEAR_ICON")
	IupSetAttribute(btn_clear, "FLAT", "NO")
	IupSetAttribute(btn_clear, "PADDING", "10x2")
	IupSetAttribute(btn_clear, "TIP", "Limpar formulário")
	IupSetCallback(btn_clear, "ACTION", cast(Icallback, @item_clear_action_cb))

	var btn_stop = IupButton("Interromper", NULL)
	IupSetHandle("EFD_BTN_STOP", btn_stop)
	IupSetAttribute(btn_stop, "IMAGE", "EFD_EXIT_ICON")
	IupSetAttribute(btn_stop, "FLAT", "NO")
	IupSetAttribute(btn_stop, "PADDING", "10x2")
	IupSetAttribute(btn_stop, "TIP", "Interromper extração")
	IupSetAttribute(btn_stop, "ACTIVE", "NO")
	IupSetCallback(btn_stop, "ACTION", cast(Icallback, @item_stop_action_cb))

	var hbox = IupHBox _
	( _
		IupFill(), _
		btn_exec, _
		IupFill(), _
		btn_stop, _
		IupFill(), _
		btn_clear, _
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
	IupSetAttribute(dlg, "MARGIN", "2x2")
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
	IupSetHandle("EFD_CLEAR_ICON", IupImageRGBA(32, 32, @clear_icon(0)))
end constructor

function EfdGUI.build() as boolean
	var dlg = buildDlg( _
		buildFileGrid(FG_EFD, "EFD's", "SPED*.txt", "Arquivos do SPED|SPED*.txt"), _
		buildFileGrid(FG_DFE, "DFe's", "*.xlsx;*.csv", "Arquivos Excel do BO Launch PAD|*.xlsx;Arquivos CSV do SAFI|*.csv"))
		
	IupSetAttribute(dlg, "_SELF", cast(zstring ptr, @this))
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
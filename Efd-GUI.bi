#include once "IUP/iup.bi"
#include once "IUP/IupMatrix.bi"
#include once "list.bi"
#include once "strings.bi"
#include once "Efd.bi"

enum FILE_GRID
	FG_EFD
	FG_DFE
	__FG_LEN__
end enum

type TFile
	path 	as string
	name 	as string
	num 	as integer
end type

type FileGridData
	filter as string
	filterInfo as string
	files as TList ptr
	mat as IHandle ptr
	num as integer
end type

type EfdGUI
public:
	declare constructor()
	declare destructor()
	declare function build() as boolean
	declare sub run()
	
	fileGrids(0 to __FG_LEN__-1) as FileGridData
	opcoes as OpcoesExtracao
	cnpjsList as Ihandle ptr
	chavesList as Ihandle ptr
	outPathEdit as IHandle ptr

private:
	declare function buildFileGrid(grid as FILE_GRID, title as string, filter as string, filterInfo as string) as IHandle ptr
	declare function buildDlg(lfrm as IHandle ptr, rfrm as IHandle ptr) as IHandle ptr
	declare function buildMenu() as IHandle ptr
	declare function buildToolbar() as Ihandle ptr
	declare function buildStatusBar() as Ihandle ptr
	declare function buildOptionsFrame() as Ihandle ptr
	declare function buildCnpjFilterBox() as IHandle ptr
	declare function buildChavesFilterBox() as IHandle ptr
	declare function buildActionsFrame() as IHandle ptr
	declare function buildOutFormatBox() as Ihandle ptr
end type

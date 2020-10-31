#include once "Efd.bi"
#include once "EfdBaseImport.bi"

type EfdSintegraImport extends EfdBaseImport
public:
	declare constructor(opcoes as OpcoesExtracao ptr)
	declare destructor()
	declare function carregar(nomeArquivo as string) as boolean

private:
	sintegraDict 		as TDict ptr
	
	declare function lerRegistroSintegra(bf as bfile) as TRegistro ptr
end type
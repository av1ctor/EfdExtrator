#include once "EfdBaseImport.bi"

type EfdSintegraImport extends EfdBaseImport
public:
	declare constructor(opcoes as OpcoesExtracao ptr)
	declare destructor()
	declare function carregar(nomeArquivo as string) as boolean

private:
	sintegraDict 		as TDict ptr
	
	declare function lerRegistroSintegra(bf as bfile, reg as TRegistro ptr) as Boolean
end type
#include once "efd.bi"
#include once "bfile.bi"
#include once "trycatch.bi"

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
	
	'' devolução? inverter emit <-> dest
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
	
	bf.charCsv				'' pular versão
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
	
	'' NOTA: só será possível saber se é operação de entrada ou saída quando pegarmos 
	''       o CNPJ base do contribuinte, que só vem no final do arquivo.......
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


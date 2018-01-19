
function getCustomCallbacks()
	return {
		-- registro D500 (NOTA FISCAL DE SERVIÇO DE COMUNICAÇÃO (CÓDIGO 21) E NOTA FISCAL DE SERVIÇO DE TELECOMUNICAÇÃO)
		D500 = {
			reader = "NFSCT_ler", 
			writer = "NFSCT_gravar",
			rel = "NFSCT_rel"
		},
		D590 = {
			reader = "NFSCT_RegAnalitico_ler"
		}
	}
end

-- readers (EFD)

	ultimoReg = nil

function NFSCT_ler(f)
	bf_char1(f) -- pular |
	
	reg = { }
	reg.tipo = "D500"
	reg.operacao = bf_int1(f)
	bf_char1(f) -- pular |
	reg.emitente = bf_int1(f)
	bf_char1(f) -- pular |
	reg.idParticipante = bf_varint(f)
	reg.modelo = bf_int2(f)
	bf_char1(f) -- pular |
	reg.situacao = bf_int2(f)
	bf_char1(f) -- pular |
	reg.serie = bf_varchar(f)
	reg.subserie = bf_varchar(f)
	reg.numero = bf_varint(f)
	reg.dataEmi = bf_varchar(f)
	reg.dataEntSaida = bf_varchar(f)
	reg.vTotal = bf_vardbl(f)
	reg.vDesconto = bf_vardbl(f)
	reg.vServico = bf_vardbl(f)
	reg.vServicoNT = bf_vardbl(f)
	reg.vTerc = bf_vardbl(f)
	reg.vDesp = bf_vardbl(f)
	reg.bcICMS = bf_vardbl(f)
	reg.icms = bf_vardbl(f)
	bf_varchar(f)	-- pular cod_inf
	reg.pis = bf_vardbl(f)
	reg.cofins = bf_vardbl(f)
	bf_varchar(f)	-- pular cod_cta
	bf_varint(f)	-- pular tp_assinante
	
	bf_char1(f) -- \r
	bf_char1(f) -- \n
	
	ultimoReg = reg
	
	-- retornar o número de campos e o registro
	return 21, reg
end

function NFSCT_RegAnalitico_ler(f)
	bf_char1(f) -- pular |
	
	reg = { }
	reg.tipo = "D590"
	reg.cst = bf_varint(f)
	reg.cfop = bf_varint(f)
	reg.aliq = bf_vardbl(f)
	reg.valorOp = bf_vardbl(f)
	reg.bcICMS = bf_vardbl(f)
	reg.icms = bf_vardbl(f)
	bf_vardbl(f) -- pular VL_BC_ICMS_UF
	bf_vardbl(f) -- pular VL_ICMS_UF
	reg.redBC = bf_vardbl(f)
	bf_varchar(f) -- pular pular COD_OBS
	
	bf_char1(f) -- \r
	bf_char1(f) -- \n

	ultimoReg.analitico = reg

	-- retornar o número de campos e o registro
	return 8, reg
end

-- writers (Excel)

function criarPlanilhas()
end

function NFSCT_gravar(reg)
	
	if reg.operacao == 0 then
		row = ws_addRow(efd_plan_entradas)
	else
		row = ws_addRow(efd_plan_saidas)
	end
	
	er_addCell(row, reg.operacao)
end 

-- writers (Relatórios)

function NFSCT_rel(dfwd, relatorio, reg)
	
	if relatorio ~= reg.operacao then
		return
	end
	
	part = efd_part_get(reg.idParticipante)
	
	dfwd_setClipboardValueByStr(dfwd, "linha", "demi", YyyyMmDd2DatetimeBR(reg.dataEmi))
	dfwd_setClipboardValueByStr(dfwd, "linha", "dent", YyyyMmDd2DatetimeBR(reg.dataEntSaida))
	dfwd_setClipboardValueByStr(dfwd, "linha", "nro", reg.numero)
	dfwd_setClipboardValueByStr(dfwd, "linha", "mod", reg.modelo)
	dfwd_setClipboardValueByStr(dfwd, "linha", "ser", reg.serie)
	dfwd_setClipboardValueByStr(dfwd, "linha", "sit", format(cdbl(reg.situacao), "00"))
	dfwd_setClipboardValueByStr(dfwd, "linha", "cnpj", STR2CNPJ(part.cnpj))
	dfwd_setClipboardValueByStr(dfwd, "linha", "ie", STR2IE(part.ie))
	dfwd_setClipboardValueByStr(dfwd, "linha", "uf", MUNICIPIO2SIGLA(part.municip))
	dfwd_setClipboardValueByStr(dfwd, "linha", "municip", codMunicipio2Nome(part.municip))
	dfwd_setClipboardValueByStr(dfwd, "linha", "razao", left(part.nome, 64))
	
	dfwd_paste(dfwd, "linha")
	
	efd_rel_addItemAnalitico(reg.situacao, reg.analitico)
end 

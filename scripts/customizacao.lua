
function getCustomCallbacks()
	return {
		-- registro D500 (NOTA FISCAL DE SERVIÇO DE COMUNICAÇÃO (CÓDIGO 21) E NOTA FISCAL DE SERVIÇO DE TELECOMUNICAÇÃO)
		D500 = {
			reader = "NFSCT_ler", 
			writer = "NFSCT_gravar",
			rel_entradas = "NFSCT_rel_entradas"
		},
		-- registro analítico do D500
		D590 = {
			reader = "NFSCT_RegAnalitico_ler"
		}
	}
end

-- readers (EFD)

	ultimoReg = nil
	
function ddMmYyyy2YyyyMmDd(s)
	return string.sub(s, 5, 8) .. string.sub(s, 3, 4) .. string.sub(s, 1, 2)
end

function NFSCT_ler(f, reg)
	bf_char1(f) -- pular |
	
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
	reg.dataEmi = ddMmYyyy2YyyyMmDd(bf_varchar(f))
	reg.dataEntSaida = ddMmYyyy2YyyyMmDd(bf_varchar(f))
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
	
	reg.analitico = nil
	
	bf_char1(f) -- \r
	bf_char1(f) -- \n
	
	ultimoReg = reg
end

function NFSCT_RegAnalitico_ler(f, reg)
	bf_char1(f) -- pular |
	
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
end

-- writers (Excel)

function yyyyMmDd2Datetime(s)
	return string.sub(s, 1, 4) .. "-" .. string.sub(s, 5, 6) .. "-" .. string.sub(s, 7, 8) .. "T00:00:00.000"
end 

function criarPlanilhas()
end

function NFSCT_gravar(reg)
	
	if reg.operacao == 0 then
		row = ws_addRow(efd_plan_entradas)
	else
		row = ws_addRow(efd_plan_saidas)
	end
	
	part = efd_participante_get(reg.idParticipante, false)
	
	er_addCell(row, part.cnpj)
	er_addCell(row, part.ie)
	er_addCell(row, part.uf)
	er_addCell(row, part.nome)
	er_addCell(row, reg.modelo)
	er_addCell(row, reg.subserie)
	er_addCell(row, reg.numero)
	er_addCell(row, yyyyMmDd2Datetime(reg.dataEmi))
	er_addCell(row, yyyyMmDd2Datetime(reg.dataEntSaida))
	er_addCell(row, "")
	er_addCell(row, "REG")
	er_addCell(row, reg.bcICMS)
	er_addCell(row, "")
	er_addCell(row, reg.icms)
	er_addCell(row, "")
	er_addCell(row, "")
	er_addCell(row, "")
	er_addCell(row, "")
	er_addCell(row, reg.vTotal)
	
end 

-- writers (Relatórios)

function YyyyMmDd2DatetimeBR(s)
	return string.sub(s, 7, 8) .. "/" .. string.sub(s, 5, 6) .. "/" .. string.sub(s, 1, 4)
end

function NFSCT_rel_entradas(dfw, reg)
	
	part = efd_participante_get(reg.idParticipante, true)
	
	dfw_setClipboardValueByStr(dfw, "linha", "demi", YyyyMmDd2DatetimeBR(reg.dataEmi))
	dfw_setClipboardValueByStr(dfw, "linha", "dent", YyyyMmDd2DatetimeBR(reg.dataEntSaida))
	dfw_setClipboardValueByStr(dfw, "linha", "nro", reg.numero)
	dfw_setClipboardValueByStr(dfw, "linha", "mod", reg.modelo)
	dfw_setClipboardValueByStr(dfw, "linha", "ser", reg.serie)
	dfw_setClipboardValueByStr(dfw, "linha", "subser", reg.subserie)
	dfw_setClipboardValueByStr(dfw, "linha", "sit", string.sub("00" .. reg.situacao, -2))
	dfw_setClipboardValueByStr(dfw, "linha", "cnpj", part.cnpj)
	dfw_setClipboardValueByStr(dfw, "linha", "ie", part.ie)
	dfw_setClipboardValueByStr(dfw, "linha", "uf", part.uf)
	dfw_setClipboardValueByStr(dfw, "linha", "municip", part.municip)
	dfw_setClipboardValueByStr(dfw, "linha", "razao", part.nome)
	
	dfw_paste(dfw, "linha")
	
	efd_rel_addItemAnalitico(reg.situacao, reg.analitico)
end 

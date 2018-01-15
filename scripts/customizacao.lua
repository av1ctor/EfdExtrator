
function getCustomCallbacks()
	return {
		D500 = {
			reader = "lerNFSCT", 
			writer = "gravarNFSCT"
		}
	}
end

-- readers

function lerNFSCT(f)
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
	
	return 6, reg
end


-- writers

function criarPlanilhas()
end

function gravarNFSCT(reg)
	if reg.operacao == 1 then
		row = ws_addRow(efd_plan_entradas)
	else
		row = ws_addRow(efd_plan_saidas)
	end
	
	er_addCell(row, reg.situacao)
end 

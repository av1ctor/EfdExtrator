#include once "winmin.bi"
#include once "DocxFactoryDyn.bi"

constructor DocxFactoryDyn
	SetDllDirectory(ExePath + "\DocxFactory")
	libh = LoadLibrary("DocxFactory.dll")
end constructor
	

sub DocxFactoryDyn.load(byval p_fileName as const zstring ptr)
	if load_p = null then
		load_p = GetProcAddress(libh, "dfw_load")
	end if
	
	load_p(p_fileName)
end sub

sub DocxFactoryDyn.save(byval p_fileName as const zstring ptr)
	if save_p = null then
		save_p = GetProcAddress(libh, "dfw_save")
	end if
	
	save_p(p_fileName)
end sub

sub DocxFactoryDyn.close()
	if close_p = null then
		close_p = GetProcAddress(libh, "dfw_close")
	end if
	
	close_p()
end sub

sub DocxFactoryDyn.paste(byval p_itemName as const zstring ptr)
	if paste_p = null then
		paste_p = GetProcAddress(libh, "dfw_paste")
	end if
	
	paste_p(p_itemName)
end sub

sub DocxFactoryDyn.merge(byval p_data as const zstring ptr)
	if merge_p = null then
		merge_p = GetProcAddress(libh, "dfw_merge")
	end if
	
	merge_p(p_data)
end sub

sub DocxFactoryDyn.mergeW(byval p_data as const wstring ptr)
	if mergeW_p = null then
		mergeW_p = GetProcAddress(libh, "dfw_mergeW")
	end if
	
	mergeW_p(p_data)
end sub

sub DocxFactoryDyn.setClipboardValueByStr(byval p_itemName as const zstring ptr, byval p_fieldName as const zstring ptr, byval p_value as const zstring ptr)
	if setClipboardValueByStr_p = null then
		setClipboardValueByStr_p = GetProcAddress(libh, "dfw_setClipboardValueByStr")
	end if
	
	setClipboardValueByStr_p(p_itemName, p_fieldName, p_value)
end sub

sub DocxFactoryDyn.setClipboardValueByStrW(byval p_itemName as const wstring ptr, byval p_fieldName as const wstring ptr, byval p_value as const wstring ptr)
	if setClipboardValueByStrW_p = null then
		setClipboardValueByStrW_p = GetProcAddress(libh, "dfw_setClipboardValueByStrW")
	end if
	
	setClipboardValueByStrW_p(p_itemName, p_fieldName, p_value)
end sub
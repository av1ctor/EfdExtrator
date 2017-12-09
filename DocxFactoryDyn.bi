

type DocxFactoryDyn
	declare	constructor
	declare sub load(byval p_fileName as const zstring ptr)
	declare sub save(byval p_fileName as const zstring ptr)
	declare sub close()
	declare sub paste(byval p_itemName as const zstring ptr)
	declare sub merge(byval p_data as const zstring ptr)
	declare sub mergeW(byval p_data as const wstring ptr)
	declare sub setClipboardValueByStr(byval p_itemName as const zstring ptr, byval p_fieldName as const zstring ptr, byval p_value as const zstring ptr)
	declare sub setClipboardValueByStrW(byval p_itemName as const wstring ptr, byval p_fieldName as const wstring ptr, byval p_value as const wstring ptr)

private:
	libh		as long
	load_p		as sub (byval p_fileName as const zstring ptr)
	save_p		as sub (byval p_fileName as const zstring ptr)
	close_p		as sub ()
	paste_p		as sub (byval p_itemName as const zstring ptr)
	merge_p		as sub (byval p_data as const zstring ptr)
	mergeW_p	as sub (byval p_data as const wstring ptr)
	setClipboardValueByStr_p as sub (byval p_itemName as const zstring ptr, byval p_fieldName as const zstring ptr, byval p_value as const zstring ptr)
	setClipboardValueByStrW_p as sub (byval p_itemName as const wstring ptr, byval p_fieldName as const wstring ptr, byval p_value as const wstring ptr)
end type


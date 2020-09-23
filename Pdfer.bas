'' PDFium Helper Library for FreeBASIC
'' Copyright 2020 by Andre Victor (av1ctortv[@]gmail.com)

#include once "Pdfer.bi"
#include once "libiconv.bi"

#ifdef WITH_PARSER
	dim shared cd8to16le as iconv_t
	dim shared cd16leto8 as iconv_t
#endif	

private sub initialize() constructor
	FPDF_InitLibrary()
#ifdef WITH_PARSER
	cd8to16le = iconv_open("UTF-16LE", "UTF-8")
	cd16leto8 = iconv_open("UTF-8", "UTF-16LE")
#endif
end sub

private sub shutdown() destructor
#ifdef WITH_PARSER
	iconv_close(cd16leto8)
	iconv_close(cd8to16le)
#endif
	FPDF_DestroyLibrary()
end sub

'''''
destructor PdfObj
	if obj <> null then
		FPDFPageObj_Destroy(obj)
	end if
end destructor

sub PdfObj.setBorder(r as ulong, g as ulong, b as ulong, a as ulong, width_ as single)
	FPDFPageObj_SetStrokeColor(obj, r, g, b, a)
	if width_ >= 0 then
		FPDFPageObj_SetStrokeWidth(obj, width_)
	end if
end sub

sub PdfObj.setBackground(r as ulong, g as ulong, b as ulong, a as ulong)
	FPDFPageObj_SetFillColor(obj, r, g, b, a)
end sub

'''''
constructor PdfRect(x as single, y as single, w as single, h as single)
	obj = FPDFPageObj_CreateNewRect(x, y, w, h)
end constructor

'''''
constructor PdfPath(x as single, y as single)
	obj = FPDFPageObj_CreateNewPath(x, y)
end constructor

sub PdfPath.moveTo(x as single, y as single)
	FPDFPath_MoveTo(obj, x, y)
end sub

sub PdfPath.lineTo(x as single, y as single)
	FPDFPath_LineTo(obj, x, y)
end sub

sub PdfPath.bezierTo(x1 as single, y1 as single, x2 as single, y2 as single, x3 as single, y3 as single)
	FPDFPath_BezierTo(obj, x1, y1, x2, y2, x3, y3)
end sub

sub PdfPath.close()
	FPDFPath_Close(obj)
end sub

'''''
constructor PdfFinder(handle as FPDF_SCHHANDLE)
	this.handle = handle
end constructor

destructor PdfFinder()
	if handle <> null then
		FPDFText_FindClose(handle)
	end if
end destructor

function PdfFinder.find(dir as PdfFinderDirection) as PdfFinderResult
	dim res as PdfFinderResult
	
	var found = iif(dir = PdfFinderDirection.DOWN, _
		FPDFText_FindNext(handle), _
		FPDFText_FindPrev(handle))
		
	if found then
		res.index = FPDFText_GetSchResultIndex(handle)
		res.count = FPDFText_GetSchCount(handle)
	end if
	
	return res
end function

'''''
constructor PdfText(text as FPDF_TEXTPAGE)
	this.text = text
	len_ = FPDFText_CountChars(text)
end constructor

destructor PdfText()
	if text <> null then
		FPDFText_ClosePage(text)
	end if
end destructor

property PdfText.length() as integer
	return len_
end property

property PdfText.value() as wstring ptr
	var buf = allocate((len_+1) * 2)
	var chars = cint(FPDFText_GetText(text, 0, len_, buf))
	return buf
end property

function PdfText.find(what as wstring ptr, index as long, flags as PdfFindFlags) as PdfFinder ptr
	var handle = FPDFText_FindStart(text, what, flags, index)
	return new PdfFinder(handle)
end function

function PdfText.getRect(index as long, count as long) as PdfRectCoords ptr
	var rects = FPDFText_CountRects(text, index, count)
	var rect = new PdfRectCoords
	if rects > 0 then
		FPDFText_GetRect(text, 0, @rect->left, @rect->top, @rect->right, @rect->bottom)
	end if
	return rect
end function

'''''
constructor PdfPage(page as FPDF_PAGE)
	this.page = page
end constructor

destructor PdfPage()
	if page <> null then
		FPDF_ClosePage(page)
	end if
end destructor

property PdfPage.text() as PdfText ptr
	var textPage = FPDFText_LoadPage(page)
	return new PdfText(textPage)
end property

sub PdfPage.highlight(rect as PdfRectCoords ptr)
	var annot = FPDFPage_CreateAnnot(page, FPDF_ANNOT_HIGHLIGHT)
	FPDFAnnot_SetFlags(annot, FPDF_ANNOT_FLAG_PRINT or FPDF_ANNOT_FLAG_READONLY)
	dim rectf as FS_RECTF = (rect->left-1, rect->top-1, rect->right+1, rect->bottom+1)
	FPDFAnnot_SetRect(annot, @rectf)
	dim quad as FS_QUADPOINTSF = (rect->left-1, rect->top+1, rect->right+1, rect->top+1, rect->left-1, rect->bottom-1, rect->right+1, rect->bottom-1)
	FPDFAnnot_AppendAttachmentPoints(annot, @quad)
	FPDFAnnot_SetColor(annot, FPDFANNOT_COLORTYPE_Color, 255, 209, 0, 102)
	FPDFPage_CloseAnnot(annot)
end sub

sub PdfPage.highlight(txt as PdfText ptr, byref res as PdfFinderResult)
	var rect = txt->getRect(res.index, res.count)
	highlight(rect)
	delete rect
end sub

function PdfPage.getHandle() as FPDF_PAGE
	return page
end function

'''''
constructor PdfDoc()
	this.doc = FPDF_CreateNewDocument()
end constructor

constructor PdfDoc(doc as FPDF_DOCUMENT)
	this.doc = doc
end constructor

constructor PdfDoc(path as string)
	doc = FPDF_LoadDocument(path, null)
end constructor

destructor PdfDoc()
	if doc <> null then
		FPDF_CloseDocument(doc)
	end if
end destructor

property PdfDoc.count as integer
	return FPDF_GetPageCount(doc)
end property

property PdfDoc.page(index as integer) as PdfPage ptr
	var pg = FPDF_LoadPage(doc, index)
	return new PdfPage(pg)
end property

sub PdfDoc.importPages(src as PdfDoc ptr, fromPage as integer, toPage as integer)
	var range = fromPage & "-" & toPage
	FPDF_ImportPages(doc, src->doc, range, count())
end sub

sub PdfDoc.importPages(src as PdfDoc ptr, range as string)
	FPDF_ImportPages(doc, src->doc, range, count())
end sub

static function PdfDoc.blockWriterCb(byval pThis as PdfFileWriter ptr, byval pData as const any ptr, byval size as culong) as long
	put #pThis->bf, , *cast(const byte ptr, pData), size
	return true
end function

function PdfDoc.saveTo(path as string, version as integer) as boolean
	dim pl as PdfFileWriter
	pl.version = version
	pl.WriteBlock = cast(any ptr, @blockWriterCb)
	pl.bf = FreeFile
	if open(path for binary access write as #pl.bf) <> 0 then
		return false
	end if
	
	FPDF_SaveWithVersion(doc, @pl, 0, version)
	
	close #pl.bf
	
	return true
end function

function PdfDoc.getDoc() as FPDF_DOCUMENT
	return this.doc
end function

#ifdef WITH_PARSER
''FIXME: Windows-only
private function utf8ToUtf16le(src as zstring ptr) as ushort ptr
	var bytes = len(*src)
	var dst = allocate((bytes+1) * len(ushort))
	var srcp = src
	var srcleft = bytes
	var dstp = dst
	var dstleft = bytes*2
	iconv(cd8to16le, @srcp, @srcleft, @dstp, @dstleft)
	*cast(ushort ptr, dstp) = 0
	function = dst
end function

private function utf16leToUtf8(src as ushort ptr, bytes as integer) as zstring ptr
	var dst = allocate(bytes+1)
	var srcp = cast(any ptr, src)
	var srcleft = bytes
	var dstp = dst
	var dstleft = bytes
	iconv(cd16leto8, @srcp, @srcleft, @dstp, @dstleft)
	*cast(zstring ptr, dstp) = 0
	function = dst
end function

private function hex2ushort(h as const zstring ptr) as uinteger
	
	dim res as uinteger
	
	for i as integer = 0 to 3
		if h[i] <= asc("9") then
			res = (res shl 4) or (h[i] - asc("0"))
		elseif h[i] >= asc("a") and h[i] <= asc("f") then
			res = (res shl 4) or (h[i] - asc("a") + 10)
		else
			res = (res shl 4) or (h[i] - asc("A") + 10)
		end if
	next 
	
	return res
end function

private sub splitStr(text as zstring ptr, delim as zstring ptr, toArray() as zstring ptr)

	var items = 10
	redim pos_(0 to items-1) as integer
	
	var x = 0
	var p = 0
	do 
		x = instr(x + 1, *text, *delim)
		if( x > 0 ) then
			if( p >= items ) then
				items += 10
				redim preserve pos_(0 to items-1)
			end if
			pos_(p) = x
		end if
		p += 1
	loop until x = 0
	
	var cnt = p - 1
	if( cnt = 0 ) then
		redim toArray(0 to 0)
		toArray(0) = allocate((len(*text)+1) * len(zstring))
		*toArray(0) = *text
		return
	end if
	
	redim toArray(0 to cnt)
	
	dim buff as string
	buff = left(*text, pos_(0) - 1)
	toArray(0) = allocate((len(buff)+1) * len(zstring))
	*toArray(0) = buff
	p = 1
	do until p = cnt
		buff = mid(*text, pos_(p - 1) + 1, pos_(p) - pos_(p - 1) - 1)
		toArray(p) = allocate((len(buff)+1) * len(zstring))
		*toArray(p) = buff
		p += 1
	loop
	
	buff = mid(*text, pos_(cnt - 1) + 1)
	toArray(cnt) = allocate((len(buff)+1) * len(zstring))
	*toArray(cnt) = buff
   
end sub

private function typeToString(type_ as PdfTemplateNodeType) as string
	select case as const type_
	case PdfTemplateNodeType.INVALID
		return "INVALID"
	case PdfTemplateNodeType.DOCUMENT
		return "DOCUMENT"
	case PdfTemplateNodeType.PAGE
		return "PAGE"
	case PdfTemplateNodeType.GROUP
		return "GROUP"
	case PdfTemplateNodeType.TEMPLATE
		return "TEMPLATE"
	case PdfTemplateNodeType.FILL
		return "FILL"
	case PdfTemplateNodeType.STROKE
		return "STROKE"
	case PdfTemplateNodeType.MOVE_TO
		return "MOVE_TO"
	case PdfTemplateNodeType.LINE_TO
		return "LINE_TO"
	case PdfTemplateNodeType.BEZIER_TO
		return "BEZIER_TO"
	case PdfTemplateNodeType.CLOSE_PATH
		return "CLOSE_PATH"
	case PdfTemplateNodeType.TEXT
		return "TEXT"
	end select
end function

'''''
constructor PdfTemplateNode()
end constructor

constructor PdfTemplateNode(type_ as PdfTemplateNodeType)
	this.type_ = type_
end constructor

constructor PdfTemplateNode(type_ as PdfTemplateNodeType, parent as PdfTemplateNode ptr)
	this.type_ = type_
	this.parent = parent
	if parent then
		if parent->head = null then
			parent->head = @this
			parent->tail = @this
		else
			parent->tail->next_ = @this
			parent->tail = @this
		end if
	end if
end constructor

constructor PdfTemplateNode(type_ as PdfTemplateNodeType, id as string, idDict as TDict ptr, parent as PdfTemplateNode ptr)
	constructor(type_, parent)
	if len(id) > 0 then
		this.id = id
		if (*idDict)[id] = null then
			idDict->add(this.id, @this)
		end if
	end if
end constructor

destructor PdfTemplateNode()
	if obj <> null then
		FPDFPageObj_Destroy(obj)
	end if
	
	var child = this.head
	do while child <> null
		var next_ = child->next_
		delete child
		child = next_
	loop
end destructor

function PdfTemplateNode.clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode_ ptr) as PdfTemplateNode ptr
	var dup = new PdfTemplateNode(type_, id, page->getIdDict(), parent)
	cloneChildren(dup, page)
	return dup
end function

sub PdfTemplateNode.cloneChildren(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode_ ptr)
	var child = this.head
	do while child <> null
		child->clone(parent, page)
		child = child->next_
	loop
end sub

function PdfTemplateNode.lookupAttrib(name_ as string, byref type_ as PdfTemplateAttribType) as any ptr
	select case name_
	case "hidden"
		type_ = PdfTemplateAttribType.TP_BOOLEAN
		return @this.hidden
	end select
	return null
end function

sub PdfTemplateNode.setAttrib(name_ as string, value as boolean)
	dim type_ as PdfTemplateAttribType
	var attrib = lookupAttrib(name_, type_)
	if attrib <> null then
		if type_ = PdfTemplateAttribType.TP_BOOLEAN then
			*cast(boolean ptr, attrib) = value
		end if
	end if
end sub

sub PdfTemplateNode.setAttrib(name_ as string, value as integer)
	dim type_ as PdfTemplateAttribType
	var attrib = lookupAttrib(name_, type_)
	if attrib <> null then
		if type_ = PdfTemplateAttribType.TP_INTEGER then
			*cast(integer ptr, attrib) = value
		end if
	end if
end sub

sub PdfTemplateNode.setAttrib(name_ as string, value as single)
	dim type_ as PdfTemplateAttribType
	var attrib = lookupAttrib(name_, type_)
	if attrib <> null then
		if type_ = PdfTemplateAttribType.TP_SINGLE then
			*cast(single ptr, attrib) = value
		end if
	end if
end sub

sub PdfTemplateNode.setAttrib(name_ as string, value as double)
	dim type_ as PdfTemplateAttribType
	var attrib = lookupAttrib(name_, type_)
	if attrib <> null then
		if type_ = PdfTemplateAttribType.TP_DOUBLE then
			*cast(double ptr, attrib) = value
		end if
	end if
end sub

sub PdfTemplateNode.setAttrib(name_ as string, value as wstring ptr)
	dim type_ as PdfTemplateAttribType
	var attrib = cast(any ptr ptr, lookupAttrib(name_, type_))
	if attrib <> null then
		if type_ = PdfTemplateAttribType.TP_WSTRINGPTR then
			if *attrib <> null then
				deallocate(*attrib)
			end if
			if value <> null andalso len(*value) > 0 then
				*attrib = allocate((len(*value)+1) * len(wstring))
				**cast(wstring ptr ptr, attrib) = *value
			else
				*attrib = null
			end if
		end if
	end if
end sub

sub PdfTemplateNode.setAttrib(name_ as string, value as zstring ptr)
	dim type_ as PdfTemplateAttribType
	var attrib = cast(any ptr ptr, lookupAttrib(name_, type_))
	if attrib <> null then
		if type_ = PdfTemplateAttribType.TP_WSTRINGPTR then
			if *attrib <> null then
				deallocate(*attrib)
			end if
			if value <> null then
				if len(*value) > 0 then
					*attrib = utf8ToUtf16le(value)
				else
					*attrib = null
				end if
			end if
		end if
	end if
end sub

sub PdfTemplateNode.translate(xi as single, yi as single)
	var child = this.head
	do while child <> null
		child->translate(xi, yi)
		child = child->next_
	loop
end sub

sub PdfTemplateNode.translateX(xi as single)
	var child = this.head
	do while child <> null
		child->translateX(xi)
		child = child->next_
	loop
end sub

sub PdfTemplateNode.translateY(yi as single)
	var child = this.head
	do while child <> null
		child->translateY(yi)
		child = child->next_
	loop
end sub

function PdfTemplateNode.emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	return null
end function

function PdfTemplateNode.emitAndInsert(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	if hidden then
		return null
	end if
		
	obj = this.emit(doc, page, parent)

	emitChildren(doc, page, obj)
	
	if obj <> null then
		FPDFPage_InsertObject(page, obj)
	end if
end function

sub PdfTemplateNode.emitChildren(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT)
	static d as integer = 0
	'print space(d*4); typeToString(this.type_)
	d += 1
	var child = this.head
	do while child <> null
		child->emitAndInsert(doc, page, parent)
		child = child->next_
	loop
	d -= 1
end sub

function PdfTemplateNode.getChild(id as string) as PdfTemplateNode ptr 
	var child = this.head
	do while child <> null
		if child->id = id then
			return child
		end if
		child = child->next_
	loop
	return null
end function

function PdfTemplateNode.getHead() as PdfTemplateNode ptr
	return this.head
end function

function PdfTemplateNode.getTail() as PdfTemplateNode ptr
	return this.tail
end function

function PdfTemplateNode.getNext() as PdfTemplateNode ptr
	return this.next_
end function

'''''
constructor PdfRGB(r as ulong, g as ulong, b as ulong, a as ulong)
	this.r = r
	this.g = g
	this.b = b
	this.a = a
end constructor

function PdfRGB.clone() as PdfRGB ptr
	return new PdfRGB(r, g, b, a)
end function

'''''
private function cloneTranform(transf as FS_MATRIX ptr) as FS_MATRIX ptr
	var dup = new FS_MATRIX
	dup->a = transf->a
	dup->b = transf->b
	dup->c = transf->c
	dup->d = transf->d
	dup->e = transf->e
	dup->f = transf->f
	return dup
end function

'''''
constructor PdfTemplateStrokeNode()
	base()
end constructor

constructor PdfTemplateStrokeNode(width_ as single, miterlin as single, join as integer, cap as integer, colorspace as integer, color_ as PdfRGB ptr, transf as FS_MATRIX ptr, parent as PdfTemplateNode ptr)
	base(PdfTemplateNodeType.STROKE, parent)
	this.width_ = width_
	this.miterlin = miterlin
	this.join = join
	this.cap = cap
	this.colorspace = colorspace
	this.color_ = color_
	this.transf = transf
end constructor

function PdfTemplateStrokeNode.clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	var color2 = iif(color_ <> null, color_->clone(), null)
	var transf2 = iif(transf <> null, cloneTranform(transf), null)
	var dup = new PdfTemplateStrokeNode(width_, miterlin, join, cap, colorspace, color2, transf2, parent)
	cloneChildren(dup, page)
	return dup
end function

destructor PdfTemplateStrokeNode()
	if color_ <> null then
		delete color_
	end if
	if transf <> null then
		delete transf
	end if
end destructor

function PdfTemplateStrokeNode.emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	var path = FPDFPageObj_CreateNewPath(0, 0)
	
	FPDFPath_SetDrawMode(path, FPDF_FILLMODE_NONE, 1)
	
	if color_ <> null then
		FPDFPageObj_SetStrokeColor(path, color_->r, color_->g, color_->b, color_->a)
	end if
	
	if width_ > 0 then
		FPDFPageObj_SetStrokeWidth(path, width_)
	end if
	
	FPDFPageObj_SetLineCap(path, cap)
	FPDFPageObj_SetLineJoin(path, join)
	
	if transf <> null then
		'FPDFPageObj_Transform(path, transf->a, transf->b, transf->c, transf->d, transf->e, transf->f)
	end if
	
	return path
end function

'''''
constructor PdfTemplateFillNode()
	base()
end constructor

constructor PdfTemplateFillNode(mode as integer, colorspace as integer, color_ as PdfRGB ptr, transf as FS_MATRIX ptr, parent as PdfTemplateNode ptr)
	base(PdfTemplateNodeType.FILL, parent)
	this.mode = mode
	this.colorspace = colorspace
	this.color_ = color_
	this.transf = transf
end constructor

constructor PdfTemplateFillNode(mode as integer, colorspace as integer, r as ulong, g as ulong, b as ulong, parent as PdfTemplateNode ptr)
	base(PdfTemplateNodeType.FILL, parent)
	this.mode = mode
	this.colorspace = colorspace
	this.color_ = new PdfRGB(r, g, b)
end constructor

function PdfTemplateFillNode.clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	var color2 = iif(color_ <> null, color_->clone(), null)
	var transf2 = iif(transf <> null, cloneTranform(transf), null)
	var dup = new PdfTemplateFillNode(mode, colorspace, color2, transf2, parent)
	cloneChildren(dup, page)
	return dup
end function

destructor PdfTemplateFillNode()
	if color_ <> null then
		delete color_
	end if
	if transf <> null then
		delete transf
	end if
end destructor

function PdfTemplateFillNode.emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	var path = FPDFPageObj_CreateNewPath(0, 0)
	
	FPDFPath_SetDrawMode(path, mode, 0)
	
	if color_ <> null then
		FPDFPageObj_SetFillColor(path, color_->r, color_->g, color_->b, color_->a)
	end if
	
	if transf <> null then
		'FPDFPageObj_Transform(path, transf->a, transf->b, transf->c, transf->d, transf->e, transf->f)
	end if
	
	return path
end function

'''''
constructor PdfTemplateMoveToNode(x as single, y as single, parent as PdfTemplateNode ptr)
	base(PdfTemplateNodeType.MOVE_TO, parent)
	this.x = x 
	this.y = y
end constructor

function PdfTemplateMoveToNode.clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	var dup = new PdfTemplateMoveToNode(x, y, parent)
	cloneChildren(dup, page)
	return dup
end function

sub PdfTemplateMoveToNode.translate(xi as single, yi as single)
	this.x += xi
	this.y += yi
end sub

sub PdfTemplateMoveToNode.translateX(xi as single)
	this.x += xi
end sub

sub PdfTemplateMoveToNode.translateY(yi as single)
	this.y += yi
end sub

function PdfTemplateMoveToNode.emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	FPDFPath_MoveTo(parent, this.x, this.y)
	return null
end function

'''''
constructor PdfTemplateLineToNode(x as single, y as single, parent as PdfTemplateNode ptr)
	base(PdfTemplateNodeType.LINE_TO, parent)
	this.x = x 
	this.y = y
end constructor

function PdfTemplateLineToNode.clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	var dup = new PdfTemplateLineToNode(x, y, parent)
	cloneChildren(dup, page)
	return dup
end function

sub PdfTemplateLineToNode.translate(xi as single, yi as single)
	this.x += xi
	this.y += yi
end sub

sub PdfTemplateLineToNode.translateX(xi as single)
	this.x += xi
end sub

sub PdfTemplateLineToNode.translateY(yi as single)
	this.y += yi
end sub

function PdfTemplateLineToNode.emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	FPDFPath_LineTo(parent, this.x, this.y)
	return null
end function

'''''
constructor PdfTemplateBezierToNode(x1 as single, y1 as single, x2 as single, y2 as single, x3 as single, y3 as single, parent as PdfTemplateNode ptr)
	base(PdfTemplateNodeType.BEZIER_TO, parent)
	this.x1 = x1 
	this.y1 = y1
	this.x2 = x2 
	this.y2 = y2
	this.x3 = x3 
	this.y3 = y3
end constructor

function PdfTemplateBezierToNode.clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	var dup = new PdfTemplateBezierToNode(x1, y1, x2, y2, x3, y3, parent)
	cloneChildren(dup, page)
	return dup
end function

sub PdfTemplateBezierToNode.translate(xi as single, yi as single)
	this.x1 += xi
	this.y1 += yi
	this.x2 += xi
	this.y2 += yi
	this.x3 += xi
	this.y3 += yi
end sub

sub PdfTemplateBezierToNode.translateX(xi as single)
	this.x1 += xi
	this.x2 += xi
	this.x3 += xi
end sub

sub PdfTemplateBezierToNode.translateY(yi as single)
	this.y1 += yi
	this.y2 += yi
	this.y3 += yi
end sub

function PdfTemplateBezierToNode.emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	FPDFPath_BezierTo(parent, this.x1, this.y1, this.x2, this.y2, this.x3, this.y3)
	return null
end function

'''''
constructor PdfTemplateClosePathNode(parent as PdfTemplateNode ptr)
	base(PdfTemplateNodeType.CLOSE_PATH, parent)
end constructor

function PdfTemplateClosePathNode.clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	var dup = new PdfTemplateClosePathNode(parent)
	cloneChildren(dup, page)
	return dup
end function

function PdfTemplateClosePathNode.emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	FPDFPath_Close(parent)
	return null
end function

'''''
constructor PdfTemplateTextNode(id as string, idDict as TDict ptr, font as string, size as single, mode as FPDF_TEXT_RENDERMODE, x as single, y as single, width_ as single, align as PdfTextAlignment, text as wstring ptr, colorspace as integer, color_ as PdfRGB ptr, transf as FS_MATRIX ptr, parent as PdfTemplateNode ptr)
	base(PdfTemplateNodeType.TEXT, id, idDict, parent)
	this.font = font
	this.size = size
	this.mode = mode
	this.x = x
	this.y = y
	this.width_ = width_
	this.align = align
	this.text = text
	this.colorspace = colorspace
	this.color_ = color_
	this.transf = transf
end constructor

function PdfTemplateTextNode.clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	dim text2 as wstring ptr
	if text <> null then
		text2 = cast(wstring ptr, allocate((len(*text)+1) * len(wstring)))
		*text2 = *text
	end if
	var color2 = iif(color_ <> null, color_->clone(), null)
	var transf2 = iif(transf <> null, cloneTranform(transf), null)
	var dup = new PdfTemplateTextNode(id, page->getIdDict(), font, size, mode, x, y, width_, align, text2, colorspace, color2, transf2, parent)
	cloneChildren(dup, page)
	return dup
end function

sub PdfTemplateTextNode.translate(xi as single, yi as single)
	this.x += xi
	this.y += yi
end sub

sub PdfTemplateTextNode.translateX(xi as single)
	this.x += xi
end sub

sub PdfTemplateTextNode.translateY(yi as single)
	this.y += yi
end sub

function PdfTemplateTextNode.lookupAttrib(name_ as string, byref type_ as PdfTemplateAttribType) as any ptr
	select case name_
	case "text"
		type_ = PdfTemplateAttribType.TP_WSTRINGPTR
		return @this.text
	case "x"
		type_ = PdfTemplateAttribType.TP_SINGLE
		return @this.x
	case "y"
		type_ = PdfTemplateAttribType.TP_SINGLE
		return @this.y
	case else
		return base.lookupAttrib(name_, type_)
	end select
end function

destructor PdfTemplateTextNode()
	if text <> null then
		deallocate text
	end if
	if color_ <> null then
		delete color_
	end if
	if transf <> null then
		delete transf
	end if
end destructor

function PdfTemplateTextNode.emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	if text = null orelse len(*text) = 0 then
		return null
	end if
	
	var fon = FPDFText_LoadStandardFont(doc, font)
	var tex = FPDFPageObj_CreateTextObj(doc, fon, size)
	FPDFText_SetText(tex, text)
	var xpos = x
	if align <> PdfTextAlignment.TA_LEFT then
		dim left as single = any, bottom as single = any, right as single = any, top as single = any
		FPDFPageObj_GetBounds(tex, @left, @bottom, @right, @top)
		var dx = (right - left) + 1
		if align = PdfTextAlignment.TA_CENTER then
			xpos += (width_ / 2) - (dx / 2)
		else
			xpos += width_ - dx
		end if
	end if
	FPDFPageObj_Transform(tex, 1, 0, 0, 1, xpos, y)
	FPDFPageObj_SetFillColor(tex, color_->r, color_->g, color_->b, color_->a)
	FPDFTextObj_SetTextRenderMode(tex, mode)
	return tex
end function

'''''
constructor PdfRectCoords()
end constructor

constructor PdfRectCoords(left as double, top as double, right as double, bottom as double)
	this.left = left
	this.right = right
	this.top = top
	this.bottom = bottom
end constructor

function PdfRectCoords.clone() as PdfRectCoords ptr
	return new PdfRectCoords(left, top, right, bottom)
end function

'''''
constructor PdfTemplateGroupNode(bbox as PdfRectCoords ptr, isolated as boolean, knockout as boolean, blendMode as zstring ptr, alpha as single, parent as PdfTemplateNode ptr)
	base(PdfTemplateNodeType.GROUP, parent)
	this.bbox = bbox
	this.isolated = isolated
	this.knockout = knockout
	this.blendMode = blendMode
	this.alpha = alpha
end constructor

function PdfTemplateGroupNode.clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	var bbox2 = iif(bbox <> null, bbox->clone(), null)
	var dup = new PdfTemplateGroupNode(bbox2, isolated, knockout, blendMode, alpha, parent)
	cloneChildren(dup, page)
	return dup
end function

destructor PdfTemplateGroupNode()
	if bbox <> null then
		delete bbox
	end if
end destructor

function PdfTemplateGroupNode.emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	return null
end function

'''''
constructor PdfTemplateTemplateNode(id as string, idDict as TDict ptr, parent as PdfTemplateNode ptr, hidden as boolean)
	base(PdfTemplateNodeType.TEMPLATE, id, idDict, parent)
	base.hidden = hidden
end constructor

function PdfTemplateTemplateNode.clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	var dup = new PdfTemplateTemplateNode(id, page->getIdDict(), parent, hidden)
	cloneChildren(dup, page)
	return dup
end function

function PdfTemplateTemplateNode.emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	return null
end function

'''''
constructor PdfTemplatePageNode(x1 as single, y1 as single, x2 as single, y2 as single, parent as PdfTemplateNode ptr)
	base(PdfTemplateNodeType.PAGE, parent)
	this.x1 = x1
	this.y1 = y1
	this.x2 = x2
	this.y2 = y2
	idDict.init(2^10)
end constructor

destructor PdfTemplatePageNode()
	idDict.end_()
end destructor

function PdfTemplatePageNode.clone() as PdfTemplatePageNode ptr
	var dup = new PdfTemplatePageNode(x1, y1, x2, y2, null)
	cloneChildren(dup, dup)
	return dup
end function

sub PdfTemplatePageNode.emit(doc as FPDF_DOCUMENT, index as integer, flush_ as boolean)
	if hidden then
		return
	end if
	
	page = FPDFPage_New(doc, index, x2 - x1, y2 - y1)
	FPDFPage_SetMediaBox(page, x1, y1, x2, y2)
	FPDFPage_SetCropBox(page, x1, y1, x2, y2)
	
	emitChildren(doc, page, null)
	
	if flush_ then
		FPDFPage_GenerateContent(page)
		page = null
	end if
end sub

sub PdfTemplatePageNode.emit(doc as PdfDoc ptr, index as integer, flush_ as boolean)
	emit(doc->getDoc(), index, flush_)
end sub

sub PdfTemplatePageNode.flush()
	if page <> null then
		FPDFPage_GenerateContent(page)
		page = null
	end if
end sub

function PdfTemplatePageNode.getIdDict() as TDict ptr
	return @idDict
end function

function PdfTemplatePageNode.getNode(id as string) as PdfTemplateNode ptr
	return cast(PdfTemplateNode ptr, idDict[id])
end function

'''''
constructor PdfTemplate(buff as zstring ptr, size as integer, encoding_ as zstring ptr)
	index = 0
	reader = xmlReaderForMemory(buff, size, null, encoding_, XML_PARSE_NOBLANKS)
end constructor

constructor PdfTemplate(path as string)
	index = 0
	reader = xmlReaderForFile(path, null, XML_PARSE_NOBLANKS)
end constructor

destructor PdfTemplate()
	delete root
	if reader <> null then
		xmlFreeTextReader(reader)
	end if
end destructor

function PdfTemplate.getXmlConstName() as string
	var s = xmlTextReaderConstName(reader)
	if s = null then
		return ""
	end if
	function = trim(*cast(const zstring ptr, s))
end function

function PdfTemplate.getXmlAttrib(name_ as zstring ptr) as string
	var s = xmlTextReaderGetAttribute(reader, cast(xmlChar ptr, name_))
	if s = null then
		return ""
	end if
	function = trim(*cast(const zstring ptr, s))
	deallocate s
end function

function PdfTemplate.getXmlAttribAsLong(name_ as zstring ptr) as longint
	function = vallng(getXmlAttrib(name_))
end function

function PdfTemplate.getXmlAttribAsDouble(name_ as zstring ptr) as double
	function = val(getXmlAttrib(name_))
end function

function PdfTemplate.getXmlAttribAsLongArray(name_ as zstring ptr, toArr() as longint, delim as string) as integer
	var value = getXmlAttrib(name_)
	
	if len(value) = 0 then
		return 0
	end if
	
	dim strArr() as zstring ptr
	splitStr(value, delim, strArr())
	
	var cnt = ubound(strArr) + 1
	
	redim toArr(0 to cnt-1)
	
	for i as integer = 0 to cnt-1
		toArr(i) = vallng(*strArr(i))
		deallocate strArr(i)
	next
	
	return cnt
	
end function

function PdfTemplate.getXmlAttribAsDoubleArray(name_ as zstring ptr, toArr() as double, delim as string) as integer
	var value = getXmlAttrib(name_)
	
	if len(value) = 0 then
		return 0
	end if
	
	dim strArr() as zstring ptr
	splitStr(value, delim, strArr())
	
	var cnt = ubound(strArr) + 1
	
	redim toArr(0 to cnt-1)
	
	for i as integer = 0 to cnt-1
		toArr(i) = val(*strArr(i))
		deallocate strArr(i)
	next
	
	return cnt
	
end function

function PdfTemplate.parseColorAttrib() as PdfRGB ptr
	dim colorArr() as double
	
	var colorCnt = getXmlAttribAsDoubleArray("color", colorArr())
	if colorCnt = 3 then
		return new PdfRGB(colorArr(0) * 255, colorArr(1) * 255, colorArr(2) * 255, 255)
	end if
	
	return null
end function

function PdfTemplate.parseTranformAttrib() as FS_MATRIX ptr
	dim transfArr() as double
	
	var transfCnt = getXmlAttribAsDoubleArray("transform", transfArr())
	
	if transfCnt = 6 then
		var transf = new FS_MATRIX
		transf->a = transfArr(0)
		transf->b = transfArr(1)
		transf->c = transfArr(2)
		transf->d = transfArr(3)
		transf->e = transfArr(4)
		transf->f = transfArr(5)
		return transf
	end if
	
	return null
end function

function PdfTemplate.parseColorspaceAttrib() as integer

	select case getXmlAttrib("colorspace")
	case "DeviceGray"
		return FPDF_COLORSPACE_DEVICEGRAY
	case "DeviceCMYK"
		return FPDF_COLORSPACE_DEVICECMYK
	case "None"
		return FPDF_COLORSPACE_UNKNOWN
	case else
		return FPDF_COLORSPACE_DEVICERGB
	end select
	
end function

function PdfTemplate.parseFill(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateFillNode ptr
	
	var mode = FPDF_FILLMODE_NONE
	select case getXmlAttrib("winding")
	case "eofill"
		mode = FPDF_FILLMODE_WINDING
	case "nonzero"
		mode = FPDF_FILLMODE_ALTERNATE
	end select
	
	var colorspace = parseColorspaceAttrib()
	var color_ = parseColorAttrib()
	if color_ <> null then
		var attrib = getXmlAttrib("alpha")
		if len(attrib) > 0 then
			color_->a = val(attrib) * 255
		end if
	end if
	var transf = parseTranformAttrib()
	
	return new PdfTemplateFillNode(mode, colorspace, color_, transf, parent)
	
end function

function PdfTemplate.parseStroke(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateStrokeNode ptr
	
	var width_ = getXmlAttribAsDouble("linewidth")
	var miterlin = getXmlAttribAsDouble("miterlimit")
	var join = getXmlAttribAsLong("linejoin")
	dim cap as integer = FPDF_LINECAP_BUTT
	var attrib = getXmlAttrib("linecap")
	if len(attrib) >= 1 then
		cap = valint(left(attrib, 1))
	end if
	var colorspace = parseColorspaceAttrib()
	var color_ = parseColorAttrib()
	if color_ <> null then
		var attrib = getXmlAttrib("alpha")
		if len(attrib) > 0 then
			color_->a = val(attrib) * 255
		end if
	end if
	var transf = parseTranformAttrib()
	
	return new PdfTemplateStrokeNode(width_, miterlin, join, cap, colorspace, color_, transf, parent)
	
end function

function PdfTemplate.parseMoveTo(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateMoveToNode ptr
	
	var x = getXmlAttribAsDouble("x")
	var y = getXmlAttribAsDouble("y")
	
	return new PdfTemplateMoveToNode(x, y, parent)
	
end function

function PdfTemplate.parseLineTo(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateLineToNode ptr
	
	var x = getXmlAttribAsDouble("x")
	var y = getXmlAttribAsDouble("y")
	
	return new PdfTemplateLineToNode(x, y, parent)
	
end function

function PdfTemplate.parseBezierTo(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateBezierToNode ptr
	
	var x1 = getXmlAttribAsDouble("x1")
	var y1 = getXmlAttribAsDouble("y1")
	var x2 = getXmlAttribAsDouble("x2")
	var y2 = getXmlAttribAsDouble("y2")
	var x3 = getXmlAttribAsDouble("x3")
	var y3 = getXmlAttribAsDouble("y3")
	
	return new PdfTemplateBezierToNode(x1, y1, x2, y2, x3, y3, parent)
	
end function

function PdfTemplate.parseClosePath(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateClosePathNode ptr
	
	return new PdfTemplateClosePathNode(parent)
	
end function

function PdfTemplate.parseGroup(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateGroupNode ptr
	dim as PdfRectCoords ptr bbox
	
	dim bboxArr() as double
	var bboxCnt = getXmlAttribAsDoubleArray("bbox", bboxArr())
	if bboxCnt = 4 then
		bbox = new PdfRectCoords
		bbox->left = bboxArr(0)
		bbox->top = bboxArr(1)
		bbox->right = bboxArr(2)
		bbox->bottom = bboxArr(3)
	end if
	
	var isolated = getXmlAttribAsLong("isolated")
	var knockout = getXmlAttribAsLong("knockout")
	var blendMode = getXmlAttrib("blendmode")
	var alpha = getXmlAttribAsDouble("alpha")
	
	return new PdfTemplateGroupNode(bbox, isolated, knockout, blendMode, alpha, parent)
	
end function

function PdfTemplate.parseTemplate(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateTemplateNode ptr
	
	var id = getXmlAttrib("id")
	var attrib = getXmlAttrib("hidden")
	var hidden = true
	if len(attrib) > 0 then
		hidden = attrib = "true" orelse attrib = "1"
	end if
	
	return new PdfTemplateTemplateNode(id, page->getIdDict(), parent, hidden)
	
end function

function PdfTemplate.parseText(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateTextNode ptr
	
	var colorspace = parseColorspaceAttrib()
	var color_ = parseColorAttrib()
	if color_ <> null then
		var attrib = getXmlAttrib("alpha")
		if len(attrib) > 0 then
			color_->a = val(attrib) * 255
		end if
	end if
	var transf = parseTranformAttrib()
	
	dim id as string
	dim font as string
	var size = 0.0
	var x = 0.0, y = 0.0, width_ = 0.0
	var align = PdfTextAlignment.TA_LEFT
	var mode = FPDF_TEXTRENDERMODE_FILL
	''FIXME: Windows-only
	var text = cast(ushort ptr, null)
	var g = 0
	
	do while xmlTextReaderRead(reader) = 1 
		select case xmlTextReaderNodeType(reader)
		case XML_READER_TYPE_ELEMENT
			var name_ = getXmlConstName()
			select case name_
			case "g"
				if g < 1024 then
					if g = 0 then
						text = callocate((1024+1) * len(ushort))
						var attrib = getXmlAttrib("x")
						if len(attrib) > 0 then
							x = val(attrib)
							y = getXmlAttribAsDouble("y")
						end if
					end if
					
					var code = getXmlAttrib("unicode")
					select case as const len(code)
					case 0
						text[g] = asc(" ")
					case 1
						text[g] = strptr(code)[0]
					case 6
						text[g] = hex2ushort(strptr(code) + 2)
					end select
					
					g += 1
				end if
				
			case "span"
				id = getXmlAttrib("id")
				font = getXmlAttrib("font")
				var trm = getXmlAttrib("trm")
				size = val(left(trm, instr(trm, " ")))
				mode = getXmlAttribAsLong("wmode")
				x = getXmlAttribAsDouble("x")
				y = getXmlAttribAsDouble("y")
				width_ = getXmlAttribAsDouble("width")
				var attrib = getXmlAttrib("align")
				if len(attrib) > 0 then
					select case attrib
					case "center"
						align = PdfTextAlignment.TA_CENTER
					case "right"
						align = PdfTextAlignment.TA_RIGHT
					end select
				end if
				g = 0
			end select
			
		case XML_READER_TYPE_TEXT
			if g = 0 then
				var value = cast(zstring ptr, xmlTextReaderConstValue(reader))
				if value <> null andalso len(*value) > 0 then
					text = utf8ToUtf16le(value)
				else
					text = null
				end if
				g = 1024
			end if
		
		case XML_READER_TYPE_END_ELEMENT
			var name_ = getXmlConstName()
			if name_ = "span" then
				exit do
			end if
		end select
	loop
		
	return new PdfTemplateTextNode(id, page->getIdDict(), font, size, mode, x, y, width_, align, text, colorspace, color_, transf, parent)
	
end function

function PdfTemplate.parseObject(tag as zstring ptr, parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	
	dim obj as PdfTemplateNode ptr
	select case *tag
	case "fill_text"
		obj = parseText(parent, page)
	case "fill_path"
		obj = parseFill(parent, page)
	case "stroke_path"
		obj = parseStroke(parent, page)
	case "moveto"
		obj = parseMoveTo(parent, page)
	case "lineto"
		obj = parseLineTo(parent, page)
	case "curveto"
		obj = parseBezierTo(parent, page)
	case "closepath"
		obj = parseClosePath(parent, page)
	case "group"
		obj = parseGroup(parent, page)
	case "template"
		obj = parseTemplate(parent, page)
	end select

	if xmlTextReaderIsEmptyElement(reader) then
		return obj
	end if

	do while xmlTextReaderRead(reader) = 1 
		select case xmlTextReaderNodeType(reader)
		case XML_READER_TYPE_ELEMENT
			var name_ = getXmlConstName()
			parseObject(name_, obj, page)
			
		case XML_READER_TYPE_END_ELEMENT
			var name_ = getXmlConstName()
			if name_ = *tag then
				exit do
			end if
		end select
	loop
	
	return obj

end function

sub PdfTemplate.parsePage(parent as PdfTemplateNode ptr)

	dim mediaboxArr() as double
	var arrCnt = getXmlAttribAsDoubleArray("mediabox", mediaboxArr())
	var x1 = 0.0, y1 = 0.0, x2 = 0.0, y2 = 0.0
	if arrCnt = 4 then
		x1 = mediaboxArr(0)
		y1 = mediaboxArr(1)
		x2 = mediaboxArr(2)
		y2 = mediaboxArr(3)
	end if
	
	var page = new PdfTemplatePageNode(x1, y1, x2, y2, parent)

	do while xmlTextReaderRead(reader) = 1 
		select case xmlTextReaderNodeType(reader)
		case XML_READER_TYPE_ELEMENT
			var name_ = getXmlConstName()
			parseObject(name_, page, page)
			
		case XML_READER_TYPE_END_ELEMENT
			var name_ = getXmlConstName()
			if name_ = "page" then
				exit do
			end if
		end select
	loop

end sub

sub PdfTemplate.parseDocument(parent as PdfTemplateNode ptr)

	version = getXmlAttribAsDouble("version") * 10

	do while xmlTextReaderRead(reader) = 1 
		
		select case xmlTextReaderNodeType(reader)
		case XML_READER_TYPE_ELEMENT
			var name_ = getXmlConstName()
			if name_ = "page" then
				parsePage(root)
			end if
			
		case XML_READER_TYPE_END_ELEMENT
			var name_ = getXmlConstName()
			if name_ = "document" then
				exit do
			end if
		end select
	loop

end sub

function PdfTemplate.load() as boolean
	do while xmlTextReaderRead(reader) = 1 
		if xmlTextReaderNodeType(reader) = XML_READER_TYPE_ELEMENT then
			var name_ = getXmlConstName()
			if name_ = "document" then
				root = new PdfTemplateNode(PdfTemplateNodeType.DOCUMENT)
				parseDocument(root)
				return true
			end if
		end if
	loop
	return false
end function

sub PdfTemplate.emitTo(doc as PdfDoc ptr, flush_ as boolean)
	var page = root->getHead()
	do while page <> null
		cast(PdfTemplatePageNode ptr, page)->emit(doc->getDoc(), index, flush_)
		index += 1
		page = page->getNext()
	loop
end sub

sub PdfTemplate.flush()
	var page = root->getHead()
	do while page <> null
		cast(PdfTemplatePageNode ptr, page)->flush()
		page = page->getNext()
	loop
end sub

function PdfTemplate.clonePage(index as integer) as PdfTemplatePageNode ptr
	var page = root->getHead()
	var cnt = 0
	do while page <> null
		if cnt = index then
			return cast(PdfTemplatePageNode ptr, page)->clone()
		end if
		page = page->getNext()
		cnt += 1
	loop
	return null
end function

function PdfTemplate.simplifyXml(inFile as string, outFile as string) as boolean
	
	var reader = xmlReaderForFile(inFIle, null, XML_PARSE_NOBLANKS)
	if reader = null then
		return false
	end if
	
	var outf = freefile
	if open(outFile for output as #outf) <> 0 then
		return false
	end if
	
	print #outf, "<?xml version=""1.0"" encoding=""UTF-8""?>"
	
	var utf16leStr = cast(ushort ptr, allocate((1024+1) * len(ushort)))
	var g = 0
	var isOpen = false
	
	do while xmlTextReaderRead(reader) = 1 		
		select case xmlTextReaderNodeType(reader)
		case XML_READER_TYPE_ELEMENT
			var tag = cast(const zstring ptr, xmlTextReaderConstName(reader))
			if *tag = "g" then
				if g = 0 then
					var x = cast(const zstring ptr, xmlTextReaderGetAttribute(reader, cast(xmlChar ptr, @"x")))
					if x <> null then
						print #outf, " x=""" + *x + """";
					end if
					var y = cast(const zstring ptr, xmlTextReaderGetAttribute(reader, cast(xmlChar ptr, @"y")))
					if y <> null then
						print #outf, " y=""" + *y + """";
					end if
					
					print #outf, ">";
					isOpen = false
				end if
				
				var code = cast(const zstring ptr, xmlTextReaderGetAttribute(reader, cast(xmlChar ptr, @"unicode")))
				select case as const len(*code)
				case 1
					utf16leStr[g] = code[0]
				case 6
					utf16leStr[g] = hex2ushort(code + 2)
				end select
				
				g += 1
			else
				var isEmpty = xmlTextReaderIsEmptyElement(reader)
				print #outf, "<" + *tag;
				
				do while xmlTextReaderMoveToNextAttribute(reader) = 1
					var attrib = cast(const zstring ptr, xmlTextReaderConstName(reader))
					var value = cast(const zstring ptr, xmlTextReaderConstValue(reader))
					print #outf, " " + *attrib + "=""" + *value + """";
				loop

				if *tag = "span" then
					isOpen = true
					g = 0
				else
					isOpen = false
					print #outf, iif(isEmpty, "/>", ">")
				end if
			end if
			
		case XML_READER_TYPE_END_ELEMENT
			if isOpen then
				isOpen = false
				print #outf, ">"
			end if
			
			var tag = cast(const zstring ptr, xmlTextReaderConstName(reader))
			if *tag = "span" then
				if g > 0 then
					var utf8str = utf16leToUtf8(utf16leStr, g * len(ushort))
					print #outf, *utf8str;
					deallocate utf8str
					g = 0
				end if
			end if
			print #outf, "</" + *tag + ">"
		
		case XML_READER_TYPE_TEXT
			if isOpen then
				isOpen = false
				print #outf, ">";
			end if
			var text = cast(const zstring ptr, xmlTextReaderConstValue(reader))
			if text <> null then
				print #outf, *text;
			end if
		end select
	loop
	
	return true
	
end function

function PdfTemplate.getVersion() as integer
	return version
end function

function PdfTemplate.getPage(index as integer) as PdfTemplatePageNode ptr
	var page = root->getHead()
	var cnt = 0
	do while page <> null
		if cnt = index then
			return cast(PdfTemplatePageNode ptr, page)
		end if
		page = page->getNext()
		cnt += 1
	loop
	return null
end function

#endif 'WITH_PARSER

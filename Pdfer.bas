'' PDFium Helper Library for FreeBASIC
'' Copyright 2020 by Andre Victor (av1ctortv[@]gmail.com)

#include once "Pdfer.bi"
#include once "libiconv.bi"

	dim shared cd8to16le as iconv_t
	dim shared cd16leto8 as iconv_t

'''''
private sub initialize() constructor
	FPDF_InitLibrary()
	cd8to16le = iconv_open("UTF-16LE", "UTF-8")
	cd16leto8 = iconv_open("UTF-8", "UTF-16LE")
end sub

private sub shutdown() destructor
	iconv_close(cd16leto8)
	iconv_close(cd8to16le)
	FPDF_DestroyLibrary()
end sub

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

private sub highlight(left_ as single, bottom as single, right_ as single, top as single, page as FPDF_PAGE)
	var annot = FPDFPage_CreateAnnot(page, FPDF_ANNOT_HIGHLIGHT)
	FPDFAnnot_SetFlags(annot, FPDF_ANNOT_FLAG_PRINT or FPDF_ANNOT_FLAG_READONLY)
	dim rectf as FS_RECTF = (left_-1, top-1, right_+1, bottom+1)
	FPDFAnnot_SetRect(annot, @rectf)
	dim quad as FS_QUADPOINTSF = (left_-1, top+1, right_+1, top+1, left_-1, bottom-1, right_+1, bottom-1)
	FPDFAnnot_AppendAttachmentPoints(annot, @quad)
	FPDFAnnot_SetColor(annot, FPDFANNOT_COLORTYPE_Color, 255, 209, 0, 102)
	FPDFPage_CloseAnnot(annot)
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
	..highlight(rect->left, rect->bottom, rect->right, rect->top, page)
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
constructor PdfDoc(doc as FPDF_DOCUMENT)
	this.doc = doc
end constructor

constructor PdfDoc()
	constructor(FPDF_CreateNewDocument())
end constructor

constructor PdfDoc(path as string)
	constructor(FPDF_LoadDocument(path, null))
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

'''''
constructor PdfElement()
end constructor

constructor PdfElement(parent as PdfElement ptr)
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

constructor PdfElement(id as zstring ptr, idDict as TDict ptr, parent as PdfElement ptr)
	constructor(parent)
	if id <> null andalso id[0] <> 0 then
		this.id = allocate(len(*id)+1)
		*this.id = *id
		if (*idDict)[id] = null then
			idDict->add(this.id, @this)
		end if
	end if
end constructor

destructor PdfElement()
	if id <> null then
		deallocate(id)
	end if
	
	var child = head
	do while child <> null
		var next_ = child->next_
		delete child
		child = next_
	loop
	head = null
end destructor

function PdfElement.clone(parent as PdfElement ptr, page as PdfPageElement_ ptr) as PdfElement ptr
	var dup = new PdfElement(id, page->getIdDict(), parent)
	cloneChildren(dup, page)
	return dup
end function

sub PdfElement.cloneChildren(parent as PdfElement ptr, page as PdfPageElement_ ptr)
	var child = this.head
	do while child <> null
		child->clone(parent, page)
		child = child->next_
	loop
end sub

function PdfElement.lookupAttrib(name_ as zstring ptr, byref type_ as PdfElementAttribType) as any ptr
	if *name_ = "hidden" then
		type_ = PdfElementAttribType.TP_BOOLEAN
		return @this.hidden
	end if
	return null
end function

sub PdfElement.setAttrib(name_ as zstring ptr, value as boolean)
	dim type_ as PdfElementAttribType
	var attrib = lookupAttrib(name_, type_)
	if attrib <> null then
		if type_ = PdfElementAttribType.TP_BOOLEAN then
			*cast(boolean ptr, attrib) = value
		end if
	end if
end sub

sub PdfElement.setAttrib(name_ as zstring ptr, value as integer)
	dim type_ as PdfElementAttribType
	var attrib = lookupAttrib(name_, type_)
	if attrib <> null then
		if type_ = PdfElementAttribType.TP_INTEGER then
			*cast(integer ptr, attrib) = value
		end if
	end if
end sub

sub PdfElement.setAttrib(name_ as zstring ptr, value as single)
	dim type_ as PdfElementAttribType
	var attrib = lookupAttrib(name_, type_)
	if attrib <> null then
		if type_ = PdfElementAttribType.TP_SINGLE then
			*cast(single ptr, attrib) = value
		end if
	end if
end sub

sub PdfElement.setAttrib(name_ as zstring ptr, value as double)
	dim type_ as PdfElementAttribType
	var attrib = lookupAttrib(name_, type_)
	if attrib <> null then
		if type_ = PdfElementAttribType.TP_DOUBLE then
			*cast(double ptr, attrib) = value
		end if
	end if
end sub

sub PdfElement.setAttrib(name_ as zstring ptr, value as wstring ptr)
	dim type_ as PdfElementAttribType
	var attrib = cast(any ptr ptr, lookupAttrib(name_, type_))
	if attrib <> null then
		if type_ = PdfElementAttribType.TP_WSTRINGPTR then
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

sub PdfElement.setAttrib(name_ as zstring ptr, value as zstring ptr)
	dim type_ as PdfElementAttribType
	var attrib = cast(any ptr ptr, lookupAttrib(name_, type_))
	if attrib <> null then
		if type_ = PdfElementAttribType.TP_WSTRINGPTR then
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

function PdfElement.getWidth() as single
	return getChildrenWidth()
end function

function PdfElement.getChildrenWidth() as single
	var maxW = 0.0
	var child = this.head
	do while child <> null
		var w = child->getWidth()
		if w > maxW then
			maxW = w
		end if
		child = child->next_
	loop
	return maxW
end function

function PdfElement.getHeight() as single
	return getChildrenHeight()
end function

function PdfElement.getChildrenHeight() as single
	var maxH = 0.0
	var child = this.head
	do while child <> null
		var h = child->getHeight()
		if h > maxH then
			maxH = h
		end if
		child = child->next_
	loop
	return maxH
end function

sub PdfElement.translate(xi as single, yi as single)
	translateChildren(xi, yi)
end sub

sub PdfElement.translateChildren(xi as single, yi as single)
	var child = this.head
	do while child <> null
		child->translate(xi, yi)
		child = child->next_
	loop
end sub

sub PdfElement.translateX(xi as single)
	translateXChildren(xi)
end sub

sub PdfElement.translateXChildren(xi as single)
	var child = this.head
	do while child <> null
		child->translateX(xi)
		child = child->next_
	loop
end sub

sub PdfElement.translateY(yi as single)
	translateYChildren(yi)
end sub

sub PdfElement.translateYChildren(yi as single)
	var child = this.head
	do while child <> null
		child->translateY(yi)
		child = child->next_
	loop
end sub


function PdfElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	return null
end function

sub PdfElement.renderInsert(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT)
	if hidden then
		return
	end if
		
	var obj = render(doc, page, parent)

	renderChildren(doc, page, obj)
	
	if obj <> null then
		page->insert(obj)
	end if
end sub

sub PdfElement.renderChildren(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT)
	var child = this.head
	do while child <> null
		child->renderInsert(doc, page, parent)
		child = child->next_
	loop
end sub

function PdfElement.getChild(id as zstring ptr) as PdfElement ptr 
	var child = this.head
	do while child <> null
		if *child->id = *id then
			return child
		end if
		child = child->next_
	loop

	if child = null then
		child = this.head
		do while child <> null
			if child->head <> null then
				var node = child->getChild(id)
				if node <> null then
					return node
				end if
			end if
			child = child->next_
		loop
	end if

	return null
end function

function PdfElement.getHead() as PdfElement ptr
	return this.head
end function

function PdfElement.getTail() as PdfElement ptr
	return this.tail
end function

function PdfElement.getNext() as PdfElement ptr
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
constructor PdfColorElement(fill as PdfRGB ptr, parent as PdfElement ptr)
	constructor(null, fill, FPDF_COLORSPACE_DEVICERGB, parent)
end constructor

constructor PdfColorElement(stroke as PdfRGB ptr, fill as PdfRGB ptr, colorspace as integer, parent as PdfElement ptr)
	base(parent)
	this.scolor = stroke
	this.fcolor = fill
	this.colorspace = colorspace
end constructor

function PdfColorElement.withStroke(r as ulong, g as ulong, b as ulong, a as ulong) as PdfColorElement ptr
	if this.scolor = null then
		this.scolor = new PdfRGB(r, g, b, a)
	else
		this.scolor->r = r
		this.scolor->g = g
		this.scolor->b = b
		this.scolor->a = a
	end if
	return @this 
end function

function PdfColorElement.withFill(r as ulong, g as ulong, b as ulong, a as ulong) as PdfColorElement ptr
	if this.fcolor = null then
		this.fcolor = new PdfRGB(r, g, b, a)
	else
		this.fcolor->r = r
		this.fcolor->g = g
		this.fcolor->b = b
		this.fcolor->a = a
	end if
	return @this 
end function

function PdfColorElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var scolor2 = iif(scolor <> null, scolor->clone(), null)
	var fcolor2 = iif(fcolor <> null, fcolor->clone(), null)
	var dup = new PdfColorElement(scolor2, fcolor2, colorspace, parent)
	cloneChildren(dup, page)
	return dup
end function

destructor PdfColorElement()
	if fcolor <> null then
		delete fcolor
	end if
	if scolor <> null then
		delete scolor
	end if
end destructor

sub PdfColorElement.renderChildren(doc as PdfDoc ptr, page as PdfPageElement_ ptr, parent as FPDF_PAGEOBJECT)
	var olds = doc->style.scolor
	if scolor <> null then
		doc->style.scolor = scolor
	end if
	var oldf = doc->style.fcolor
	if fcolor <> null then
		doc->style.fcolor = fcolor
	end if
	
	var child = this.head
	do while child <> null
		child->renderInsert(doc, page, parent)
		child = child->next_
	loop
	
	doc->style.fcolor = olds
	doc->style.scolor = oldf
end sub

'''''
constructor PdfTransformElement(transf as FS_MATRIX ptr, parent as PdfElement ptr)
	base(parent)
	this.transf = transf
end constructor

function PdfTransformElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var transf2 = iif(transf <> null, cloneTranform(transf), null)
	var dup = new PdfTransformElement(transf2, parent)
	cloneChildren(dup, page)
	return dup
end function

destructor PdfTransformElement()
	if transf <> null then
		delete transf
	end if
end destructor

sub PdfTransformElement.renderChildren(doc as PdfDoc ptr, page as PdfPageElement_ ptr, parent as FPDF_PAGEOBJECT)
	var old = doc->style.transf
	doc->style.transf = transf

	var child = this.head
	do while child <> null
		child->renderInsert(doc, page, parent)
		child = child->next_
	loop
	
	doc->style.transf = old
end sub

'''''
constructor PdfFontElement(name_ as zstring ptr, size as single, parent as PdfElement ptr)
	base(parent)
	this.name_ = *name_
	this.size = size
end constructor

function PdfFontElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var dup = new PdfFontElement(name_, size, parent)
	cloneChildren(dup, page)
	return dup
end function

destructor PdfFontElement()
	if font <> null then
		FPDFFont_Close(font)
	end if
end destructor

sub PdfFontElement.renderChildren(doc as PdfDoc ptr, page as PdfPageElement_ ptr, parent as FPDF_PAGEOBJECT)
	if font = null then
		font = FPDFText_LoadStandardFont(doc->getDoc(), name_)
	end if
	
	var oldf = doc->style.font
	if font <> null then
		doc->style.font = font
	end if
	var olds = doc->style.size
	if size > 0.0! then
		doc->style.size = size
	end if

	var child = this.head
	do while child <> null
		child->renderInsert(doc, page, parent)
		child = child->next_
	loop
	
	doc->style.size = olds
	doc->style.font = oldf
end sub

'''''
private function buildPath(doc as PdfDoc ptr, fillMode as integer, lineWidth as single, miterlin as single = 0.0, join as integer = 0, cap as integer = 0) as FPDF_PAGEOBJECT
	var path = FPDFPageObj_CreateNewPath(0, 0)
	
	FPDFPath_SetDrawMode(path, fillMode, iif(lineWidth > 0.0, 1, 0))
	
	if lineWidth > 0.0 then
		var col = doc->style.scolor
		if col <> null then
			FPDFPageObj_SetStrokeColor(path, col->r, col->g, col->b, col->a)
		end if
		FPDFPageObj_SetStrokeWidth(path, lineWidth)
		FPDFPageObj_SetLineCap(path, cap)
		FPDFPageObj_SetLineJoin(path, join)
	end if
	
	if fillMode <> FPDF_FILLMODE_NONE then
		var col = doc->style.fcolor
		if col <> null then
			FPDFPageObj_SetFillColor(path, col->r, col->g, col->b, col->a)
		end if
	end if
	
	var transf = doc->style.transf
	if transf <> null then
		FPDFPageObj_Transform(path, transf->a, transf->b, transf->c, transf->d, transf->e, transf->f)
	end if
	
	return path
end function

'''''
constructor PdfStrokeElement(width_ as single, miterlin as single, join as integer, cap as integer, parent as PdfElement ptr)
	base(parent)
	this.width_ = width_
	this.miterlin = miterlin
	this.join = join
	this.cap = cap
end constructor

function PdfStrokeElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var dup = new PdfStrokeElement(width_, miterlin, join, cap, parent)
	cloneChildren(dup, page)
	return dup
end function

destructor PdfStrokeElement()
end destructor

function PdfStrokeElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	return buildPath(doc, FPDF_FILLMODE_NONE, width_, miterlin, join, cap)
end function

'''''
constructor PdfFillElement(mode as integer, parent as PdfElement ptr)
	base(parent)
	this.mode = mode
end constructor

function PdfFillElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var dup = new PdfFillElement(mode, parent)
	cloneChildren(dup, page)
	return dup
end function

destructor PdfFillElement()
end destructor

function PdfFillElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	return buildPath(doc, mode, 0.0)
end function

'''''
constructor PdfMoveToElement(x as single, y as single, parent as PdfElement ptr)
	base(parent)
	this.x = x 
	this.y = y
end constructor

function PdfMoveToElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var dup = new PdfMoveToElement(x, y, parent)
	return dup
end function

sub PdfMoveToElement.translate(xi as single, yi as single)
	this.x += xi
	this.y += yi
end sub

sub PdfMoveToElement.translateX(xi as single)
	this.x += xi
end sub

sub PdfMoveToElement.translateY(yi as single)
	this.y += yi
end sub

function PdfMoveToElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	FPDFPath_MoveTo(parent, this.x, this.y)
	return null
end function

'''''
constructor PdfLineToElement(x as single, y as single, parent as PdfElement ptr)
	base(parent)
	this.x = x 
	this.y = y
end constructor

function PdfLineToElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var dup = new PdfLineToElement(x, y, parent)
	return dup
end function

sub PdfLineToElement.translate(xi as single, yi as single)
	this.x += xi
	this.y += yi
end sub

sub PdfLineToElement.translateX(xi as single)
	this.x += xi
end sub

sub PdfLineToElement.translateY(yi as single)
	this.y += yi
end sub

function PdfLineToElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	FPDFPath_LineTo(parent, this.x, this.y)
	return null
end function

'''''
constructor PdfBezierToElement(x1 as single, y1 as single, x2 as single, y2 as single, x3 as single, y3 as single, parent as PdfElement ptr)
	base(parent)
	this.x1 = x1 
	this.y1 = y1
	this.x2 = x2 
	this.y2 = y2
	this.x3 = x3 
	this.y3 = y3
end constructor

function PdfBezierToElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var dup = new PdfBezierToElement(x1, y1, x2, y2, x3, y3, parent)
	return dup
end function

sub PdfBezierToElement.translate(xi as single, yi as single)
	this.x1 += xi
	this.y1 += yi
	this.x2 += xi
	this.y2 += yi
	this.x3 += xi
	this.y3 += yi
end sub

sub PdfBezierToElement.translateX(xi as single)
	this.x1 += xi
	this.x2 += xi
	this.x3 += xi
end sub

sub PdfBezierToElement.translateY(yi as single)
	this.y1 += yi
	this.y2 += yi
	this.y3 += yi
end sub

function PdfBezierToElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	FPDFPath_BezierTo(parent, this.x1, this.y1, this.x2, this.y2, this.x3, this.y3)
	return null
end function

'''''
constructor PdfClosePathElement(parent as PdfElement ptr)
	base(parent)
end constructor

function PdfClosePathElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var dup = new PdfClosePathElement(parent)
	cloneChildren(dup, page)
	return dup
end function

function PdfClosePathElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	FPDFPath_Close(parent)
	return null
end function

'''''
constructor PdfRectElement(x as single, y as single, w as single, h as single, mode as integer, lineWidth as single, miterlin as single, join as integer, cap as integer, parent as PdfElement ptr)
	base(parent)
	this.mode = mode
	this.x = x 
	this.y = y
	this.w = w 
	this.h = h
	this.linew = lineWidth
	this.miterlin = miterlin
	this.join = join
	this.cap = cap
end constructor

function PdfRectElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var dup = new PdfRectElement(x, y, w, h, mode, linew, miterlin, join, cap, parent)
	cloneChildren(dup, page)
	return dup
end function

function PdfRectElement.getWidth() as single
	var cw = getChildrenWidth()
	return iif(cw > w, cw, w)
end function

function PdfRectElement.getHeight() as single
	var ch = getChildrenHeight()
	return iif(ch > h, ch, h)
end function

sub PdfRectElement.translate(xi as single, yi as single)
	this.x += xi
	this.y += yi
	translateChildren(xi, yi)
end sub

sub PdfRectElement.translateX(xi as single)
	this.x += xi
	translateXChildren(xi)
end sub

sub PdfRectElement.translateY(yi as single)
	this.y += yi
	translateYChildren(yi)
end sub

function PdfRectElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	var path = buildPath(doc, mode, linew, miterlin, join, cap)
	FPDFPath_MoveTo(path, x, y)
	FPDFPath_LineTo(path, x+w, y)
	FPDFPath_LineTo(path, x+w, y-h)
	FPDFPath_LineTo(path, x, y-h)
	FPDFPath_LineTo(path, x, y)
	FPDFPath_Close(path)
	return path
end function

/'sub PdfRectElement.renderInsert(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT)
	if hidden then
		return
	end if
		
	var obj = render(doc, page, parent)

	renderChildren(doc, page, obj)
	
	if obj <> null then
		page->insert(obj)
	end if
end sub'/

'''''
constructor PdfHighlightElement(left as single, bottom as single, right as single, top as single, parent as PdfElement ptr)
	base(parent)
	this.left = left
	this.bottom = bottom
	this.right = right
	this.top = top
end constructor

function PdfHighlightElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var dup = new PdfHighlightElement(left, bottom, right, top, parent)
	cloneChildren(dup, page)
	return dup
end function

function PdfHighlightElement.getWidth() as single
	return abs(right - left) + 1
end function

function PdfHighlightElement.getHeight() as single
	return abs(top - bottom) + 1
end function

sub PdfHighlightElement.translate(xi as single, yi as single)
	this.left += xi
	this.bottom += yi
	this.right += xi
	this.top += yi
	translateChildren(xi, yi)
end sub

sub PdfHighlightElement.translateX(xi as single)
	this.left += xi
	this.right += xi
	translateXChildren(xi)
end sub

sub PdfHighlightElement.translateY(yi as single)
	this.bottom += yi
	this.top += yi
	translateYChildren(yi)
end sub

function PdfHighlightElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	highlight(left, bottom, right, top, page->getPage())
	return null
end function

'''''
constructor PdfTextElement(id as zstring ptr, idDict as TDict ptr, mode as FPDF_TEXT_RENDERMODE, x as single, y as single, width_ as single, align as PdfTextAlignment, text as wstring ptr, parent as PdfElement ptr)
	base(id, idDict, parent)
	this.mode = mode
	this.x = x
	this.y = y
	this.width_ = width_
	this.align = align
	this.text = text
end constructor

constructor PdfTextElement(mode as FPDF_TEXT_RENDERMODE, x as single, y as single, width_ as single, align as PdfTextAlignment, text as wstring ptr, parent as PdfElement ptr)
	constructor(null, null, mode, x, y, width_, align, text, parent)
end constructor

constructor PdfTextElement(x as single, y as single, width_ as single, align as PdfTextAlignment, text as wstring ptr, parent as PdfElement ptr)
	constructor(null, null, FPDF_TEXTRENDERMODE_FILL, x, y, width_, align, text, parent)
end constructor

constructor PdfTextElement(x as single, y as single, text as wstring ptr, parent as PdfElement ptr)
	constructor(null, null, FPDF_TEXTRENDERMODE_FILL, x, y, 0.0, PdfTextAlignment.TA_LEFT, text, parent)
end constructor

function PdfTextElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	dim text2 as wstring ptr
	if text <> null then
		text2 = cast(wstring ptr, allocate((len(*text)+1) * len(wstring)))
		*text2 = *text
	end if
	var dup = new PdfTextElement(id, page->getIdDict(), mode, x, y, width_, align, text2, parent)
	return dup
end function

private sub calcDim(obj as FPDF_PAGEOBJECT, byref w as single, byref h as single)
	if obj = null then
		w = 0.0
		h = 0.0
		return
	end if
	
	dim left_ as single = any, bottom as single = any, right_ as single = any, top as single = any
	FPDFPageObj_GetBounds(obj, @left_, @bottom, @right_, @top)
	w = abs(right_ - left_) + 1
	h = abs(top - bottom) + 1
end sub

sub PdfTextElement.translate(xi as single, yi as single)
	x += xi
	y += yi
end sub

sub PdfTextElement.translateX(xi as single)
	x += xi
end sub

sub PdfTextElement.translateY(yi as single)
	y += yi
end sub

function PdfTextElement.lookupAttrib(name_ as zstring ptr, byref type_ as PdfElementAttribType) as any ptr
	if *name_ = "text" then
		type_ = PdfElementAttribType.TP_WSTRINGPTR
		return @text
	elseif *name_ = "x" then
		type_ = PdfElementAttribType.TP_SINGLE
		return @x
	elseif *name_ = "y" then
		type_ = PdfElementAttribType.TP_SINGLE
		return @y
	else
		return base.lookupAttrib(name_, type_)
	end if
end function

destructor PdfTextElement()
	if text <> null then
		deallocate text
	end if
end destructor

function PdfTextElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	if text = null orelse len(*text) = 0 then
		return null
	end if
	
	var obj = FPDFPageObj_CreateTextObj(doc->getDoc(), doc->style.font, doc->style.size)

	FPDFText_SetText(obj, text)

	if mode = FPDF_TEXTRENDERMODE_FILL orelse mode = FPDF_TEXTRENDERMODE_FILL_STROKE then
		var col = doc->style.fcolor
		if col <> null then
			FPDFPageObj_SetFillColor(obj, col->r, col->g, col->b, col->a)
		end if
	end if

	if mode = FPDF_TEXTRENDERMODE_STROKE orelse mode = FPDF_TEXTRENDERMODE_FILL_STROKE then
		var col = doc->style.scolor
		if col <> null then
			FPDFPageObj_SetStrokeColor(obj, col->r, col->g, col->b, col->a)
		end if
	end if
	
	FPDFTextObj_SetTextRenderMode(obj, mode)
	
	var xpos = x
	if align <> PdfTextAlignment.TA_LEFT then
		dim dx as single = any, dy as single = any
		calcDim(obj, dx, dy)
		if align = PdfTextAlignment.TA_CENTER then
			xpos += (width_ / 2) - (dx / 2)
		else
			xpos += width_ - dx
		end if
	end if
	FPDFPageObj_Transform(obj, 1, 0, 0, 1, xpos, y)
	
	return obj
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
constructor PdfGroupElement(bbox as PdfRectCoords ptr, isolated as boolean, knockout as boolean, blendMode as zstring ptr, alpha as single, parent as PdfElement ptr)
	base(parent)
	this.bbox = bbox
	this.isolated = isolated
	this.knockout = knockout
	this.blendMode = blendMode
	this.alpha = alpha
end constructor

function PdfGroupElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var bbox2 = iif(bbox <> null, bbox->clone(), null)
	var dup = new PdfGroupElement(bbox2, isolated, knockout, blendMode, alpha, parent)
	cloneChildren(dup, page)
	return dup
end function

destructor PdfGroupElement()
	if bbox <> null then
		delete bbox
	end if
end destructor

function PdfGroupElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	return null
end function

'''''
constructor PdfTemplateElement(id as zstring ptr, idDict as TDict ptr, parent as PdfElement ptr, hidden as boolean)
	base(id, idDict, parent)
	base.hidden = hidden
end constructor

function PdfTemplateElement.clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	var dup = new PdfTemplateElement(id, page->getIdDict(), parent, hidden)
	cloneChildren(dup, page)
	return dup
end function

function PdfTemplateElement.render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	return null
end function

'''''
constructor PdfPageElement(x1 as single, y1 as single, x2 as single, y2 as single, parent as PdfElement ptr)
	base(parent)
	this.x1 = x1
	this.y1 = y1
	this.x2 = x2
	this.y2 = y2
	idDict = new TDict(64)
end constructor

destructor PdfPageElement()
	if page <> null then
		FPDF_ClosePage(page)
	end if
	delete idDict
end destructor

function PdfPageElement.clone() as PdfPageElement ptr
	var dup = new PdfPageElement(x1, y1, x2, y2, null)
	cloneChildren(dup, dup)
	return dup
end function

sub PdfPageElement.render(doc as PdfDoc ptr, index as integer, flush_ as boolean)
	if hidden then
		return
	end if
	
	page = FPDFPage_New(doc->getDoc(), index, x2 - x1, y2 - y1)
	FPDFPage_SetMediaBox(page, x1, y1, x2, y2)
	FPDFPage_SetCropBox(page, x1, y1, x2, y2)
	
	renderChildren(doc, @this, null)
	
	if flush_ then
		flush(doc)
	end if
end sub

sub PdfPageElement.flush(doc as PdfDoc ptr)
	if page <> null then
		FPDFPage_GenerateContent(page)
		FPDF_ClosePage(page)
		page = null
	end if
end sub

sub PdfPageElement.insert(obj as FPDF_PAGEOBJECT)
	FPDFPage_InsertObject(page, obj)
	'NOTE: the obj will be automatically freed by FPDF_ClosePage()
end sub

function PdfPageElement.getPage() as FPDF_PAGE
	return page
end function

function PdfPageElement.getIdDict() as TDict ptr
	return idDict
end function

function PdfPageElement.getNode(id as zstring ptr) as PdfElement ptr
	return cast(PdfElement ptr, (*idDict)[id])
end function

#ifdef WITH_PARSER
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

function PdfTemplate.parseColorAttrib(name_ as zstring ptr) as PdfRGB ptr
	dim colorArr() as double
	
	var cnt = getXmlAttribAsDoubleArray(name_, colorArr())
	select case cnt
	case 3
		return new PdfRGB(colorArr(0) * 255, colorArr(1) * 255, colorArr(2) * 255, 255)
	case 4
		return new PdfRGB(colorArr(0) * 255, colorArr(1) * 255, colorArr(2) * 255, colorArr(3) * 255)
	end select
	
	return null
end function

function PdfTemplate.parseTransformAttrib(name_ as zstring ptr) as FS_MATRIX ptr
	dim transfArr() as double
	
	var transfCnt = getXmlAttribAsDoubleArray(name_, transfArr())
	
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

function PdfTemplate.parseFill(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfFillElement ptr
	
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
	var transf = parseTransformAttrib()
	
	if color_ <> null then
		parent = new PdfColorElement(null, color_, colorspace, parent)
	end if
	
	if transf <> null then
		parent = new PdfTransformElement(transf, parent)
	end if
	
	return new PdfFillElement(mode, parent)
	
end function

function PdfTemplate.parseStroke(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfStrokeElement ptr
	
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
	var transf = parseTransformAttrib()

	if color_ <> null then
		parent = new PdfColorElement(null, color_, colorspace, parent)
	end if
	
	if transf <> null then
		parent = new PdfTransformElement(transf, parent)
	end if
	
	return new PdfStrokeElement(width_, miterlin, join, cap, parent)
	
end function

function PdfTemplate.parseMoveTo(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfMoveToElement ptr
	
	var x = getXmlAttribAsDouble("x")
	var y = getXmlAttribAsDouble("y")
	
	return new PdfMoveToElement(x, y, parent)
	
end function

function PdfTemplate.parseLineTo(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfLineToElement ptr
	
	var x = getXmlAttribAsDouble("x")
	var y = getXmlAttribAsDouble("y")
	
	return new PdfLineToElement(x, y, parent)
	
end function

function PdfTemplate.parseBezierTo(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfBezierToElement ptr
	
	var x1 = getXmlAttribAsDouble("x1")
	var y1 = getXmlAttribAsDouble("y1")
	var x2 = getXmlAttribAsDouble("x2")
	var y2 = getXmlAttribAsDouble("y2")
	var x3 = getXmlAttribAsDouble("x3")
	var y3 = getXmlAttribAsDouble("y3")
	
	return new PdfBezierToElement(x1, y1, x2, y2, x3, y3, parent)
	
end function

function PdfTemplate.parseClosePath(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfClosePathElement ptr
	
	return new PdfClosePathElement(parent)
	
end function

function PdfTemplate.parseRect(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfRectElement ptr
	
	var x = getXmlAttribAsDouble("x")
	var y = getXmlAttribAsDouble("y")
	var w = getXmlAttribAsDouble("w")
	var h = getXmlAttribAsDouble("h")
	var mode = getXmlAttribAsLong("mode")
	var linew = getXmlAttribAsDouble("linewidth")
	var miterlin = getXmlAttribAsDouble("miterlimit")
	var join = getXmlAttribAsLong("linejoin")
	dim cap as integer = FPDF_LINECAP_BUTT
	var attrib = getXmlAttrib("linecap")
	if len(attrib) >= 1 then
		cap = valint(left(attrib, 1))
	end if
	
	return new PdfRectElement(x, y, w, h, mode, linew, miterlin, join, cap, parent)
	
end function

function PdfTemplate.parseFillRect(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfRectElement ptr
	
	var x = getXmlAttribAsDouble("x")
	var y = getXmlAttribAsDouble("y")
	var w = getXmlAttribAsDouble("w")
	var h = getXmlAttribAsDouble("h")
	
	return new PdfRectElement(x, y, w, h, FPDF_FILLMODE_ALTERNATE, 0, 0, 0, 0, parent)
	
end function

function PdfTemplate.parseStrokeRect(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfRectElement ptr
	
	var x = getXmlAttribAsDouble("x")
	var y = getXmlAttribAsDouble("y")
	var w = getXmlAttribAsDouble("w")
	var h = getXmlAttribAsDouble("h")
	var linew = getXmlAttribAsDouble("linewidth")
	var miterlin = getXmlAttribAsDouble("miterlimit")
	var join = getXmlAttribAsLong("linejoin")
	dim cap as integer = FPDF_LINECAP_BUTT
	var attrib = getXmlAttrib("linecap")
	if len(attrib) >= 1 then
		cap = valint(left(attrib, 1))
	end if
	
	return new PdfRectElement(x, y, w, h, FPDF_FILLMODE_NONE, linew, miterlin, join, cap, parent)
	
end function

function PdfTemplate.parseGroup(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfGroupElement ptr
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
	
	return new PdfGroupElement(bbox, isolated, knockout, blendMode, alpha, parent)
	
end function

function PdfTemplate.parseTemplate(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfTemplateElement ptr
	
	var id = getXmlAttrib("id")
	var attrib = getXmlAttrib("hidden")
	var hidden = true
	if len(attrib) > 0 then
		hidden = attrib = "true" orelse attrib = "1"
	end if
	
	return new PdfTemplateElement(id, page->getIdDict(), parent, hidden)
	
end function

function PdfTemplate.parseText(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfTextElement ptr
	
	var colorspace = parseColorspaceAttrib()
	var color_ = parseColorAttrib()
	if color_ <> null then
		var attrib = getXmlAttrib("alpha")
		if len(attrib) > 0 then
			color_->a = val(attrib) * 255
		end if
	end if
	var transf = parseTransformAttrib()
	
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
	
	if len(font) > 0 then
		parent = new PdfFontElement(font, size, parent)
	end if
		
	if color_ <> null then
		parent = new PdfColorElement(null, color_, colorspace, parent)
	end if
	
	if transf <> null then
		parent = new PdfTransformElement(transf, parent)
	end if
	
	return new PdfTextElement(id, page->getIdDict(), mode, x, y, width_, align, text, parent)
	
end function

function PdfTemplate.parseColor(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfColorElement ptr
	
	var colorspace = parseColorspaceAttrib()
	var stroke = parseColorAttrib("stroke")
	var fill = parseColorAttrib("fill")
	
	return new PdfColorElement(stroke, fill, colorspace, parent)
	
end function

function PdfTemplate.parseTransform(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfTransformElement ptr
	
	var transf = parseTransformAttrib("matrix")
	
	return new PdfTransformElement(transf, parent)
	
end function

function PdfTemplate.parseFont(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfFontElement ptr
	
	var font = getXmlAttrib("name")
	var size = getXmlAttribAsDouble("size")
	
	return new PdfFontElement(font, size, parent)
	
end function

function PdfTemplate.parseObject(tag as zstring ptr, parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	
	dim obj as PdfElement ptr
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
	case "color"
		obj = parseColor(parent, page)
	case "transform"
		obj = parseTransform(parent, page)
	case "font"
		obj = parseFont(parent, page)
	case "rect"
		obj = parseRect(parent, page)
	case "fill_rect"
		obj = parseFillRect(parent, page)
	case "stroke_rect"
		obj = parseStrokeRect(parent, page)
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

sub PdfTemplate.parsePage(parent as PdfElement ptr)

	dim mediaboxArr() as double
	var arrCnt = getXmlAttribAsDoubleArray("mediabox", mediaboxArr())
	var x1 = 0.0, y1 = 0.0, x2 = 0.0, y2 = 0.0
	if arrCnt = 4 then
		x1 = mediaboxArr(0)
		y1 = mediaboxArr(1)
		x2 = mediaboxArr(2)
		y2 = mediaboxArr(3)
	end if
	
	var page = new PdfPageElement(x1, y1, x2, y2, parent)

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

sub PdfTemplate.parseDocument(parent as PdfElement ptr)

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
				root = new PdfElement()
				parseDocument(root)
				return true
			end if
		end if
	loop
	return false
end function

sub PdfTemplate.renderTo(doc as PdfDoc ptr, flush_ as boolean)
	var page = root->getHead()
	do while page <> null
		cast(PdfPageElement ptr, page)->render(doc, index, flush_)
		index += 1
		page = page->getNext()
	loop
end sub

sub PdfTemplate.flush(doc as PdfDoc ptr)
	var page = root->getHead()
	do while page <> null
		cast(PdfPageElement ptr, page)->flush(doc)
		page = page->getNext()
	loop
end sub

function PdfTemplate.clonePage(index as integer) as PdfPageElement ptr
	var page = root->getHead()
	var cnt = 0
	do while page <> null
		if cnt = index then
			return cast(PdfPageElement ptr, page)->clone()
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

function PdfTemplate.getPage(index as integer) as PdfPageElement ptr
	var page = root->getHead()
	var cnt = 0
	do while page <> null
		if cnt = index then
			return cast(PdfPageElement ptr, page)
		end if
		page = page->getNext()
		cnt += 1
	loop
	return null
end function

#endif 'WITH_PARSER

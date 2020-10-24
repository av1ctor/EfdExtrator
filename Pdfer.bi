#include once "fpdf.bi"
#include once "List.bi"
#ifdef WITH_PARSER
#	include once "libxml/xmlreader.bi"
#	include once "Dict.bi"
#endif

type PdfObj
public:
	declare destructor()
	declare sub setBorder(r as ulong, g as ulong, b as ulong, a as ulong = 255, width_ as single = -1.0)
	declare sub setBackground(r as ulong, g as ulong, b as ulong, a as ulong = 255)
protected:
	obj as FPDF_PAGEOBJECT
end type

type PdfRect extends PdfObj
public:
	declare constructor(x as single, y as single, w as single, h as single)
end type

type PdfPath extends PdfObj
public:
	declare constructor(x as single, y as single)
	declare sub moveTo(x as single, y as single)
	declare sub lineTo(x as single, y as single)
	declare sub bezierTo(x1 as single, y1 as single, x2 as single, y2 as single, x3 as single, y3 as single)
	declare sub close()
end type

type PdfFinderResult
	index	as long
	count	as long
end type

enum PdfFinderDirection explicit
	DOWN
	UP
end enum

type PdfFinder
public:
	declare constructor(handle as FPDF_SCHHANDLE)
	declare destructor()
	declare function find(dir as PdfFinderDirection = PdfFinderDirection.DOWN) as PdfFinderResult
	
private:
	handle as FPDF_SCHHANDLE
end type

enum PdfFindFlags explicit
	DEFAULT 		= &h00000000
	MATCHCASE 		= &h00000001
	MATCHWHOLEWORD 	= &h00000002
	CONSECUTIVE 	= &h00000004
end enum

type PdfRectCoords
	left	as double
    top		as double
    right	as double
    bottom	as double
	declare constructor()
	declare constructor(left as double, top as double, right as double, bottom as double)
	declare function clone() as PdfRectCoords ptr
end type

type PdfRGB
public:
	declare constructor(r as ulong, g as ulong, b as ulong, a as ulong = 255)
	declare function clone() as PdfRGB ptr
	r as ulong
	g as ulong
	b as ulong
	a as ulong
end type

type PdfText
public:
	declare constructor(text as FPDF_TEXTPAGE)
	declare destructor()
	declare property length() as integer
	declare property value() as wstring ptr
	declare function find(what as wstring ptr, index as long, flags as PdfFindFlags = PdfFindFlags.DEFAULT) as PdfFinder ptr
	declare function getRect(index as long, count as long) as PdfRectCoords ptr
private:
	text as FPDF_TEXTPAGE
	len_ as integer
end type

type PdfPage
public:
	declare constructor(page as FPDF_PAGE)
	declare destructor()
	declare property text() as PdfText ptr
	declare sub highlight(text as PdfText ptr, byref res as PdfFinderResult)
	declare sub highlight(rect as PdfRectCoords ptr)
	declare function getHandle() as FPDF_PAGE
private:
	page as FPDF_PAGE
end type

type PdfFileWriter extends FPDF_FILEWRITE
	bf as integer
	curSize as ulong
	maxSize as ulong
	mask as string
	count as integer
end type

type PdfStyle
	scolor as PdfRGB ptr
	fcolor as PdfRGB ptr
	font as FPDF_FONT
	size as single = 10.0
	transf as FS_MATRIX ptr
end type

type PdfDoc
public:
	style as PdfStyle
	declare constructor()
	declare constructor(doc as FPDF_DOCUMENT)
	declare constructor(path as string)
	declare destructor()
	declare property count as integer
	declare property page(index as integer) as PdfPage ptr
	declare sub importPages(src as PdfDoc ptr, fromPage as integer, toPage as integer)
	declare sub importPages(src as PdfDoc ptr, range as string)
	declare function saveTo(path as string, version as integer = 17) as boolean
	declare function getDoc() as FPDF_DOCUMENT
	
private:
	doc as FPDF_DOCUMENT
	declare static function blockWriterCb(byval pThis as PdfFileWriter ptr, byval pData as const any ptr, byval size as culong) as long
end type

enum PdfElementAttribType explicit
	TP_BOOLEAN
	TP_INTEGER
	TP_SINGLE
	TP_DOUBLE
	TP_WSTRINGPTR
end enum

type PdfPageElement_ as PdfPageElement

type PdfElement extends object
public:
	declare constructor()
	declare constructor(parent as PdfElement ptr)
	declare constructor(id as zstring ptr, idDict as TDict ptr, parent as PdfElement ptr)
	declare virtual destructor()
	declare function getHead() as PdfElement ptr
	declare function getTail() as PdfElement ptr
	declare function getNext() as PdfElement ptr
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement_ ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement_ ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare virtual function getWidth() as single
	declare virtual function getHeight() as single
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
	declare function getChild(id as zstring ptr) as PdfElement ptr 
	declare sub setAttrib(name_ as zstring ptr, value as boolean)
	declare sub setAttrib(name_ as zstring ptr, value as integer)
	declare sub setAttrib(name_ as zstring ptr, value as single)
	declare sub setAttrib(name_ as zstring ptr, value as double)
	declare sub setAttrib(name_ as zstring ptr, value as zstring ptr)
	declare sub setAttrib(name_ as zstring ptr, value as wstring ptr)
protected:
	id as zstring ptr
	hidden as boolean
	parent as PdfElement ptr
	next_ as PdfElement ptr
	head as PdfElement ptr
	tail as PdfElement ptr

	declare sub cloneChildren(parent as PdfElement ptr, page as PdfPageElement_ ptr)
	declare virtual sub renderChildren(doc as PdfDoc ptr, page as PdfPageElement_ ptr, parent as FPDF_PAGEOBJECT)
	declare sub renderInsert(doc as PdfDoc ptr, page as PdfPageElement_ ptr, parent as FPDF_PAGEOBJECT)
	declare function getChildrenWidth() as single
	declare function getChildrenHeight() as single
	declare sub translateChildren(xi as single, yi as single)
	declare sub translateXChildren(xi as single)
	declare sub translateYChildren(yi as single)
	declare virtual function lookupAttrib(name_ as zstring ptr, byref type_ as PdfElementAttribType) as any ptr
end type

type PdfPageElement extends PdfElement
public:
	declare constructor(x1 as single, y1 as single, x2 as single, y2 as single, parent as PdfElement ptr)
	declare virtual destructor()
	declare sub render(doc as PdfDoc ptr, index as integer, flush_ as boolean = true)
	declare sub flush(doc as PdfDoc ptr)
	declare function clone() as PdfPageElement ptr
	declare function getNode(id as zstring ptr) as PdfElement ptr
	declare function getIdDict() as TDict ptr
	declare function getPage() as FPDF_PAGE
	declare sub insert(obj as FPDF_PAGEOBJECT)
private:
	x1 as single
	y1 as single
	x2 as single
	y2 as single
	idDict as TDict ptr
	page as FPDF_PAGE
end type

type PdfColorElement extends PdfElement
public:
	declare constructor (fill as PdfRGB ptr, parent as PdfElement ptr)
	declare constructor (stroke as PdfRGB ptr, fill as PdfRGB ptr, colorspace as integer, parent as PdfElement ptr)
	declare virtual destructor ()
	declare function withStroke(r as ulong, g as ulong, b as ulong, a as ulong) as PdfColorElement ptr
	declare function withFill(r as ulong, g as ulong, b as ulong, a as ulong) as PdfColorElement ptr
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
protected:
	declare virtual sub renderChildren(doc as PdfDoc ptr, page as PdfPageElement_ ptr, parent as FPDF_PAGEOBJECT)
private:
	fcolor as PdfRGB ptr
	scolor as PdfRGB ptr
	colorspace as integer
end type

type PdfTransformElement extends PdfElement
public:
	declare constructor(transf as FS_MATRIX ptr, parent as PdfElement ptr)
	declare virtual destructor()
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
protected:
	declare virtual sub renderChildren(doc as PdfDoc ptr, page as PdfPageElement_ ptr, parent as FPDF_PAGEOBJECT)
private:
	transf as FS_MATRIX ptr
end type

type PdfFontElement extends PdfElement
public:
	declare constructor(name_ as zstring ptr, size as single, parent as PdfElement ptr)
	declare virtual destructor()
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
protected:
	declare virtual sub renderChildren(doc as PdfDoc ptr, page as PdfPageElement_ ptr, parent as FPDF_PAGEOBJECT)
private:
	name_ as string
	size as single
	font as FPDF_FONT
end type

type PdfFillElement extends PdfElement
public:
	declare constructor(mode as integer, parent as PdfElement ptr)
	declare virtual destructor()
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
private:
	mode as integer
end type

type PdfStrokeElement extends PdfElement
public:
	declare constructor(width_ as single, miterlin as single, join as integer, cap as integer, parent as PdfElement ptr)
	declare virtual destructor()
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
private:
	width_ as single
	miterlin as single
	join as integer
	cap as integer
end type

type PdfMoveToElement extends PdfElement
public:
	declare constructor(x as single, y as single, parent as PdfElement ptr = null)
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
private:
	x as single
	y as single
end type

type PdfLineToElement extends PdfElement
public:
	declare constructor(x as single, y as single, parent as PdfElement ptr = null)
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
private:
	x as single
	y as single
end type

type PdfBezierToElement extends PdfElement
public:
	declare constructor(x1 as single, y1 as single, x2 as single, y2 as single, x3 as single, y3 as single, parent as PdfElement ptr = null)
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
private:
	x1 as single
	y1 as single
	x2 as single
	y2 as single
	x3 as single
	y3 as single
end type

type PdfClosePathElement extends PdfElement
public:
	declare constructor(parent as PdfElement ptr = null)
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
private:
end type

type PdfRectElement extends PdfElement
public:
	declare constructor(x as single, y as single, w as single, h as single, mode as integer, lineWidth as single, miterlin as single, join as integer, cap as integer, parent as PdfElement ptr)
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare virtual function getWidth() as single
	declare virtual function getHeight() as single
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
private:
	mode as integer
	x as single
	y as single
	w as single
	h as single
	linew as single
	miterlin as single
	join as integer
	cap as integer
end type

type PdfHighlightElement extends PdfElement
public:
	declare constructor(left as single, bottom as single, right as single, top as single, parent as PdfElement ptr = null)
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare virtual function getWidth() as single
	declare virtual function getHeight() as single
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
private:
	left as single
	bottom as single
	right as single
	top as single
end type

enum PdfTextAlignment explicit
	TA_LEFT
	TA_CENTER
	TA_RIGHT
end enum

type PdfTextElement extends PdfElement
public:
	declare constructor(id as zstring ptr, idDict as TDict ptr, mode as FPDF_TEXT_RENDERMODE, x as single, y as single, width_ as single, align as PdfTextAlignment, text as wstring ptr, parent as PdfElement ptr)
	declare constructor(mode as FPDF_TEXT_RENDERMODE, x as single, y as single, width_ as single, align as PdfTextAlignment, text as wstring ptr, parent as PdfElement ptr)
	declare constructor(x as single, y as single, width_ as single, align as PdfTextAlignment, text as wstring ptr, parent as PdfElement ptr)
	declare constructor(x as single, y as single, text as wstring ptr, parent as PdfElement ptr)
	declare virtual destructor()
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
protected:
	declare virtual function lookupAttrib(name_ as zstring ptr, byref type_ as PdfElementAttribType) as any ptr
private:
	mode as FPDF_TEXT_RENDERMODE
	x as single
	y as single
	width_ as single
	align as PdfTextAlignment
	text as wstring ptr
end type

type PdfGroupElement extends PdfElement
public:
	declare constructor(bbox as PdfRectCoords ptr, isolated as boolean, knockout as boolean, blendMode as zstring ptr, alpha as single, parent as PdfElement ptr)
	declare virtual destructor()
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
private:
	bbox as PdfRectCoords ptr
	isolated as boolean
	knockout as boolean
	blendMode as zstring ptr
	alpha as single
end type

type PdfTemplateElement extends PdfElement
public:
	declare constructor(id as zstring ptr, idDict as TDict ptr, parent as PdfElement ptr, hidden as boolean = true)
	declare virtual function clone(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare virtual function render(doc as PdfDoc ptr, page as PdfPageElement ptr, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
end type

#ifdef WITH_PARSER
type PdfTemplate
public:
	declare constructor(buff as zstring ptr, size as integer, encoding_ as zstring ptr = null)
	declare constructor(path as string)
	declare destructor()
	declare function load() as boolean
	declare sub renderTo(doc as PdfDoc ptr, flush_ as boolean = true)
	declare sub flush(doc as PdfDoc ptr)
	declare function clonePage(index as integer) as PdfPageElement ptr
	declare function getPage(index as integer) as PdfPageElement ptr
	declare static function simplifyXml(inFile as string, outFile as string) as boolean
	declare function getVersion() as integer
	
private:
	declare sub parseDocument(parent as PdfElement ptr)
	declare sub parsePage(parent as PdfElement ptr)
	declare function parseObject(tag as zstring ptr, parent as PdfElement ptr, page as PdfPageElement ptr) as PdfElement ptr
	declare function parseGroup(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfGroupElement ptr
	declare function parseTemplate(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfTemplateElement ptr
	declare function parseFill(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfFillElement ptr
	declare function parseStroke(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfStrokeElement ptr
	declare function parseMoveTo(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfMoveToElement ptr
	declare function parseLineTo(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfLineToElement ptr
	declare function parseBezierTo(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfBezierToElement ptr
	declare function parseClosePath(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfClosePathElement ptr
	declare function parseRect(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfRectElement ptr
	declare function parseFillRect(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfRectElement ptr
	declare function parseStrokeRect(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfRectElement ptr
	declare function parseText(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfTextElement ptr
	declare function parseColor(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfColorElement ptr
	declare function parseTransform(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfTransformElement ptr
	declare function parseFont(parent as PdfElement ptr, page as PdfPageElement ptr) as PdfFontElement ptr
	declare function getXmlConstName() as string
	declare function getXmlAttrib(name_ as zstring ptr) as string
	declare function getXmlAttribAsLong(name_ as zstring ptr) as longint
	declare function getXmlAttribAsDouble(name_ as zstring ptr) as double
	declare function getXmlAttribAsLongArray(name_ as zstring ptr, toArr() as longint, delim as string = " ") as integer
	declare function getXmlAttribAsDoubleArray(name_ as zstring ptr, toArr() as double, delim as string = " ") as integer
	declare function parseColorAttrib(name_ as zstring ptr = @"color") as PdfRGB ptr
	declare function parseTransformAttrib(name_ as zstring ptr = @"transform") as FS_MATRIX ptr
	declare function parseColorspaceAttrib() as integer
	
	reader as xmlTextReaderPtr
	index as integer
	root as PdfElement ptr
	version as integer
end type
#endif 'WITH_PARSER

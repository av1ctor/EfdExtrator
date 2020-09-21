#include once "fpdf.bi"
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

type PdfDoc
public:
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

#ifdef WITH_PARSER
enum PdfTemplateNodeType explicit
	INVALID
	DOCUMENT
	PAGE
	GROUP
	TEMPLATE
	FILL
	STROKE
	MOVE_TO
	LINE_TO
	BEZIER_TO
	CLOSE_PATH
	TEXT
end enum

enum PdfTemplateAttribType explicit
	TP_BOOLEAN
	TP_INTEGER
	TP_SINGLE
	TP_DOUBLE
	TP_WSTRINGPTR
end enum

type PdfTemplatePageNode_ as PdfTemplatePageNode

type PdfTemplateNode extends object
public:
	declare constructor()
	declare constructor(type_ as PdfTemplateNodeType)
	declare constructor(type_ as PdfTemplateNodeType, parent as PdfTemplateNode ptr)
	declare constructor(type_ as PdfTemplateNodeType, id as string, idDict as TDict ptr, parent as PdfTemplateNode ptr)
	declare destructor()
	declare sub cloneChildren(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode_ ptr)
	declare function getHead() as PdfTemplateNode ptr
	declare function getTail() as PdfTemplateNode ptr
	declare function getNext() as PdfTemplateNode ptr
	declare virtual function clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode_ ptr) as PdfTemplateNode ptr
	declare virtual function emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare function emitAndInsert(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare sub emitChildren(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT)
	declare virtual function lookupAttrib(name_ as string, byref type_ as PdfTemplateAttribType) as any ptr
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
	declare function getChild(id as string) as PdfTemplateNode ptr 
	declare sub setAttrib(name_ as string, value as boolean)
	declare sub setAttrib(name_ as string, value as integer)
	declare sub setAttrib(name_ as string, value as single)
	declare sub setAttrib(name_ as string, value as double)
	declare sub setAttrib(name_ as string, value as zstring ptr)
	declare sub setAttrib(name_ as string, value as wstring ptr)
protected:
	type_ as PdfTemplateNodeType
	id as string
	hidden as boolean
	parent as PdfTemplateNode ptr
	next_ as PdfTemplateNode ptr
	head as PdfTemplateNode ptr
	tail as PdfTemplateNode ptr
private:
	obj as FPDF_PAGEOBJECT
end type

type PdfTemplatePageNode extends PdfTemplateNode
public:
	declare constructor(x1 as single, y1 as single, x2 as single, y2 as single, parent as PdfTemplateNode ptr)
	declare destructor()
	declare sub emit(doc as FPDF_DOCUMENT, index as integer, flush_ as boolean)
	declare sub emit(doc as PdfDoc ptr, index as integer, flush_ as boolean = true)
	declare sub flush()
	declare function clone() as PdfTemplatePageNode ptr
	declare function getIdDict() as TDict ptr
	declare function getNode(id as string) as PdfTemplateNode ptr
private:
	x1 as single
	y1 as single
	x2 as single
	y2 as single
	idDict as TDict
	page as FPDF_PAGE
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

type PdfTemplateFillNode extends PdfTemplateNode
public:
	declare constructor()
	declare constructor(mode as integer, colorspace as integer, color_ as PdfRGB ptr, transf as FS_MATRIX ptr, parent as PdfTemplateNode ptr)
	declare constructor(mode as integer, colorspace as integer, r as ulong, g as ulong, b as ulong, parent as PdfTemplateNode ptr)
	declare destructor()
	declare virtual function clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	declare virtual function emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
private:
	mode as integer
	colorspace as integer 
	color_ as PdfRGB ptr
	transf as FS_MATRIX ptr
end type

type PdfTemplateStrokeNode extends PdfTemplateNode
public:
	declare constructor()
	declare constructor(width_ as single, miterlin as single, join as integer, cap as integer, colorspace as integer, color_ as PdfRGB ptr, transf as FS_MATRIX ptr, parent as PdfTemplateNode ptr)
	declare destructor()
	declare virtual function clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	declare virtual function emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
private:
	width_ as single
	miterlin as single
	join as integer
	cap as integer
	colorspace as single 
	color_ as PdfRGB ptr
	transf as FS_MATRIX ptr
end type

type PdfTemplateMoveToNode extends PdfTemplateNode
public:
	declare constructor(x as single, y as single, parent as PdfTemplateNode ptr = null)
	declare virtual function clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	declare virtual function emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
private:
	x as single
	y as single
end type

type PdfTemplateLineToNode extends PdfTemplateNode
public:
	declare constructor(x as single, y as single, parent as PdfTemplateNode ptr = null)
	declare virtual function clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	declare virtual function emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
private:
	x as single
	y as single
end type

type PdfTemplateBezierToNode extends PdfTemplateNode
public:
	declare constructor(x1 as single, y1 as single, x2 as single, y2 as single, x3 as single, y3 as single, parent as PdfTemplateNode ptr = null)
	declare virtual function clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	declare virtual function emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
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

type PdfTemplateClosePathNode extends PdfTemplateNode
public:
	declare constructor(parent as PdfTemplateNode ptr = null)
	declare virtual function clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	declare virtual function emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
private:
end type

enum PdfTextAlignment explicit
	TA_LEFT
	TA_CENTER
	TA_RIGHT
end enum

type PdfTemplateTextNode extends PdfTemplateNode
public:
	declare constructor (id as string, idDict as TDict ptr, font as string, size as single, mode as FPDF_TEXT_RENDERMODE, x as single, y as single, width_ as single, align as PdfTextAlignment, text as wstring ptr, colorspace as integer, color_ as PdfRGB ptr, transf as FS_MATRIX ptr, parent as PdfTemplateNode ptr)
	declare destructor()
	declare virtual function clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	declare virtual function emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
	declare virtual sub translate(xi as single, yi as single)
	declare virtual sub translateX(xi as single)
	declare virtual sub translateY(yi as single)
	declare virtual function lookupAttrib(name_ as string, byref type_ as PdfTemplateAttribType) as any ptr
private:
	font as string
	size as single
	mode as FPDF_TEXT_RENDERMODE
	x as single
	y as single
	width_ as single
	align as PdfTextAlignment
	text as wstring ptr
	colorspace as integer
	color_ as PdfRGB ptr
	transf as FS_MATRIX ptr
end type

type PdfTemplateGroupNode extends PdfTemplateNode
public:
	declare constructor(bbox as PdfRectCoords ptr, isolated as boolean, knockout as boolean, blendMode as zstring ptr, alpha as single, parent as PdfTemplateNode ptr)
	declare destructor()
	declare virtual function clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	declare virtual function emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
private:
	bbox as PdfRectCoords ptr
	isolated as boolean
	knockout as boolean
	blendMode as zstring ptr
	alpha as single
end type

type PdfTemplateTemplateNode extends PdfTemplateNode
public:
	declare constructor(id as string, idDict as TDict ptr, parent as PdfTemplateNode ptr, hidden as boolean = true)
	declare virtual function clone(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	declare virtual function emit(doc as FPDF_DOCUMENT, page as FPDF_PAGE, parent as FPDF_PAGEOBJECT) as FPDF_PAGEOBJECT
end type

type PdfTemplate
public:
	declare constructor(buff as zstring ptr, size as integer, encoding_ as zstring ptr = null)
	declare constructor(path as string)
	declare destructor()
	declare function load() as boolean
	declare sub emitTo(doc as PdfDoc ptr, flush_ as boolean = true)
	declare sub flush()
	declare sub parseDocument(parent as PdfTemplateNode ptr)
	declare sub parsePage(parent as PdfTemplateNode ptr)
	declare function parseObject(tag as zstring ptr, parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateNode ptr
	declare function parseGroup(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateGroupNode ptr
	declare function parseTemplate(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateTemplateNode ptr
	declare function parseFill(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateFillNode ptr
	declare function parseStroke(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateStrokeNode ptr
	declare function parseMoveTo(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateMoveToNode ptr
	declare function parseLineTo(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateLineToNode ptr
	declare function parseBezierTo(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateBezierToNode ptr
	declare function parseClosePath(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateClosePathNode ptr
	declare function parseText(parent as PdfTemplateNode ptr, page as PdfTemplatePageNode ptr) as PdfTemplateTextNode ptr
	declare function clonePage(index as integer) as PdfTemplatePageNode ptr
	declare function getPage(index as integer) as PdfTemplatePageNode ptr
	declare static function simplifyXml(inFile as string, outFile as string) as boolean
	declare function getVersion() as integer
	
private:
	declare function getXmlConstName() as string
	declare function getXmlAttrib(name_ as zstring ptr) as string
	declare function getXmlAttribAsLong(name_ as zstring ptr) as longint
	declare function getXmlAttribAsDouble(name_ as zstring ptr) as double
	declare function getXmlAttribAsLongArray(name_ as zstring ptr, toArr() as longint, delim as string = " ") as integer
	declare function getXmlAttribAsDoubleArray(name_ as zstring ptr, toArr() as double, delim as string = " ") as integer
	declare function parseColorAttrib() as PdfRGB ptr
	declare function parseTranformAttrib() as FS_MATRIX ptr
	declare function parseColorspaceAttrib() as integer
	
	reader as xmlTextReaderPtr
	index as integer
	root as PdfTemplateNode ptr
	version as integer
end type
#endif 'WITH_PARSER

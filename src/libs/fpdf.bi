#pragma once

#include once "crt/long.bi"
#include once "crt/stddef.bi"
#include once "crt/stdint.bi"
#include once "crt/time.bi"

#inclib "pdfium"

extern "Windows-MS"

#define PUBLIC_FPDF_ANNOT_H_
#define PUBLIC_FPDFVIEW_H_
const FPDF_OBJECT_UNKNOWN = 0
const FPDF_OBJECT_BOOLEAN = 1
const FPDF_OBJECT_NUMBER = 2
const FPDF_OBJECT_STRING = 3
const FPDF_OBJECT_NAME = 4
const FPDF_OBJECT_ARRAY = 5
const FPDF_OBJECT_DICTIONARY = 6
const FPDF_OBJECT_STREAM = 7
const FPDF_OBJECT_NULLOBJ = 8
const FPDF_OBJECT_REFERENCE = 9

type FPDF_TEXT_RENDERMODE as long
enum
	FPDF_TEXTRENDERMODE_UNKNOWN = -1
	FPDF_TEXTRENDERMODE_FILL = 0
	FPDF_TEXTRENDERMODE_STROKE = 1
	FPDF_TEXTRENDERMODE_FILL_STROKE = 2
	FPDF_TEXTRENDERMODE_INVISIBLE = 3
	FPDF_TEXTRENDERMODE_FILL_CLIP = 4
	FPDF_TEXTRENDERMODE_STROKE_CLIP = 5
	FPDF_TEXTRENDERMODE_FILL_STROKE_CLIP = 6
	FPDF_TEXTRENDERMODE_CLIP = 7
	FPDF_TEXTRENDERMODE_LAST = FPDF_TEXTRENDERMODE_CLIP
end enum

type FPDF_ACTION as fpdf_action_t__ ptr
type FPDF_ANNOTATION as fpdf_annotation_t__ ptr
type FPDF_ATTACHMENT as fpdf_attachment_t__ ptr
type FPDF_BITMAP as fpdf_bitmap_t__ ptr
type FPDF_BOOKMARK as fpdf_bookmark_t__ ptr
type FPDF_CLIPPATH as fpdf_clippath_t__ ptr
type FPDF_DEST as fpdf_dest_t__ ptr
type FPDF_DOCUMENT as fpdf_document_t__ ptr
type FPDF_FONT as fpdf_font_t__ ptr
type FPDF_FORMHANDLE as fpdf_form_handle_t__ ptr
type FPDF_JAVASCRIPT_ACTION as fpdf_javascript_action_t ptr
type FPDF_LINK as fpdf_link_t__ ptr
type FPDF_PAGE as fpdf_page_t__ ptr
type FPDF_PAGELINK as fpdf_pagelink_t__ ptr
type FPDF_PAGEOBJECT as fpdf_pageobject_t__ ptr
type FPDF_PAGEOBJECTMARK as fpdf_pageobjectmark_t__ ptr
type FPDF_PAGERANGE as fpdf_pagerange_t__ ptr
type FPDF_PATHSEGMENT as const fpdf_pathsegment_t ptr
type FPDF_RECORDER as any ptr
type FPDF_SCHHANDLE as fpdf_schhandle_t__ ptr
type FPDF_SIGNATURE as fpdf_signature_t__ ptr
type FPDF_STRUCTELEMENT as fpdf_structelement_t__ ptr
type FPDF_STRUCTTREE as fpdf_structtree_t__ ptr
type FPDF_TEXTPAGE as fpdf_textpage_t__ ptr
type FPDF_WIDGET as fpdf_widget_t__ ptr
type FPDF_BOOL as long
type FPDF_RESULT as long
type FPDF_DWORD as culong
type FS_FLOAT as single

type _FPDF_DUPLEXTYPE_ as long
enum
	DuplexUndefined = 0
	Simplex
	DuplexFlipShortEdge
	DuplexFlipLongEdge
end enum

type FPDF_DUPLEXTYPE as _FPDF_DUPLEXTYPE_
type FPDF_WCHAR as ushort
type FPDF_BYTESTRING as const zstring ptr
#ifdef __FB_WIN32__
type FPDF_WIDESTRING as const wstring ptr
#else
type FPDF_WIDESTRING as const ushort ptr
#endif

type FPDF_BSTR_
	str as zstring ptr
	len as long
end type

type FPDF_BSTR as FPDF_BSTR_
type FPDF_STRING as const zstring ptr

type _FS_MATRIX_
	a as single
	b as single
	c as single
	d as single
	e as single
	f as single
end type

type FS_MATRIX as _FS_MATRIX_

type _FS_RECTF_
	left as single
	top as single
	right as single
	bottom as single
end type

type FS_LPRECTF as _FS_RECTF_ ptr
type FS_RECTF as _FS_RECTF_
type FS_LPCRECTF as const FS_RECTF ptr

type FS_SIZEF_
	width as single
	height as single
end type

type FS_LPSIZEF as FS_SIZEF_ ptr
type FS_SIZEF as FS_SIZEF_
type FS_LPCSIZEF as const FS_SIZEF ptr

type FS_POINTF_
	x as single
	y as single
end type

type FS_LPPOINTF as FS_POINTF_ ptr
type FS_POINTF as FS_POINTF_
type FS_LPCPOINTF as const FS_POINTF ptr
type FPDF_ANNOTATION_SUBTYPE as long
type FPDF_ANNOT_APPEARANCEMODE as long
type FPDF_OBJECT_TYPE as long

#define FPDF_CALLCONV
declare sub FPDF_InitLibrary()

type FPDF_LIBRARY_CONFIG_
	version as long
	m_pUserFontPaths as const zstring ptr ptr
	m_pIsolate as any ptr
	m_v8EmbedderSlot as ulong
	m_pPlatform as any ptr
end type

type FPDF_LIBRARY_CONFIG as FPDF_LIBRARY_CONFIG_
declare sub FPDF_InitLibraryWithConfig(byval config as const FPDF_LIBRARY_CONFIG ptr)
declare sub FPDF_DestroyLibrary()
const FPDF_POLICY_MACHINETIME_ACCESS = 0
declare sub FPDF_SetSandBoxPolicy(byval policy as FPDF_DWORD, byval enable as FPDF_BOOL)

#ifdef __FB_WIN32__
	declare function FPDF_SetPrintMode(byval mode as long) as FPDF_BOOL
#endif

declare function FPDF_LoadDocument(byval file_path as FPDF_STRING, byval password as FPDF_BYTESTRING) as FPDF_DOCUMENT
declare function FPDF_LoadMemDocument(byval data_buf as const any ptr, byval size as long, byval password as FPDF_BYTESTRING) as FPDF_DOCUMENT
declare function FPDF_LoadMemDocument64(byval data_buf as const any ptr, byval size as uinteger, byval password as FPDF_BYTESTRING) as FPDF_DOCUMENT

type FPDF_FILEACCESS
	m_FileLen as culong
	m_GetBlock as function(byval param as any ptr, byval position as culong, byval pBuf as ubyte ptr, byval size as culong) as long
	m_Param as any ptr
end type

type FPDF_FILEHANDLER_
	clientData as any ptr
	Release as sub(byval clientData as any ptr)
	GetSize as function(byval clientData as any ptr) as FPDF_DWORD
	ReadBlock as function(byval clientData as any ptr, byval offset as FPDF_DWORD, byval buffer as any ptr, byval size as FPDF_DWORD) as FPDF_RESULT
	WriteBlock as function(byval clientData as any ptr, byval offset as FPDF_DWORD, byval buffer as const any ptr, byval size as FPDF_DWORD) as FPDF_RESULT
	Flush as function(byval clientData as any ptr) as FPDF_RESULT
	Truncate as function(byval clientData as any ptr, byval size as FPDF_DWORD) as FPDF_RESULT
end type

type FPDF_FILEHANDLER as FPDF_FILEHANDLER_
declare function FPDF_LoadCustomDocument(byval pFileAccess as FPDF_FILEACCESS ptr, byval password as FPDF_BYTESTRING) as FPDF_DOCUMENT
declare function FPDF_GetFileVersion(byval doc as FPDF_DOCUMENT, byval fileVersion as long ptr) as FPDF_BOOL
const FPDF_ERR_SUCCESS = 0
const FPDF_ERR_UNKNOWN = 1
const FPDF_ERR_FILE = 2
const FPDF_ERR_FORMAT = 3
const FPDF_ERR_PASSWORD = 4
const FPDF_ERR_SECURITY = 5
const FPDF_ERR_PAGE = 6

declare function FPDF_GetLastError() as culong
declare function FPDF_DocumentHasValidCrossReferenceTable(byval document as FPDF_DOCUMENT) as FPDF_BOOL
declare function FPDF_GetDocPermissions(byval document as FPDF_DOCUMENT) as culong
declare function FPDF_GetSecurityHandlerRevision(byval document as FPDF_DOCUMENT) as long
declare function FPDF_GetPageCount(byval document as FPDF_DOCUMENT) as long
declare function FPDF_LoadPage(byval document as FPDF_DOCUMENT, byval page_index as long) as FPDF_PAGE
declare function FPDF_GetPageWidthF(byval page as FPDF_PAGE) as single
declare function FPDF_GetPageWidth(byval page as FPDF_PAGE) as double
declare function FPDF_GetPageHeightF(byval page as FPDF_PAGE) as single
declare function FPDF_GetPageHeight(byval page as FPDF_PAGE) as double
declare function FPDF_GetPageBoundingBox(byval page as FPDF_PAGE, byval rect as FS_RECTF ptr) as FPDF_BOOL
declare function FPDF_GetPageSizeByIndexF(byval document as FPDF_DOCUMENT, byval page_index as long, byval size as FS_SIZEF ptr) as FPDF_BOOL
declare function FPDF_GetPageSizeByIndex(byval document as FPDF_DOCUMENT, byval page_index as long, byval width as double ptr, byval height as double ptr) as long

const FPDF_ANNOT = &h01
const FPDF_LCD_TEXT = &h02
const FPDF_NO_NATIVETEXT = &h04
const FPDF_GRAYSCALE = &h08
const FPDF_DEBUG_INFO = &h80
const FPDF_NO_CATCH = &h100
const FPDF_RENDER_LIMITEDIMAGECACHE = &h200
const FPDF_RENDER_FORCEHALFTONE = &h400
const FPDF_PRINTING = &h800
const FPDF_RENDER_NO_SMOOTHTEXT = &h1000
const FPDF_RENDER_NO_SMOOTHIMAGE = &h2000
const FPDF_RENDER_NO_SMOOTHPATH = &h4000
const FPDF_REVERSE_BYTE_ORDER = &h10
const FPDF_CONVERT_FILL_TO_STROKE = &h20

type FPDF_COLORSCHEME_
	path_fill_color as FPDF_DWORD
	path_stroke_color as FPDF_DWORD
	text_fill_color as FPDF_DWORD
	text_stroke_color as FPDF_DWORD
end type

type FPDF_COLORSCHEME as FPDF_COLORSCHEME_

'#ifdef __FB_WIN32__
	'declare sub FPDF_RenderPage(byval dc as HDC, byval page as FPDF_PAGE, byval start_x as long, byval start_y as long, byval size_x as long, byval size_y as long, byval rotate as long, byval flags as long)
'#endif

declare sub FPDF_RenderPageBitmap(byval bitmap as FPDF_BITMAP, byval page as FPDF_PAGE, byval start_x as long, byval start_y as long, byval size_x as long, byval size_y as long, byval rotate as long, byval flags as long)
declare sub FPDF_RenderPageBitmapWithMatrix(byval bitmap as FPDF_BITMAP, byval page as FPDF_PAGE, byval matrix as const FS_MATRIX ptr, byval clipping as const FS_RECTF ptr, byval flags as long)
declare sub FPDF_ClosePage(byval page as FPDF_PAGE)
declare sub FPDF_CloseDocument(byval document as FPDF_DOCUMENT)
declare function FPDF_DeviceToPage(byval page as FPDF_PAGE, byval start_x as long, byval start_y as long, byval size_x as long, byval size_y as long, byval rotate as long, byval device_x as long, byval device_y as long, byval page_x as double ptr, byval page_y as double ptr) as FPDF_BOOL
declare function FPDF_PageToDevice(byval page as FPDF_PAGE, byval start_x as long, byval start_y as long, byval size_x as long, byval size_y as long, byval rotate as long, byval page_x as double, byval page_y as double, byval device_x as long ptr, byval device_y as long ptr) as FPDF_BOOL
declare function FPDFBitmap_Create(byval width as long, byval height as long, byval alpha as long) as FPDF_BITMAP

const FPDFBitmap_Unknown = 0
const FPDFBitmap_Gray = 1
const FPDFBitmap_BGR = 2
const FPDFBitmap_BGRx = 3
const FPDFBitmap_BGRA = 4

declare function FPDFBitmap_CreateEx(byval width as long, byval height as long, byval format as long, byval first_scan as any ptr, byval stride as long) as FPDF_BITMAP
declare function FPDFBitmap_GetFormat(byval bitmap as FPDF_BITMAP) as long
declare sub FPDFBitmap_FillRect(byval bitmap as FPDF_BITMAP, byval left as long, byval top as long, byval width as long, byval height as long, byval color as FPDF_DWORD)
declare function FPDFBitmap_GetBuffer(byval bitmap as FPDF_BITMAP) as any ptr
declare function FPDFBitmap_GetWidth(byval bitmap as FPDF_BITMAP) as long
declare function FPDFBitmap_GetHeight(byval bitmap as FPDF_BITMAP) as long
declare function FPDFBitmap_GetStride(byval bitmap as FPDF_BITMAP) as long
declare sub FPDFBitmap_Destroy(byval bitmap as FPDF_BITMAP)
declare function FPDF_VIEWERREF_GetPrintScaling(byval document as FPDF_DOCUMENT) as FPDF_BOOL
declare function FPDF_VIEWERREF_GetNumCopies(byval document as FPDF_DOCUMENT) as long
declare function FPDF_VIEWERREF_GetPrintPageRange(byval document as FPDF_DOCUMENT) as FPDF_PAGERANGE
declare function FPDF_VIEWERREF_GetPrintPageRangeCount(byval pagerange as FPDF_PAGERANGE) as uinteger
declare function FPDF_VIEWERREF_GetPrintPageRangeElement(byval pagerange as FPDF_PAGERANGE, byval index as uinteger) as long
declare function FPDF_VIEWERREF_GetDuplex(byval document as FPDF_DOCUMENT) as FPDF_DUPLEXTYPE
declare function FPDF_VIEWERREF_GetName(byval document as FPDF_DOCUMENT, byval key as FPDF_BYTESTRING, byval buffer as zstring ptr, byval length as culong) as culong
declare function FPDF_CountNamedDests(byval document as FPDF_DOCUMENT) as FPDF_DWORD
declare function FPDF_GetNamedDestByName(byval document as FPDF_DOCUMENT, byval name as FPDF_BYTESTRING) as FPDF_DEST
declare function FPDF_GetNamedDest(byval document as FPDF_DOCUMENT, byval index as long, byval buffer as any ptr, byval buflen as clong ptr) as FPDF_DEST
declare function FPDF_GetXFAPacketCount(byval document as FPDF_DOCUMENT) as long
declare function FPDF_GetXFAPacketName(byval document as FPDF_DOCUMENT, byval index as long, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDF_GetXFAPacketContent(byval document as FPDF_DOCUMENT, byval index as long, byval buffer as any ptr, byval buflen as culong, byval out_buflen as culong ptr) as FPDF_BOOL

#define PUBLIC_FPDF_DOC_H_
const PDFACTION_UNSUPPORTED = 0
const PDFACTION_GOTO = 1
const PDFACTION_REMOTEGOTO = 2
const PDFACTION_URI = 3
const PDFACTION_LAUNCH = 4
const PDFACTION_EMBEDDEDGOTO = 5
const PDFDEST_VIEW_UNKNOWN_MODE = 0
const PDFDEST_VIEW_XYZ = 1
const PDFDEST_VIEW_FIT = 2
const PDFDEST_VIEW_FITH = 3
const PDFDEST_VIEW_FITV = 4
const PDFDEST_VIEW_FITR = 5
const PDFDEST_VIEW_FITB = 6
const PDFDEST_VIEW_FITBH = 7
const PDFDEST_VIEW_FITBV = 8

type FPDF_FILEIDTYPE as long
enum
	FILEIDTYPE_PERMANENT = 0
	FILEIDTYPE_CHANGING = 1
end enum

type _FS_QUADPOINTSF
	x1 as FS_FLOAT
	y1 as FS_FLOAT
	x2 as FS_FLOAT
	y2 as FS_FLOAT
	x3 as FS_FLOAT
	y3 as FS_FLOAT
	x4 as FS_FLOAT
	y4 as FS_FLOAT
end type

type FS_QUADPOINTSF as _FS_QUADPOINTSF
declare function FPDFBookmark_GetFirstChild(byval document as FPDF_DOCUMENT, byval bookmark as FPDF_BOOKMARK) as FPDF_BOOKMARK
declare function FPDFBookmark_GetNextSibling(byval document as FPDF_DOCUMENT, byval bookmark as FPDF_BOOKMARK) as FPDF_BOOKMARK
declare function FPDFBookmark_GetTitle(byval bookmark as FPDF_BOOKMARK, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDFBookmark_Find(byval document as FPDF_DOCUMENT, byval title as FPDF_WIDESTRING) as FPDF_BOOKMARK
declare function FPDFBookmark_GetDest(byval document as FPDF_DOCUMENT, byval bookmark as FPDF_BOOKMARK) as FPDF_DEST
declare function FPDFBookmark_GetAction(byval bookmark as FPDF_BOOKMARK) as FPDF_ACTION
declare function FPDFAction_GetType(byval action as FPDF_ACTION) as culong
declare function FPDFAction_GetDest(byval document as FPDF_DOCUMENT, byval action as FPDF_ACTION) as FPDF_DEST
declare function FPDFAction_GetFilePath(byval action as FPDF_ACTION, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDFAction_GetURIPath(byval document as FPDF_DOCUMENT, byval action as FPDF_ACTION, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDFDest_GetDestPageIndex(byval document as FPDF_DOCUMENT, byval dest as FPDF_DEST) as long
declare function FPDFDest_GetView(byval dest as FPDF_DEST, byval pNumParams as culong ptr, byval pParams as FS_FLOAT ptr) as culong
declare function FPDFDest_GetLocationInPage(byval dest as FPDF_DEST, byval hasXVal as FPDF_BOOL ptr, byval hasYVal as FPDF_BOOL ptr, byval hasZoomVal as FPDF_BOOL ptr, byval x as FS_FLOAT ptr, byval y as FS_FLOAT ptr, byval zoom as FS_FLOAT ptr) as FPDF_BOOL
declare function FPDFLink_GetLinkAtPoint(byval page as FPDF_PAGE, byval x as double, byval y as double) as FPDF_LINK
declare function FPDFLink_GetLinkZOrderAtPoint(byval page as FPDF_PAGE, byval x as double, byval y as double) as long
declare function FPDFLink_GetDest(byval document as FPDF_DOCUMENT, byval link as FPDF_LINK) as FPDF_DEST
declare function FPDFLink_GetAction(byval link as FPDF_LINK) as FPDF_ACTION
declare function FPDFLink_Enumerate(byval page as FPDF_PAGE, byval start_pos as long ptr, byval link_annot as FPDF_LINK ptr) as FPDF_BOOL
declare function FPDFLink_GetAnnot(byval page as FPDF_PAGE, byval link_annot as FPDF_LINK) as FPDF_ANNOTATION
declare function FPDFLink_GetAnnotRect(byval link_annot as FPDF_LINK, byval rect as FS_RECTF ptr) as FPDF_BOOL
declare function FPDFLink_CountQuadPoints(byval link_annot as FPDF_LINK) as long
declare function FPDFLink_GetQuadPoints(byval link_annot as FPDF_LINK, byval quad_index as long, byval quad_points as FS_QUADPOINTSF ptr) as FPDF_BOOL
declare function FPDF_GetPageAAction(byval page as FPDF_PAGE, byval aa_type as long) as FPDF_ACTION
declare function FPDF_GetFileIdentifier(byval document as FPDF_DOCUMENT, byval id_type as FPDF_FILEIDTYPE, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDF_GetMetaText(byval document as FPDF_DOCUMENT, byval tag as FPDF_BYTESTRING, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDF_GetPageLabel(byval document as FPDF_DOCUMENT, byval page_index as long, byval buffer as any ptr, byval buflen as culong) as culong

#define PUBLIC_FPDF_FORMFILL_H_
const FORMTYPE_NONE = 0
const FORMTYPE_ACRO_FORM = 1
const FORMTYPE_XFA_FULL = 2
const FORMTYPE_XFA_FOREGROUND = 3
const FORMTYPE_COUNT = 4
const JSPLATFORM_ALERT_BUTTON_OK = 0
const JSPLATFORM_ALERT_BUTTON_OKCANCEL = 1
const JSPLATFORM_ALERT_BUTTON_YESNO = 2
const JSPLATFORM_ALERT_BUTTON_YESNOCANCEL = 3
const JSPLATFORM_ALERT_BUTTON_DEFAULT = JSPLATFORM_ALERT_BUTTON_OK
const JSPLATFORM_ALERT_ICON_ERROR = 0
const JSPLATFORM_ALERT_ICON_WARNING = 1
const JSPLATFORM_ALERT_ICON_QUESTION = 2
const JSPLATFORM_ALERT_ICON_STATUS = 3
const JSPLATFORM_ALERT_ICON_ASTERISK = 4
const JSPLATFORM_ALERT_ICON_DEFAULT = JSPLATFORM_ALERT_ICON_ERROR
const JSPLATFORM_ALERT_RETURN_OK = 1
const JSPLATFORM_ALERT_RETURN_CANCEL = 2
const JSPLATFORM_ALERT_RETURN_NO = 3
const JSPLATFORM_ALERT_RETURN_YES = 4
const JSPLATFORM_BEEP_ERROR = 0
const JSPLATFORM_BEEP_WARNING = 1
const JSPLATFORM_BEEP_QUESTION = 2
const JSPLATFORM_BEEP_STATUS = 3
const JSPLATFORM_BEEP_DEFAULT = 4

type _IPDF_JsPlatform
	version as long
	app_alert as function(byval pThis as _IPDF_JsPlatform ptr, byval Msg as FPDF_WIDESTRING, byval Title as FPDF_WIDESTRING, byval Type as long, byval Icon as long) as long
	app_beep as sub(byval pThis as _IPDF_JsPlatform ptr, byval nType as long)
	app_response as function(byval pThis as _IPDF_JsPlatform ptr, byval Question as FPDF_WIDESTRING, byval Title as FPDF_WIDESTRING, byval Default as FPDF_WIDESTRING, byval cLabel as FPDF_WIDESTRING, byval bPassword as FPDF_BOOL, byval response as any ptr, byval length as long) as long
	Doc_getFilePath as function(byval pThis as _IPDF_JsPlatform ptr, byval filePath as any ptr, byval length as long) as long
	Doc_mail as sub(byval pThis as _IPDF_JsPlatform ptr, byval mailData as any ptr, byval length as long, byval bUI as FPDF_BOOL, byval To as FPDF_WIDESTRING, byval Subject as FPDF_WIDESTRING, byval CC as FPDF_WIDESTRING, byval BCC as FPDF_WIDESTRING, byval Msg as FPDF_WIDESTRING)
	Doc_print as sub(byval pThis as _IPDF_JsPlatform ptr, byval bUI as FPDF_BOOL, byval nStart as long, byval nEnd as long, byval bSilent as FPDF_BOOL, byval bShrinkToFit as FPDF_BOOL, byval bPrintAsImage as FPDF_BOOL, byval bReverse as FPDF_BOOL, byval bAnnotations as FPDF_BOOL)
	Doc_submitForm as sub(byval pThis as _IPDF_JsPlatform ptr, byval formData as any ptr, byval length as long, byval URL as FPDF_WIDESTRING)
	Doc_gotoPage as sub(byval pThis as _IPDF_JsPlatform ptr, byval nPageNum as long)
	Field_browse as function(byval pThis as _IPDF_JsPlatform ptr, byval filePath as any ptr, byval length as long) as long
	m_pFormfillinfo as any ptr
	m_isolate as any ptr
	m_v8EmbedderSlot as ulong
end type

type IPDF_JSPLATFORM as _IPDF_JsPlatform
const FXCT_ARROW = 0
const FXCT_NESW = 1
const FXCT_NWSE = 2
const FXCT_VBEAM = 3
const FXCT_HBEAM = 4
const FXCT_HAND = 5
type TimerCallback as sub(byval idEvent as long)

type _FPDF_SYSTEMTIME
	wYear as ushort
	wMonth as ushort
	wDayOfWeek as ushort
	wDay as ushort
	wHour as ushort
	wMinute as ushort
	wSecond as ushort
	wMilliseconds as ushort
end type

type FPDF_SYSTEMTIME as _FPDF_SYSTEMTIME

type _FPDF_FORMFILLINFO
	version as long
	Release as sub(byval pThis as _FPDF_FORMFILLINFO ptr)
	FFI_Invalidate as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval page as FPDF_PAGE, byval left as double, byval top as double, byval right as double, byval bottom as double)
	FFI_OutputSelectedRect as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval page as FPDF_PAGE, byval left as double, byval top as double, byval right as double, byval bottom as double)
	FFI_SetCursor as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval nCursorType as long)
	FFI_SetTimer as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval uElapse as long, byval lpTimerFunc as TimerCallback) as long
	FFI_KillTimer as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval nTimerID as long)
	FFI_GetLocalTime as function(byval pThis as _FPDF_FORMFILLINFO ptr) as FPDF_SYSTEMTIME
	FFI_OnChange as sub(byval pThis as _FPDF_FORMFILLINFO ptr)
	FFI_GetPage as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval document as FPDF_DOCUMENT, byval nPageIndex as long) as FPDF_PAGE
	FFI_GetCurrentPage as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval document as FPDF_DOCUMENT) as FPDF_PAGE
	FFI_GetRotation as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval page as FPDF_PAGE) as long
	FFI_ExecuteNamedAction as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval namedAction as FPDF_BYTESTRING)
	FFI_SetTextFieldFocus as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval value as FPDF_WIDESTRING, byval valueLen as FPDF_DWORD, byval is_focus as FPDF_BOOL)
	FFI_DoURIAction as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval bsURI as FPDF_BYTESTRING)
	FFI_DoGoToAction as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval nPageIndex as long, byval zoomMode as long, byval fPosArray as single ptr, byval sizeofArray as long)
	m_pJsPlatform as IPDF_JSPLATFORM ptr
	xfa_disabled as FPDF_BOOL
	FFI_DisplayCaret as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval page as FPDF_PAGE, byval bVisible as FPDF_BOOL, byval left as double, byval top as double, byval right as double, byval bottom as double)
	FFI_GetCurrentPageIndex as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval document as FPDF_DOCUMENT) as long
	FFI_SetCurrentPage as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval document as FPDF_DOCUMENT, byval iCurPage as long)
	FFI_GotoURL as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval document as FPDF_DOCUMENT, byval wsURL as FPDF_WIDESTRING)
	FFI_GetPageViewRect as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval page as FPDF_PAGE, byval left as double ptr, byval top as double ptr, byval right as double ptr, byval bottom as double ptr)
	FFI_PageEvent as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval page_count as long, byval event_type as FPDF_DWORD)
	FFI_PopupMenu as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval page as FPDF_PAGE, byval hWidget as FPDF_WIDGET, byval menuFlag as long, byval x as single, byval y as single) as FPDF_BOOL
	FFI_OpenFile as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval fileFlag as long, byval wsURL as FPDF_WIDESTRING, byval mode as const zstring ptr) as FPDF_FILEHANDLER ptr
	FFI_EmailTo as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval fileHandler as FPDF_FILEHANDLER ptr, byval pTo as FPDF_WIDESTRING, byval pSubject as FPDF_WIDESTRING, byval pCC as FPDF_WIDESTRING, byval pBcc as FPDF_WIDESTRING, byval pMsg as FPDF_WIDESTRING)
	FFI_UploadTo as sub(byval pThis as _FPDF_FORMFILLINFO ptr, byval fileHandler as FPDF_FILEHANDLER ptr, byval fileFlag as long, byval uploadTo as FPDF_WIDESTRING)
	FFI_GetPlatform as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval platform as any ptr, byval length as long) as long
	FFI_GetLanguage as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval language as any ptr, byval length as long) as long
	FFI_DownloadFromURL as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval URL as FPDF_WIDESTRING) as FPDF_FILEHANDLER ptr
	FFI_PostRequestURL as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval wsURL as FPDF_WIDESTRING, byval wsData as FPDF_WIDESTRING, byval wsContentType as FPDF_WIDESTRING, byval wsEncode as FPDF_WIDESTRING, byval wsHeader as FPDF_WIDESTRING, byval response as FPDF_BSTR ptr) as FPDF_BOOL
	FFI_PutRequestURL as function(byval pThis as _FPDF_FORMFILLINFO ptr, byval wsURL as FPDF_WIDESTRING, byval wsData as FPDF_WIDESTRING, byval wsEncode as FPDF_WIDESTRING) as FPDF_BOOL
	FFI_OnFocusChange as sub(byval param as _FPDF_FORMFILLINFO ptr, byval annot as FPDF_ANNOTATION, byval page_index as long)
	FFI_DoURIActionWithKeyboardModifier as sub(byval param as _FPDF_FORMFILLINFO ptr, byval uri as FPDF_BYTESTRING, byval modifiers as long)
end type

type FPDF_FORMFILLINFO as _FPDF_FORMFILLINFO
declare function FPDFDOC_InitFormFillEnvironment(byval document as FPDF_DOCUMENT, byval formInfo as FPDF_FORMFILLINFO ptr) as FPDF_FORMHANDLE
declare sub FPDFDOC_ExitFormFillEnvironment(byval hHandle as FPDF_FORMHANDLE)
declare sub FORM_OnAfterLoadPage(byval page as FPDF_PAGE, byval hHandle as FPDF_FORMHANDLE)
declare sub FORM_OnBeforeClosePage(byval page as FPDF_PAGE, byval hHandle as FPDF_FORMHANDLE)
declare sub FORM_DoDocumentJSAction(byval hHandle as FPDF_FORMHANDLE)
declare sub FORM_DoDocumentOpenAction(byval hHandle as FPDF_FORMHANDLE)

const FPDFDOC_AACTION_WC = &h10
const FPDFDOC_AACTION_WS = &h11
const FPDFDOC_AACTION_DS = &h12
const FPDFDOC_AACTION_WP = &h13
const FPDFDOC_AACTION_DP = &h14
declare sub FORM_DoDocumentAAction(byval hHandle as FPDF_FORMHANDLE, byval aaType as long)
const FPDFPAGE_AACTION_OPEN = 0
const FPDFPAGE_AACTION_CLOSE = 1

declare sub FORM_DoPageAAction(byval page as FPDF_PAGE, byval hHandle as FPDF_FORMHANDLE, byval aaType as long)
declare function FORM_OnMouseMove(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval modifier as long, byval page_x as double, byval page_y as double) as FPDF_BOOL
declare function FORM_OnMouseWheel(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval modifier as long, byval page_coord as const FS_POINTF ptr, byval delta_x as long, byval delta_y as long) as FPDF_BOOL
declare function FORM_OnFocus(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval modifier as long, byval page_x as double, byval page_y as double) as FPDF_BOOL
declare function FORM_OnLButtonDown(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval modifier as long, byval page_x as double, byval page_y as double) as FPDF_BOOL
declare function FORM_OnRButtonDown(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval modifier as long, byval page_x as double, byval page_y as double) as FPDF_BOOL
declare function FORM_OnLButtonUp(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval modifier as long, byval page_x as double, byval page_y as double) as FPDF_BOOL
declare function FORM_OnRButtonUp(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval modifier as long, byval page_x as double, byval page_y as double) as FPDF_BOOL
declare function FORM_OnLButtonDoubleClick(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval modifier as long, byval page_x as double, byval page_y as double) as FPDF_BOOL
declare function FORM_OnKeyDown(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval nKeyCode as long, byval modifier as long) as FPDF_BOOL
declare function FORM_OnKeyUp(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval nKeyCode as long, byval modifier as long) as FPDF_BOOL
declare function FORM_OnChar(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval nChar as long, byval modifier as long) as FPDF_BOOL
declare function FORM_GetFocusedText(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval buffer as any ptr, byval buflen as culong) as culong
declare function FORM_GetSelectedText(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval buffer as any ptr, byval buflen as culong) as culong
declare sub FORM_ReplaceSelection(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval wsText as FPDF_WIDESTRING)
declare function FORM_SelectAllText(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE) as FPDF_BOOL
declare function FORM_CanUndo(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE) as FPDF_BOOL
declare function FORM_CanRedo(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE) as FPDF_BOOL
declare function FORM_Undo(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE) as FPDF_BOOL
declare function FORM_Redo(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE) as FPDF_BOOL
declare function FORM_ForceToKillFocus(byval hHandle as FPDF_FORMHANDLE) as FPDF_BOOL
declare function FORM_GetFocusedAnnot(byval handle as FPDF_FORMHANDLE, byval page_index as long ptr, byval annot as FPDF_ANNOTATION ptr) as FPDF_BOOL
declare function FORM_SetFocusedAnnot(byval handle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION) as FPDF_BOOL

const FPDF_FORMFIELD_UNKNOWN = 0
const FPDF_FORMFIELD_PUSHBUTTON = 1
const FPDF_FORMFIELD_CHECKBOX = 2
const FPDF_FORMFIELD_RADIOBUTTON = 3
const FPDF_FORMFIELD_COMBOBOX = 4
const FPDF_FORMFIELD_LISTBOX = 5
const FPDF_FORMFIELD_TEXTFIELD = 6
const FPDF_FORMFIELD_SIGNATURE = 7
const FPDF_FORMFIELD_COUNT = 8

declare function FPDFPage_HasFormFieldAtPoint(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval page_x as double, byval page_y as double) as long
declare function FPDFPage_FormFieldZOrderAtPoint(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval page_x as double, byval page_y as double) as long
declare sub FPDF_SetFormFieldHighlightColor(byval hHandle as FPDF_FORMHANDLE, byval fieldType as long, byval color as culong)
declare sub FPDF_SetFormFieldHighlightAlpha(byval hHandle as FPDF_FORMHANDLE, byval alpha as ubyte)
declare sub FPDF_RemoveFormFieldHighlight(byval hHandle as FPDF_FORMHANDLE)
declare sub FPDF_FFLDraw(byval hHandle as FPDF_FORMHANDLE, byval bitmap as FPDF_BITMAP, byval page as FPDF_PAGE, byval start_x as long, byval start_y as long, byval size_x as long, byval size_y as long, byval rotate as long, byval flags as long)
declare function FPDF_GetFormType(byval document as FPDF_DOCUMENT) as long
declare function FORM_SetIndexSelected(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval index as long, byval selected as FPDF_BOOL) as FPDF_BOOL
declare function FORM_IsIndexSelected(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval index as long) as FPDF_BOOL
declare function FPDF_LoadXFA(byval document as FPDF_DOCUMENT) as FPDF_BOOL

const FPDF_ANNOT_UNKNOWN = 0
const FPDF_ANNOT_TEXT = 1
const FPDF_ANNOT_LINK = 2
const FPDF_ANNOT_FREETEXT = 3
const FPDF_ANNOT_LINE = 4
const FPDF_ANNOT_SQUARE = 5
const FPDF_ANNOT_CIRCLE = 6
const FPDF_ANNOT_POLYGON = 7
const FPDF_ANNOT_POLYLINE = 8
const FPDF_ANNOT_HIGHLIGHT = 9
const FPDF_ANNOT_UNDERLINE = 10
const FPDF_ANNOT_SQUIGGLY = 11
const FPDF_ANNOT_STRIKEOUT = 12
const FPDF_ANNOT_STAMP = 13
const FPDF_ANNOT_CARET = 14
const FPDF_ANNOT_INK = 15
const FPDF_ANNOT_POPUP = 16
const FPDF_ANNOT_FILEATTACHMENT = 17
const FPDF_ANNOT_SOUND = 18
const FPDF_ANNOT_MOVIE = 19
const FPDF_ANNOT_WIDGET = 20
const FPDF_ANNOT_SCREEN = 21
const FPDF_ANNOT_PRINTERMARK = 22
const FPDF_ANNOT_TRAPNET = 23
const FPDF_ANNOT_WATERMARK = 24
const FPDF_ANNOT_THREED = 25
const FPDF_ANNOT_RICHMEDIA = 26
const FPDF_ANNOT_XFAWIDGET = 27
const FPDF_ANNOT_FLAG_NONE = 0
const FPDF_ANNOT_FLAG_INVISIBLE = 1 shl 0
const FPDF_ANNOT_FLAG_HIDDEN = 1 shl 1
const FPDF_ANNOT_FLAG_PRINT = 1 shl 2
const FPDF_ANNOT_FLAG_NOZOOM = 1 shl 3
const FPDF_ANNOT_FLAG_NOROTATE = 1 shl 4
const FPDF_ANNOT_FLAG_NOVIEW = 1 shl 5
const FPDF_ANNOT_FLAG_READONLY = 1 shl 6
const FPDF_ANNOT_FLAG_LOCKED = 1 shl 7
const FPDF_ANNOT_FLAG_TOGGLENOVIEW = 1 shl 8
const FPDF_ANNOT_APPEARANCEMODE_NORMAL = 0
const FPDF_ANNOT_APPEARANCEMODE_ROLLOVER = 1
const FPDF_ANNOT_APPEARANCEMODE_DOWN = 2
const FPDF_ANNOT_APPEARANCEMODE_COUNT = 3
const FPDF_FORMFLAG_NONE = 0
const FPDF_FORMFLAG_READONLY = 1 shl 0
const FPDF_FORMFLAG_REQUIRED = 1 shl 1
const FPDF_FORMFLAG_NOEXPORT = 1 shl 2
const FPDF_FORMFLAG_TEXT_MULTILINE = 1 shl 12
const FPDF_FORMFLAG_TEXT_PASSWORD = 1 shl 13
const FPDF_FORMFLAG_CHOICE_COMBO = 1 shl 17
const FPDF_FORMFLAG_CHOICE_EDIT = 1 shl 18
const FPDF_FORMFLAG_CHOICE_MULTI_SELECT = 1 shl 21

type FPDFANNOT_COLORTYPE as long
enum
	FPDFANNOT_COLORTYPE_Color = 0
	FPDFANNOT_COLORTYPE_InteriorColor
end enum

declare function FPDFAnnot_IsSupportedSubtype(byval subtype as FPDF_ANNOTATION_SUBTYPE) as FPDF_BOOL
declare function FPDFPage_CreateAnnot(byval page as FPDF_PAGE, byval subtype as FPDF_ANNOTATION_SUBTYPE) as FPDF_ANNOTATION
declare function FPDFPage_GetAnnotCount(byval page as FPDF_PAGE) as long
declare function FPDFPage_GetAnnot(byval page as FPDF_PAGE, byval index as long) as FPDF_ANNOTATION
declare function FPDFPage_GetAnnotIndex(byval page as FPDF_PAGE, byval annot as FPDF_ANNOTATION) as long
declare sub FPDFPage_CloseAnnot(byval annot as FPDF_ANNOTATION)
declare function FPDFPage_RemoveAnnot(byval page as FPDF_PAGE, byval index as long) as FPDF_BOOL
declare function FPDFAnnot_GetSubtype(byval annot as FPDF_ANNOTATION) as FPDF_ANNOTATION_SUBTYPE
declare function FPDFAnnot_IsObjectSupportedSubtype(byval subtype as FPDF_ANNOTATION_SUBTYPE) as FPDF_BOOL
declare function FPDFAnnot_UpdateObject(byval annot as FPDF_ANNOTATION, byval obj as FPDF_PAGEOBJECT) as FPDF_BOOL
declare function FPDFAnnot_AddInkStroke(byval annot as FPDF_ANNOTATION, byval points as const FS_POINTF ptr, byval point_count as uinteger) as long
declare function FPDFAnnot_RemoveInkList(byval annot as FPDF_ANNOTATION) as FPDF_BOOL
declare function FPDFAnnot_AppendObject(byval annot as FPDF_ANNOTATION, byval obj as FPDF_PAGEOBJECT) as FPDF_BOOL
declare function FPDFAnnot_GetObjectCount(byval annot as FPDF_ANNOTATION) as long
declare function FPDFAnnot_GetObject(byval annot as FPDF_ANNOTATION, byval index as long) as FPDF_PAGEOBJECT
declare function FPDFAnnot_RemoveObject(byval annot as FPDF_ANNOTATION, byval index as long) as FPDF_BOOL
declare function FPDFAnnot_SetColor(byval annot as FPDF_ANNOTATION, byval type as FPDFANNOT_COLORTYPE, byval R as ulong, byval G as ulong, byval B as ulong, byval A as ulong) as FPDF_BOOL
declare function FPDFAnnot_GetColor(byval annot as FPDF_ANNOTATION, byval type as FPDFANNOT_COLORTYPE, byval R as ulong ptr, byval G as ulong ptr, byval B as ulong ptr, byval A as ulong ptr) as FPDF_BOOL
declare function FPDFAnnot_HasAttachmentPoints(byval annot as FPDF_ANNOTATION) as FPDF_BOOL
declare function FPDFAnnot_SetAttachmentPoints(byval annot as FPDF_ANNOTATION, byval quad_index as uinteger, byval quad_points as const FS_QUADPOINTSF ptr) as FPDF_BOOL
declare function FPDFAnnot_AppendAttachmentPoints(byval annot as FPDF_ANNOTATION, byval quad_points as const FS_QUADPOINTSF ptr) as FPDF_BOOL
declare function FPDFAnnot_CountAttachmentPoints(byval annot as FPDF_ANNOTATION) as uinteger
declare function FPDFAnnot_GetAttachmentPoints(byval annot as FPDF_ANNOTATION, byval quad_index as uinteger, byval quad_points as FS_QUADPOINTSF ptr) as FPDF_BOOL
declare function FPDFAnnot_SetRect(byval annot as FPDF_ANNOTATION, byval rect as const FS_RECTF ptr) as FPDF_BOOL
declare function FPDFAnnot_GetRect(byval annot as FPDF_ANNOTATION, byval rect as FS_RECTF ptr) as FPDF_BOOL
declare function FPDFAnnot_HasKey(byval annot as FPDF_ANNOTATION, byval key as FPDF_BYTESTRING) as FPDF_BOOL
declare function FPDFAnnot_GetValueType(byval annot as FPDF_ANNOTATION, byval key as FPDF_BYTESTRING) as FPDF_OBJECT_TYPE
declare function FPDFAnnot_SetStringValue(byval annot as FPDF_ANNOTATION, byval key as FPDF_BYTESTRING, byval value as FPDF_WIDESTRING) as FPDF_BOOL
declare function FPDFAnnot_GetStringValue(byval annot as FPDF_ANNOTATION, byval key as FPDF_BYTESTRING, byval buffer as FPDF_WCHAR ptr, byval buflen as culong) as culong
declare function FPDFAnnot_GetNumberValue(byval annot as FPDF_ANNOTATION, byval key as FPDF_BYTESTRING, byval value as single ptr) as FPDF_BOOL
declare function FPDFAnnot_SetAP(byval annot as FPDF_ANNOTATION, byval appearanceMode as FPDF_ANNOT_APPEARANCEMODE, byval value as FPDF_WIDESTRING) as FPDF_BOOL
declare function FPDFAnnot_GetAP(byval annot as FPDF_ANNOTATION, byval appearanceMode as FPDF_ANNOT_APPEARANCEMODE, byval buffer as FPDF_WCHAR ptr, byval buflen as culong) as culong
declare function FPDFAnnot_GetLinkedAnnot(byval annot as FPDF_ANNOTATION, byval key as FPDF_BYTESTRING) as FPDF_ANNOTATION
declare function FPDFAnnot_GetFlags(byval annot as FPDF_ANNOTATION) as long
declare function FPDFAnnot_SetFlags(byval annot as FPDF_ANNOTATION, byval flags as long) as FPDF_BOOL
declare function FPDFAnnot_GetFormFieldFlags(byval handle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION) as long
declare function FPDFAnnot_GetFormFieldAtPoint(byval hHandle as FPDF_FORMHANDLE, byval page as FPDF_PAGE, byval point as const FS_POINTF ptr) as FPDF_ANNOTATION
declare function FPDFAnnot_GetFormFieldName(byval hHandle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION, byval buffer as FPDF_WCHAR ptr, byval buflen as culong) as culong
declare function FPDFAnnot_GetFormFieldType(byval hHandle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION) as long
declare function FPDFAnnot_GetFormFieldValue(byval hHandle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION, byval buffer as FPDF_WCHAR ptr, byval buflen as culong) as culong
declare function FPDFAnnot_GetOptionCount(byval hHandle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION) as long
declare function FPDFAnnot_GetOptionLabel(byval hHandle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION, byval index as long, byval buffer as FPDF_WCHAR ptr, byval buflen as culong) as culong
declare function FPDFAnnot_IsOptionSelected(byval handle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION, byval index as long) as FPDF_BOOL
declare function FPDFAnnot_GetFontSize(byval hHandle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION, byval value as single ptr) as FPDF_BOOL
declare function FPDFAnnot_IsChecked(byval hHandle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION) as FPDF_BOOL
declare function FPDFAnnot_SetFocusableSubtypes(byval hHandle as FPDF_FORMHANDLE, byval subtypes as const FPDF_ANNOTATION_SUBTYPE ptr, byval count as uinteger) as FPDF_BOOL
declare function FPDFAnnot_GetFocusableSubtypesCount(byval hHandle as FPDF_FORMHANDLE) as long
declare function FPDFAnnot_GetFocusableSubtypes(byval hHandle as FPDF_FORMHANDLE, byval subtypes as FPDF_ANNOTATION_SUBTYPE ptr, byval count as uinteger) as FPDF_BOOL
declare function FPDFAnnot_GetLink(byval annot as FPDF_ANNOTATION) as FPDF_LINK
declare function FPDFAnnot_GetFormControlCount(byval hHandle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION) as long
declare function FPDFAnnot_GetFormControlIndex(byval hHandle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION) as long
declare function FPDFAnnot_GetFormFieldExportValue(byval hHandle as FPDF_FORMHANDLE, byval annot as FPDF_ANNOTATION, byval buffer as FPDF_WCHAR ptr, byval buflen as culong) as culong
#define PUBLIC_FPDF_ATTACHMENT_H_
declare function FPDFDoc_GetAttachmentCount(byval document as FPDF_DOCUMENT) as long
declare function FPDFDoc_AddAttachment(byval document as FPDF_DOCUMENT, byval name as FPDF_WIDESTRING) as FPDF_ATTACHMENT
declare function FPDFDoc_GetAttachment(byval document as FPDF_DOCUMENT, byval index as long) as FPDF_ATTACHMENT
declare function FPDFDoc_DeleteAttachment(byval document as FPDF_DOCUMENT, byval index as long) as FPDF_BOOL
declare function FPDFAttachment_GetName(byval attachment as FPDF_ATTACHMENT, byval buffer as FPDF_WCHAR ptr, byval buflen as culong) as culong
declare function FPDFAttachment_HasKey(byval attachment as FPDF_ATTACHMENT, byval key as FPDF_BYTESTRING) as FPDF_BOOL
declare function FPDFAttachment_GetValueType(byval attachment as FPDF_ATTACHMENT, byval key as FPDF_BYTESTRING) as FPDF_OBJECT_TYPE
declare function FPDFAttachment_SetStringValue(byval attachment as FPDF_ATTACHMENT, byval key as FPDF_BYTESTRING, byval value as FPDF_WIDESTRING) as FPDF_BOOL
declare function FPDFAttachment_GetStringValue(byval attachment as FPDF_ATTACHMENT, byval key as FPDF_BYTESTRING, byval buffer as FPDF_WCHAR ptr, byval buflen as culong) as culong
declare function FPDFAttachment_SetFile(byval attachment as FPDF_ATTACHMENT, byval document as FPDF_DOCUMENT, byval contents as const any ptr, byval len as culong) as FPDF_BOOL
declare function FPDFAttachment_GetFile(byval attachment as FPDF_ATTACHMENT, byval buffer as any ptr, byval buflen as culong, byval out_buflen as culong ptr) as FPDF_BOOL
#define PUBLIC_FPDF_CATALOG_H_
declare function FPDFCatalog_IsTagged(byval document as FPDF_DOCUMENT) as FPDF_BOOL

#define PUBLIC_FPDF_DATAAVAIL_H_
const PDF_LINEARIZATION_UNKNOWN = -1
const PDF_NOT_LINEARIZED = 0
const PDF_LINEARIZED = 1
const PDF_DATA_ERROR = -1
const PDF_DATA_NOTAVAIL = 0
const PDF_DATA_AVAIL = 1
const PDF_FORM_ERROR = -1
const PDF_FORM_NOTAVAIL = 0
const PDF_FORM_AVAIL = 1
const PDF_FORM_NOTEXIST = 2

type _FX_FILEAVAIL
	version as long
	IsDataAvail as function(byval pThis as _FX_FILEAVAIL ptr, byval offset as uinteger, byval size as uinteger) as FPDF_BOOL
end type

type FX_FILEAVAIL as _FX_FILEAVAIL
type FPDF_AVAIL as any ptr
declare function FPDFAvail_Create(byval file_avail as FX_FILEAVAIL ptr, byval file as FPDF_FILEACCESS ptr) as FPDF_AVAIL
declare sub FPDFAvail_Destroy(byval avail as FPDF_AVAIL)

type _FX_DOWNLOADHINTS
	version as long
	AddSegment as sub(byval pThis as _FX_DOWNLOADHINTS ptr, byval offset as uinteger, byval size as uinteger)
end type

type FX_DOWNLOADHINTS as _FX_DOWNLOADHINTS
declare function FPDFAvail_IsDocAvail(byval avail as FPDF_AVAIL, byval hints as FX_DOWNLOADHINTS ptr) as long
declare function FPDFAvail_GetDocument(byval avail as FPDF_AVAIL, byval password as FPDF_BYTESTRING) as FPDF_DOCUMENT
declare function FPDFAvail_GetFirstPageNum(byval doc as FPDF_DOCUMENT) as long
declare function FPDFAvail_IsPageAvail(byval avail as FPDF_AVAIL, byval page_index as long, byval hints as FX_DOWNLOADHINTS ptr) as long
declare function FPDFAvail_IsFormAvail(byval avail as FPDF_AVAIL, byval hints as FX_DOWNLOADHINTS ptr) as long
declare function FPDFAvail_IsLinearized(byval avail as FPDF_AVAIL) as long

#define PUBLIC_FPDF_EDIT_H_
#define FPDF_ARGB(a, r, g, b) culng(culng(culng(culng(culng(b) and &hff) or culng(culng(culng(g) and &hff) shl 8)) or culng(culng(culng(r) and &hff) shl 16)) or culng(culng(culng(a) and &hff) shl 24))
#define FPDF_GetBValue(argb) cubyte(argb)
#define FPDF_GetGValue(argb) cubyte(cushort(argb) shr 8)
#define FPDF_GetRValue(argb) cubyte((argb) shr 16)
#define FPDF_GetAValue(argb) cubyte((argb) shr 24)
const FPDF_COLORSPACE_UNKNOWN = 0
const FPDF_COLORSPACE_DEVICEGRAY = 1
const FPDF_COLORSPACE_DEVICERGB = 2
const FPDF_COLORSPACE_DEVICECMYK = 3
const FPDF_COLORSPACE_CALGRAY = 4
const FPDF_COLORSPACE_CALRGB = 5
const FPDF_COLORSPACE_LAB = 6
const FPDF_COLORSPACE_ICCBASED = 7
const FPDF_COLORSPACE_SEPARATION = 8
const FPDF_COLORSPACE_DEVICEN = 9
const FPDF_COLORSPACE_INDEXED = 10
const FPDF_COLORSPACE_PATTERN = 11
const FPDF_PAGEOBJ_UNKNOWN = 0
const FPDF_PAGEOBJ_TEXT = 1
const FPDF_PAGEOBJ_PATH = 2
const FPDF_PAGEOBJ_IMAGE = 3
const FPDF_PAGEOBJ_SHADING = 4
const FPDF_PAGEOBJ_FORM = 5
const FPDF_SEGMENT_UNKNOWN = -1
const FPDF_SEGMENT_LINETO = 0
const FPDF_SEGMENT_BEZIERTO = 1
const FPDF_SEGMENT_MOVETO = 2
const FPDF_FILLMODE_NONE = 0
const FPDF_FILLMODE_ALTERNATE = 1
const FPDF_FILLMODE_WINDING = 2
const FPDF_FONT_TYPE1 = 1
const FPDF_FONT_TRUETYPE = 2
const FPDF_LINECAP_BUTT = 0
const FPDF_LINECAP_ROUND = 1
const FPDF_LINECAP_PROJECTING_SQUARE = 2
const FPDF_LINEJOIN_MITER = 0
const FPDF_LINEJOIN_ROUND = 1
const FPDF_LINEJOIN_BEVEL = 2
const FPDF_PRINTMODE_EMF = 0
const FPDF_PRINTMODE_TEXTONLY = 1
const FPDF_PRINTMODE_POSTSCRIPT2 = 2
const FPDF_PRINTMODE_POSTSCRIPT3 = 3
const FPDF_PRINTMODE_POSTSCRIPT2_PASSTHROUGH = 4
const FPDF_PRINTMODE_POSTSCRIPT3_PASSTHROUGH = 5
const FPDF_PRINTMODE_EMF_IMAGE_MASKS = 6

type FPDF_IMAGEOBJ_METADATA
	width as ulong
	height as ulong
	horizontal_dpi as single
	vertical_dpi as single
	bits_per_pixel as ulong
	colorspace as long
	marked_content_id as long
end type

declare function FPDF_CreateNewDocument() as FPDF_DOCUMENT
declare function FPDFPage_New(byval document as FPDF_DOCUMENT, byval page_index as long, byval width as double, byval height as double) as FPDF_PAGE
declare sub FPDFPage_Delete(byval document as FPDF_DOCUMENT, byval page_index as long)
declare function FPDFPage_GetRotation(byval page as FPDF_PAGE) as long
declare sub FPDFPage_SetRotation(byval page as FPDF_PAGE, byval rotate as long)
declare sub FPDFPage_InsertObject(byval page as FPDF_PAGE, byval page_obj as FPDF_PAGEOBJECT)
declare function FPDFPage_RemoveObject(byval page as FPDF_PAGE, byval page_obj as FPDF_PAGEOBJECT) as FPDF_BOOL
declare function FPDFPage_CountObjects(byval page as FPDF_PAGE) as long
declare function FPDFPage_GetObject(byval page as FPDF_PAGE, byval index as long) as FPDF_PAGEOBJECT
declare function FPDFPage_HasTransparency(byval page as FPDF_PAGE) as FPDF_BOOL
declare function FPDFPage_GenerateContent(byval page as FPDF_PAGE) as FPDF_BOOL
declare sub FPDFPageObj_Destroy(byval page_obj as FPDF_PAGEOBJECT)
declare function FPDFPageObj_HasTransparency(byval page_object as FPDF_PAGEOBJECT) as FPDF_BOOL
declare function FPDFPageObj_GetType(byval page_object as FPDF_PAGEOBJECT) as long
declare sub FPDFPageObj_Transform(byval page_object as FPDF_PAGEOBJECT, byval a as double, byval b as double, byval c as double, byval d as double, byval e as double, byval f as double)
declare sub FPDFPage_TransformAnnots(byval page as FPDF_PAGE, byval a as double, byval b as double, byval c as double, byval d as double, byval e as double, byval f as double)
declare function FPDFPageObj_NewImageObj(byval document as FPDF_DOCUMENT) as FPDF_PAGEOBJECT
declare function FPDFPageObj_CountMarks(byval page_object as FPDF_PAGEOBJECT) as long
declare function FPDFPageObj_GetMark(byval page_object as FPDF_PAGEOBJECT, byval index as culong) as FPDF_PAGEOBJECTMARK
declare function FPDFPageObj_AddMark(byval page_object as FPDF_PAGEOBJECT, byval name as FPDF_BYTESTRING) as FPDF_PAGEOBJECTMARK
declare function FPDFPageObj_RemoveMark(byval page_object as FPDF_PAGEOBJECT, byval mark as FPDF_PAGEOBJECTMARK) as FPDF_BOOL
declare function FPDFPageObjMark_GetName(byval mark as FPDF_PAGEOBJECTMARK, byval buffer as any ptr, byval buflen as culong, byval out_buflen as culong ptr) as FPDF_BOOL
declare function FPDFPageObjMark_CountParams(byval mark as FPDF_PAGEOBJECTMARK) as long
declare function FPDFPageObjMark_GetParamKey(byval mark as FPDF_PAGEOBJECTMARK, byval index as culong, byval buffer as any ptr, byval buflen as culong, byval out_buflen as culong ptr) as FPDF_BOOL
declare function FPDFPageObjMark_GetParamValueType(byval mark as FPDF_PAGEOBJECTMARK, byval key as FPDF_BYTESTRING) as FPDF_OBJECT_TYPE
declare function FPDFPageObjMark_GetParamIntValue(byval mark as FPDF_PAGEOBJECTMARK, byval key as FPDF_BYTESTRING, byval out_value as long ptr) as FPDF_BOOL
declare function FPDFPageObjMark_GetParamStringValue(byval mark as FPDF_PAGEOBJECTMARK, byval key as FPDF_BYTESTRING, byval buffer as any ptr, byval buflen as culong, byval out_buflen as culong ptr) as FPDF_BOOL
declare function FPDFPageObjMark_GetParamBlobValue(byval mark as FPDF_PAGEOBJECTMARK, byval key as FPDF_BYTESTRING, byval buffer as any ptr, byval buflen as culong, byval out_buflen as culong ptr) as FPDF_BOOL
declare function FPDFPageObjMark_SetIntParam(byval document as FPDF_DOCUMENT, byval page_object as FPDF_PAGEOBJECT, byval mark as FPDF_PAGEOBJECTMARK, byval key as FPDF_BYTESTRING, byval value as long) as FPDF_BOOL
declare function FPDFPageObjMark_SetStringParam(byval document as FPDF_DOCUMENT, byval page_object as FPDF_PAGEOBJECT, byval mark as FPDF_PAGEOBJECTMARK, byval key as FPDF_BYTESTRING, byval value as FPDF_BYTESTRING) as FPDF_BOOL
declare function FPDFPageObjMark_SetBlobParam(byval document as FPDF_DOCUMENT, byval page_object as FPDF_PAGEOBJECT, byval mark as FPDF_PAGEOBJECTMARK, byval key as FPDF_BYTESTRING, byval value as any ptr, byval value_len as culong) as FPDF_BOOL
declare function FPDFPageObjMark_RemoveParam(byval page_object as FPDF_PAGEOBJECT, byval mark as FPDF_PAGEOBJECTMARK, byval key as FPDF_BYTESTRING) as FPDF_BOOL
declare function FPDFImageObj_LoadJpegFile(byval pages as FPDF_PAGE ptr, byval count as long, byval image_object as FPDF_PAGEOBJECT, byval file_access as FPDF_FILEACCESS ptr) as FPDF_BOOL
declare function FPDFImageObj_LoadJpegFileInline(byval pages as FPDF_PAGE ptr, byval count as long, byval image_object as FPDF_PAGEOBJECT, byval file_access as FPDF_FILEACCESS ptr) as FPDF_BOOL
declare function FPDFImageObj_GetMatrix(byval image_object as FPDF_PAGEOBJECT, byval a as double ptr, byval b as double ptr, byval c as double ptr, byval d as double ptr, byval e as double ptr, byval f as double ptr) as FPDF_BOOL
declare function FPDFImageObj_SetMatrix(byval image_object as FPDF_PAGEOBJECT, byval a as double, byval b as double, byval c as double, byval d as double, byval e as double, byval f as double) as FPDF_BOOL
declare function FPDFImageObj_SetBitmap(byval pages as FPDF_PAGE ptr, byval count as long, byval image_object as FPDF_PAGEOBJECT, byval bitmap as FPDF_BITMAP) as FPDF_BOOL
declare function FPDFImageObj_GetBitmap(byval image_object as FPDF_PAGEOBJECT) as FPDF_BITMAP
declare function FPDFImageObj_GetRenderedBitmap(byval document as FPDF_DOCUMENT, byval page as FPDF_PAGE, byval image_object as FPDF_PAGEOBJECT) as FPDF_BITMAP
declare function FPDFImageObj_GetImageDataDecoded(byval image_object as FPDF_PAGEOBJECT, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDFImageObj_GetImageDataRaw(byval image_object as FPDF_PAGEOBJECT, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDFImageObj_GetImageFilterCount(byval image_object as FPDF_PAGEOBJECT) as long
declare function FPDFImageObj_GetImageFilter(byval image_object as FPDF_PAGEOBJECT, byval index as long, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDFImageObj_GetImageMetadata(byval image_object as FPDF_PAGEOBJECT, byval page as FPDF_PAGE, byval metadata as FPDF_IMAGEOBJ_METADATA ptr) as FPDF_BOOL
declare function FPDFPageObj_CreateNewPath(byval x as single, byval y as single) as FPDF_PAGEOBJECT
declare function FPDFPageObj_CreateNewRect(byval x as single, byval y as single, byval w as single, byval h as single) as FPDF_PAGEOBJECT
declare function FPDFPageObj_GetBounds(byval page_object as FPDF_PAGEOBJECT, byval left as single ptr, byval bottom as single ptr, byval right as single ptr, byval top as single ptr) as FPDF_BOOL
declare sub FPDFPageObj_SetBlendMode(byval page_object as FPDF_PAGEOBJECT, byval blend_mode as FPDF_BYTESTRING)
declare function FPDFPageObj_SetStrokeColor(byval page_object as FPDF_PAGEOBJECT, byval R as ulong, byval G as ulong, byval B as ulong, byval A as ulong) as FPDF_BOOL
declare function FPDFPageObj_GetStrokeColor(byval page_object as FPDF_PAGEOBJECT, byval R as ulong ptr, byval G as ulong ptr, byval B as ulong ptr, byval A as ulong ptr) as FPDF_BOOL
declare function FPDFPageObj_SetStrokeWidth(byval page_object as FPDF_PAGEOBJECT, byval width as single) as FPDF_BOOL
declare function FPDFPageObj_GetStrokeWidth(byval page_object as FPDF_PAGEOBJECT, byval width as single ptr) as FPDF_BOOL
declare function FPDFPageObj_GetLineJoin(byval page_object as FPDF_PAGEOBJECT) as long
declare function FPDFPageObj_SetLineJoin(byval page_object as FPDF_PAGEOBJECT, byval line_join as long) as FPDF_BOOL
declare function FPDFPageObj_GetLineCap(byval page_object as FPDF_PAGEOBJECT) as long
declare function FPDFPageObj_SetLineCap(byval page_object as FPDF_PAGEOBJECT, byval line_cap as long) as FPDF_BOOL
declare function FPDFPageObj_SetFillColor(byval page_object as FPDF_PAGEOBJECT, byval R as ulong, byval G as ulong, byval B as ulong, byval A as ulong) as FPDF_BOOL
declare function FPDFPageObj_GetFillColor(byval page_object as FPDF_PAGEOBJECT, byval R as ulong ptr, byval G as ulong ptr, byval B as ulong ptr, byval A as ulong ptr) as FPDF_BOOL
declare function FPDFPath_CountSegments(byval path as FPDF_PAGEOBJECT) as long
declare function FPDFPath_GetPathSegment(byval path as FPDF_PAGEOBJECT, byval index as long) as FPDF_PATHSEGMENT
declare function FPDFPathSegment_GetPoint(byval segment as FPDF_PATHSEGMENT, byval x as single ptr, byval y as single ptr) as FPDF_BOOL
declare function FPDFPathSegment_GetType(byval segment as FPDF_PATHSEGMENT) as long
declare function FPDFPathSegment_GetClose(byval segment as FPDF_PATHSEGMENT) as FPDF_BOOL
declare function FPDFPath_MoveTo(byval path as FPDF_PAGEOBJECT, byval x as single, byval y as single) as FPDF_BOOL
declare function FPDFPath_LineTo(byval path as FPDF_PAGEOBJECT, byval x as single, byval y as single) as FPDF_BOOL
declare function FPDFPath_BezierTo(byval path as FPDF_PAGEOBJECT, byval x1 as single, byval y1 as single, byval x2 as single, byval y2 as single, byval x3 as single, byval y3 as single) as FPDF_BOOL
declare function FPDFPath_Close(byval path as FPDF_PAGEOBJECT) as FPDF_BOOL
declare function FPDFPath_SetDrawMode(byval path as FPDF_PAGEOBJECT, byval fillmode as long, byval stroke as FPDF_BOOL) as FPDF_BOOL
declare function FPDFPath_GetDrawMode(byval path as FPDF_PAGEOBJECT, byval fillmode as long ptr, byval stroke as FPDF_BOOL ptr) as FPDF_BOOL
declare function FPDFPath_GetMatrix(byval path as FPDF_PAGEOBJECT, byval matrix as FS_MATRIX ptr) as FPDF_BOOL
declare function FPDFPath_SetMatrix(byval path as FPDF_PAGEOBJECT, byval matrix as const FS_MATRIX ptr) as FPDF_BOOL
declare function FPDFPageObj_NewTextObj(byval document as FPDF_DOCUMENT, byval font as FPDF_BYTESTRING, byval font_size as single) as FPDF_PAGEOBJECT
declare function FPDFText_SetText(byval text_object as FPDF_PAGEOBJECT, byval text as FPDF_WIDESTRING) as FPDF_BOOL
declare function FPDFText_LoadFont(byval document as FPDF_DOCUMENT, byval data as const ubyte ptr, byval size as ulong, byval font_type as long, byval cid as FPDF_BOOL) as FPDF_FONT
declare function FPDFText_LoadStandardFont(byval document as FPDF_DOCUMENT, byval font as FPDF_BYTESTRING) as FPDF_FONT
declare function FPDFTextObj_GetMatrix(byval text as FPDF_PAGEOBJECT, byval matrix as FS_MATRIX ptr) as FPDF_BOOL
declare function FPDFTextObj_GetFontSize(byval text as FPDF_PAGEOBJECT) as single
declare sub FPDFFont_Close(byval font as FPDF_FONT)
declare function FPDFPageObj_CreateTextObj(byval document as FPDF_DOCUMENT, byval font as FPDF_FONT, byval font_size as single) as FPDF_PAGEOBJECT
declare function FPDFTextObj_GetTextRenderMode(byval text as FPDF_PAGEOBJECT) as FPDF_TEXT_RENDERMODE
declare function FPDFTextObj_SetTextRenderMode(byval text as FPDF_PAGEOBJECT, byval render_mode as FPDF_TEXT_RENDERMODE) as FPDF_BOOL
declare function FPDFTextObj_GetFontName(byval text as FPDF_PAGEOBJECT, byval buffer as any ptr, byval length as culong) as culong
declare function FPDFTextObj_GetText(byval text_object as FPDF_PAGEOBJECT, byval text_page as FPDF_TEXTPAGE, byval buffer as any ptr, byval length as culong) as culong
declare function FPDFFormObj_CountObjects(byval form_object as FPDF_PAGEOBJECT) as long
declare function FPDFFormObj_GetObject(byval form_object as FPDF_PAGEOBJECT, byval index as culong) as FPDF_PAGEOBJECT
declare function FPDFFormObj_GetMatrix(byval form_object as FPDF_PAGEOBJECT, byval matrix as FS_MATRIX ptr) as FPDF_BOOL

#define PUBLIC_FPDF_EXT_H_
const FPDF_UNSP_DOC_XFAFORM = 1
const FPDF_UNSP_DOC_PORTABLECOLLECTION = 2
const FPDF_UNSP_DOC_ATTACHMENT = 3
const FPDF_UNSP_DOC_SECURITY = 4
const FPDF_UNSP_DOC_SHAREDREVIEW = 5
const FPDF_UNSP_DOC_SHAREDFORM_ACROBAT = 6
const FPDF_UNSP_DOC_SHAREDFORM_FILESYSTEM = 7
const FPDF_UNSP_DOC_SHAREDFORM_EMAIL = 8
const FPDF_UNSP_ANNOT_3DANNOT = 11
const FPDF_UNSP_ANNOT_MOVIE = 12
const FPDF_UNSP_ANNOT_SOUND = 13
const FPDF_UNSP_ANNOT_SCREEN_MEDIA = 14
const FPDF_UNSP_ANNOT_SCREEN_RICHMEDIA = 15
const FPDF_UNSP_ANNOT_ATTACHMENT = 16
const FPDF_UNSP_ANNOT_SIG = 17

type _UNSUPPORT_INFO
	version as long
	FSDK_UnSupport_Handler as sub(byval pThis as _UNSUPPORT_INFO ptr, byval nType as long)
end type

type UNSUPPORT_INFO as _UNSUPPORT_INFO
declare function FSDK_SetUnSpObjProcessHandler(byval unsp_info as UNSUPPORT_INFO ptr) as FPDF_BOOL
declare sub FSDK_SetTimeFunction(byval func as function() as time_t)
declare sub FSDK_SetLocaltimeFunction(byval func as function(byval as const time_t ptr) as tm ptr)

const PAGEMODE_UNKNOWN = -1
const PAGEMODE_USENONE = 0
const PAGEMODE_USEOUTLINES = 1
const PAGEMODE_USETHUMBS = 2
const PAGEMODE_FULLSCREEN = 3
const PAGEMODE_USEOC = 4
const PAGEMODE_USEATTACHMENTS = 5
declare function FPDFDoc_GetPageMode(byval document as FPDF_DOCUMENT) as long
#define PUBLIC_FPDF_FLATTEN_H_
const FLATTEN_FAIL = 0
const FLATTEN_SUCCESS = 1
const FLATTEN_NOTHINGTODO = 2
const FLAT_NORMALDISPLAY = 0
const FLAT_PRINT = 1
declare function FPDFPage_Flatten(byval page as FPDF_PAGE, byval nFlag as long) as long
#define PUBLIC_FPDF_FWLEVENT_H_

type FWL_EVENTFLAG as long
enum
	FWL_EVENTFLAG_ShiftKey = 1 shl 0
	FWL_EVENTFLAG_ControlKey = 1 shl 1
	FWL_EVENTFLAG_AltKey = 1 shl 2
	FWL_EVENTFLAG_MetaKey = 1 shl 3
	FWL_EVENTFLAG_KeyPad = 1 shl 4
	FWL_EVENTFLAG_AutoRepeat = 1 shl 5
	FWL_EVENTFLAG_LeftButtonDown = 1 shl 6
	FWL_EVENTFLAG_MiddleButtonDown = 1 shl 7
	FWL_EVENTFLAG_RightButtonDown = 1 shl 8
end enum

type FWL_VKEYCODE as long
enum
	FWL_VKEY_Back = &h08
	FWL_VKEY_Tab = &h09
	FWL_VKEY_NewLine = &h0A
	FWL_VKEY_Clear = &h0C
	FWL_VKEY_Return = &h0D
	FWL_VKEY_Shift = &h10
	FWL_VKEY_Control = &h11
	FWL_VKEY_Menu = &h12
	FWL_VKEY_Pause = &h13
	FWL_VKEY_Capital = &h14
	FWL_VKEY_Kana = &h15
	FWL_VKEY_Hangul = &h15
	FWL_VKEY_Junja = &h17
	FWL_VKEY_Final = &h18
	FWL_VKEY_Hanja = &h19
	FWL_VKEY_Kanji = &h19
	FWL_VKEY_Escape = &h1B
	FWL_VKEY_Convert = &h1C
	FWL_VKEY_NonConvert = &h1D
	FWL_VKEY_Accept = &h1E
	FWL_VKEY_ModeChange = &h1F
	FWL_VKEY_Space = &h20
	FWL_VKEY_Prior = &h21
	FWL_VKEY_Next = &h22
	FWL_VKEY_End = &h23
	FWL_VKEY_Home = &h24
	FWL_VKEY_Left = &h25
	FWL_VKEY_Up = &h26
	FWL_VKEY_Right = &h27
	FWL_VKEY_Down = &h28
	FWL_VKEY_Select = &h29
	FWL_VKEY_Print = &h2A
	FWL_VKEY_Execute = &h2B
	FWL_VKEY_Snapshot = &h2C
	FWL_VKEY_Insert = &h2D
	FWL_VKEY_Delete = &h2E
	FWL_VKEY_Help = &h2F
	FWL_VKEY_0 = &h30
	FWL_VKEY_1 = &h31
	FWL_VKEY_2 = &h32
	FWL_VKEY_3 = &h33
	FWL_VKEY_4 = &h34
	FWL_VKEY_5 = &h35
	FWL_VKEY_6 = &h36
	FWL_VKEY_7 = &h37
	FWL_VKEY_8 = &h38
	FWL_VKEY_9 = &h39
	FWL_VKEY_A = &h41
	FWL_VKEY_B = &h42
	FWL_VKEY_C = &h43
	FWL_VKEY_D = &h44
	FWL_VKEY_E = &h45
	FWL_VKEY_F = &h46
	FWL_VKEY_G = &h47
	FWL_VKEY_H = &h48
	FWL_VKEY_I = &h49
	FWL_VKEY_J = &h4A
	FWL_VKEY_K = &h4B
	FWL_VKEY_L = &h4C
	FWL_VKEY_M = &h4D
	FWL_VKEY_N = &h4E
	FWL_VKEY_O = &h4F
	FWL_VKEY_P = &h50
	FWL_VKEY_Q = &h51
	FWL_VKEY_R = &h52
	FWL_VKEY_S = &h53
	FWL_VKEY_T = &h54
	FWL_VKEY_U = &h55
	FWL_VKEY_V = &h56
	FWL_VKEY_W = &h57
	FWL_VKEY_X = &h58
	FWL_VKEY_Y = &h59
	FWL_VKEY_Z = &h5A
	FWL_VKEY_LWin = &h5B
	FWL_VKEY_Command = &h5B
	FWL_VKEY_RWin = &h5C
	FWL_VKEY_Apps = &h5D
	FWL_VKEY_Sleep = &h5F
	FWL_VKEY_NumPad0 = &h60
	FWL_VKEY_NumPad1 = &h61
	FWL_VKEY_NumPad2 = &h62
	FWL_VKEY_NumPad3 = &h63
	FWL_VKEY_NumPad4 = &h64
	FWL_VKEY_NumPad5 = &h65
	FWL_VKEY_NumPad6 = &h66
	FWL_VKEY_NumPad7 = &h67
	FWL_VKEY_NumPad8 = &h68
	FWL_VKEY_NumPad9 = &h69
	FWL_VKEY_Multiply = &h6A
	FWL_VKEY_Add = &h6B
	FWL_VKEY_Separator = &h6C
	FWL_VKEY_Subtract = &h6D
	FWL_VKEY_Decimal = &h6E
	FWL_VKEY_Divide = &h6F
	FWL_VKEY_F1 = &h70
	FWL_VKEY_F2 = &h71
	FWL_VKEY_F3 = &h72
	FWL_VKEY_F4 = &h73
	FWL_VKEY_F5 = &h74
	FWL_VKEY_F6 = &h75
	FWL_VKEY_F7 = &h76
	FWL_VKEY_F8 = &h77
	FWL_VKEY_F9 = &h78
	FWL_VKEY_F10 = &h79
	FWL_VKEY_F11 = &h7A
	FWL_VKEY_F12 = &h7B
	FWL_VKEY_F13 = &h7C
	FWL_VKEY_F14 = &h7D
	FWL_VKEY_F15 = &h7E
	FWL_VKEY_F16 = &h7F
	FWL_VKEY_F17 = &h80
	FWL_VKEY_F18 = &h81
	FWL_VKEY_F19 = &h82
	FWL_VKEY_F20 = &h83
	FWL_VKEY_F21 = &h84
	FWL_VKEY_F22 = &h85
	FWL_VKEY_F23 = &h86
	FWL_VKEY_F24 = &h87
	FWL_VKEY_NunLock = &h90
	FWL_VKEY_Scroll = &h91
	FWL_VKEY_LShift = &hA0
	FWL_VKEY_RShift = &hA1
	FWL_VKEY_LControl = &hA2
	FWL_VKEY_RControl = &hA3
	FWL_VKEY_LMenu = &hA4
	FWL_VKEY_RMenu = &hA5
	FWL_VKEY_BROWSER_Back = &hA6
	FWL_VKEY_BROWSER_Forward = &hA7
	FWL_VKEY_BROWSER_Refresh = &hA8
	FWL_VKEY_BROWSER_Stop = &hA9
	FWL_VKEY_BROWSER_Search = &hAA
	FWL_VKEY_BROWSER_Favorites = &hAB
	FWL_VKEY_BROWSER_Home = &hAC
	FWL_VKEY_VOLUME_Mute = &hAD
	FWL_VKEY_VOLUME_Down = &hAE
	FWL_VKEY_VOLUME_Up = &hAF
	FWL_VKEY_MEDIA_NEXT_Track = &hB0
	FWL_VKEY_MEDIA_PREV_Track = &hB1
	FWL_VKEY_MEDIA_Stop = &hB2
	FWL_VKEY_MEDIA_PLAY_Pause = &hB3
	FWL_VKEY_MEDIA_LAUNCH_Mail = &hB4
	FWL_VKEY_MEDIA_LAUNCH_MEDIA_Select = &hB5
	FWL_VKEY_MEDIA_LAUNCH_APP1 = &hB6
	FWL_VKEY_MEDIA_LAUNCH_APP2 = &hB7
	FWL_VKEY_OEM_1 = &hBA
	FWL_VKEY_OEM_Plus = &hBB
	FWL_VKEY_OEM_Comma = &hBC
	FWL_VKEY_OEM_Minus = &hBD
	FWL_VKEY_OEM_Period = &hBE
	FWL_VKEY_OEM_2 = &hBF
	FWL_VKEY_OEM_3 = &hC0
	FWL_VKEY_OEM_4 = &hDB
	FWL_VKEY_OEM_5 = &hDC
	FWL_VKEY_OEM_6 = &hDD
	FWL_VKEY_OEM_7 = &hDE
	FWL_VKEY_OEM_8 = &hDF
	FWL_VKEY_OEM_102 = &hE2
	FWL_VKEY_ProcessKey = &hE5
	FWL_VKEY_Packet = &hE7
	FWL_VKEY_Attn = &hF6
	FWL_VKEY_Crsel = &hF7
	FWL_VKEY_Exsel = &hF8
	FWL_VKEY_Ereof = &hF9
	FWL_VKEY_Play = &hFA
	FWL_VKEY_Zoom = &hFB
	FWL_VKEY_NoName = &hFC
	FWL_VKEY_PA1 = &hFD
	FWL_VKEY_OEM_Clear = &hFE
	FWL_VKEY_Unknown = 0
end enum

#define PUBLIC_FPDF_JAVASCRIPT_H_
declare function FPDFDoc_GetJavaScriptActionCount(byval document as FPDF_DOCUMENT) as long
declare function FPDFDoc_GetJavaScriptAction(byval document as FPDF_DOCUMENT, byval index as long) as FPDF_JAVASCRIPT_ACTION
declare sub FPDFDoc_CloseJavaScriptAction(byval javascript as FPDF_JAVASCRIPT_ACTION)
declare function FPDFJavaScriptAction_GetName(byval javascript as FPDF_JAVASCRIPT_ACTION, byval buffer as FPDF_WCHAR ptr, byval buflen as culong) as culong
declare function FPDFJavaScriptAction_GetScript(byval javascript as FPDF_JAVASCRIPT_ACTION, byval buffer as FPDF_WCHAR ptr, byval buflen as culong) as culong
#define PUBLIC_FPDF_PPO_H_
declare function FPDF_ImportPages(byval dest_doc as FPDF_DOCUMENT, byval src_doc as FPDF_DOCUMENT, byval pagerange as FPDF_BYTESTRING, byval index as long) as FPDF_BOOL
declare function FPDF_ImportNPagesToOne(byval src_doc as FPDF_DOCUMENT, byval output_width as single, byval output_height as single, byval num_pages_on_x_axis as uinteger, byval num_pages_on_y_axis as uinteger) as FPDF_DOCUMENT
declare function FPDF_CopyViewerPreferences(byval dest_doc as FPDF_DOCUMENT, byval src_doc as FPDF_DOCUMENT) as FPDF_BOOL

#define PUBLIC_FPDF_PROGRESSIVE_H_
const FPDF_RENDER_READY = 0
const FPDF_RENDER_TOBECONTINUED = 1
const FPDF_RENDER_DONE = 2
const FPDF_RENDER_FAILED = 3

type _IFSDK_PAUSE
	version as long
	NeedToPauseNow as function(byval pThis as _IFSDK_PAUSE ptr) as FPDF_BOOL
	user as any ptr
end type

type IFSDK_PAUSE as _IFSDK_PAUSE
declare function FPDF_RenderPageBitmapWithColorScheme_Start(byval bitmap as FPDF_BITMAP, byval page as FPDF_PAGE, byval start_x as long, byval start_y as long, byval size_x as long, byval size_y as long, byval rotate as long, byval flags as long, byval color_scheme as const FPDF_COLORSCHEME ptr, byval pause as IFSDK_PAUSE ptr) as long
declare function FPDF_RenderPageBitmap_Start(byval bitmap as FPDF_BITMAP, byval page as FPDF_PAGE, byval start_x as long, byval start_y as long, byval size_x as long, byval size_y as long, byval rotate as long, byval flags as long, byval pause as IFSDK_PAUSE ptr) as long
declare function FPDF_RenderPage_Continue(byval page as FPDF_PAGE, byval pause as IFSDK_PAUSE ptr) as long
declare sub FPDF_RenderPage_Close(byval page as FPDF_PAGE)
#define PUBLIC_FPDF_SAVE_H_

type FPDF_FILEWRITE_
	version as long
	WriteBlock as function(byval pThis as FPDF_FILEWRITE_ ptr, byval pData as const any ptr, byval size as culong) as long
end type

type FPDF_FILEWRITE as FPDF_FILEWRITE_
const FPDF_INCREMENTAL = 1
const FPDF_NO_INCREMENTAL = 2
const FPDF_REMOVE_SECURITY = 3
declare function FPDF_SaveAsCopy(byval document as FPDF_DOCUMENT, byval pFileWrite as FPDF_FILEWRITE ptr, byval flags as FPDF_DWORD) as FPDF_BOOL
declare function FPDF_SaveWithVersion(byval document as FPDF_DOCUMENT, byval pFileWrite as FPDF_FILEWRITE ptr, byval flags as FPDF_DWORD, byval fileVersion as long) as FPDF_BOOL
#define PUBLIC_FPDF_SEARCHEX_H_
declare function FPDFText_GetCharIndexFromTextIndex(byval text_page as FPDF_TEXTPAGE, byval nTextIndex as long) as long
declare function FPDFText_GetTextIndexFromCharIndex(byval text_page as FPDF_TEXTPAGE, byval nCharIndex as long) as long
#define PUBLIC_FPDF_SIGNATURE_H_

declare function FPDF_GetSignatureCount(byval document as FPDF_DOCUMENT) as long
declare function FPDF_GetSignatureObject(byval document as FPDF_DOCUMENT, byval index as long) as FPDF_SIGNATURE
declare function FPDFSignatureObj_GetContents(byval signature as FPDF_SIGNATURE, byval buffer as any ptr, byval length as culong) as culong
declare function FPDFSignatureObj_GetByteRange(byval signature as FPDF_SIGNATURE, byval buffer as long ptr, byval length as culong) as culong
declare function FPDFSignatureObj_GetSubFilter(byval signature as FPDF_SIGNATURE, byval buffer as zstring ptr, byval length as culong) as culong
declare function FPDFSignatureObj_GetReason(byval signature as FPDF_SIGNATURE, byval buffer as any ptr, byval length as culong) as culong
declare function FPDFSignatureObj_GetTime(byval signature as FPDF_SIGNATURE, byval buffer as zstring ptr, byval length as culong) as culong
#define PUBLIC_FPDF_STRUCTTREE_H_
declare function FPDF_StructTree_GetForPage(byval page as FPDF_PAGE) as FPDF_STRUCTTREE
declare sub FPDF_StructTree_Close(byval struct_tree as FPDF_STRUCTTREE)
declare function FPDF_StructTree_CountChildren(byval struct_tree as FPDF_STRUCTTREE) as long
declare function FPDF_StructTree_GetChildAtIndex(byval struct_tree as FPDF_STRUCTTREE, byval index as long) as FPDF_STRUCTELEMENT
declare function FPDF_StructElement_GetAltText(byval struct_element as FPDF_STRUCTELEMENT, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDF_StructElement_GetID(byval struct_element as FPDF_STRUCTELEMENT, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDF_StructElement_GetLang(byval struct_element as FPDF_STRUCTELEMENT, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDF_StructElement_GetStringAttribute(byval struct_element as FPDF_STRUCTELEMENT, byval attr_name as FPDF_BYTESTRING, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDF_StructElement_GetMarkedContentID(byval struct_element as FPDF_STRUCTELEMENT) as long
declare function FPDF_StructElement_GetType(byval struct_element as FPDF_STRUCTELEMENT, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDF_StructElement_GetTitle(byval struct_element as FPDF_STRUCTELEMENT, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDF_StructElement_CountChildren(byval struct_element as FPDF_STRUCTELEMENT) as long
declare function FPDF_StructElement_GetChildAtIndex(byval struct_element as FPDF_STRUCTELEMENT, byval index as long) as FPDF_STRUCTELEMENT

#define PUBLIC_FPDF_SYSFONTINFO_H_
const FXFONT_ANSI_CHARSET = 0
const FXFONT_DEFAULT_CHARSET = 1
const FXFONT_SYMBOL_CHARSET = 2
const FXFONT_SHIFTJIS_CHARSET = 128
const FXFONT_HANGEUL_CHARSET = 129
const FXFONT_GB2312_CHARSET = 134
const FXFONT_CHINESEBIG5_CHARSET = 136
const FXFONT_ARABIC_CHARSET = 178
const FXFONT_CYRILLIC_CHARSET = 204
const FXFONT_EASTERNEUROPEAN_CHARSET = 238
const FXFONT_FF_FIXEDPITCH = 1 shl 0
const FXFONT_FF_ROMAN = 1 shl 4
const FXFONT_FF_SCRIPT = 4 shl 4
const FXFONT_FW_NORMAL = 400
const FXFONT_FW_BOLD = 700

type _FPDF_SYSFONTINFO
	version as long
	Release as sub(byval pThis as _FPDF_SYSFONTINFO ptr)
	EnumFonts as sub(byval pThis as _FPDF_SYSFONTINFO ptr, byval pMapper as any ptr)
	MapFont as function(byval pThis as _FPDF_SYSFONTINFO ptr, byval weight as long, byval bItalic as FPDF_BOOL, byval charset as long, byval pitch_family as long, byval face as const zstring ptr, byval bExact as FPDF_BOOL ptr) as any ptr
	GetFont as function(byval pThis as _FPDF_SYSFONTINFO ptr, byval face as const zstring ptr) as any ptr
	GetFontData as function(byval pThis as _FPDF_SYSFONTINFO ptr, byval hFont as any ptr, byval table as ulong, byval buffer as ubyte ptr, byval buf_size as culong) as culong
	GetFaceName as function(byval pThis as _FPDF_SYSFONTINFO ptr, byval hFont as any ptr, byval buffer as zstring ptr, byval buf_size as culong) as culong
	GetFontCharset as function(byval pThis as _FPDF_SYSFONTINFO ptr, byval hFont as any ptr) as long
	DeleteFont as sub(byval pThis as _FPDF_SYSFONTINFO ptr, byval hFont as any ptr)
end type

type FPDF_SYSFONTINFO as _FPDF_SYSFONTINFO

type FPDF_CharsetFontMap_
	charset as long
	fontname as const zstring ptr
end type

type FPDF_CharsetFontMap as FPDF_CharsetFontMap_
declare function FPDF_GetDefaultTTFMap() as const FPDF_CharsetFontMap ptr
declare sub FPDF_AddInstalledFont(byval mapper as any ptr, byval face as const zstring ptr, byval charset as long)
declare sub FPDF_SetSystemFontInfo(byval pFontInfo as FPDF_SYSFONTINFO ptr)
declare function FPDF_GetDefaultSystemFontInfo() as FPDF_SYSFONTINFO ptr
declare sub FPDF_FreeDefaultSystemFontInfo(byval pFontInfo as FPDF_SYSFONTINFO ptr)
#define PUBLIC_FPDF_TEXT_H_
declare function FPDFText_LoadPage(byval page as FPDF_PAGE) as FPDF_TEXTPAGE
declare sub FPDFText_ClosePage(byval text_page as FPDF_TEXTPAGE)
declare function FPDFText_CountChars(byval text_page as FPDF_TEXTPAGE) as long
declare function FPDFText_GetUnicode(byval text_page as FPDF_TEXTPAGE, byval index as long) as ulong
declare function FPDFText_GetFontSize(byval text_page as FPDF_TEXTPAGE, byval index as long) as double
declare function FPDFText_GetFontInfo(byval text_page as FPDF_TEXTPAGE, byval index as long, byval buffer as any ptr, byval buflen as culong, byval flags as long ptr) as culong
declare function FPDFText_GetFontWeight(byval text_page as FPDF_TEXTPAGE, byval index as long) as long
declare function FPDFText_GetTextRenderMode(byval text_page as FPDF_TEXTPAGE, byval index as long) as FPDF_TEXT_RENDERMODE
declare function FPDFText_GetFillColor(byval text_page as FPDF_TEXTPAGE, byval index as long, byval R as ulong ptr, byval G as ulong ptr, byval B as ulong ptr, byval A as ulong ptr) as FPDF_BOOL
declare function FPDFText_GetStrokeColor(byval text_page as FPDF_TEXTPAGE, byval index as long, byval R as ulong ptr, byval G as ulong ptr, byval B as ulong ptr, byval A as ulong ptr) as FPDF_BOOL
declare function FPDFText_GetCharAngle(byval text_page as FPDF_TEXTPAGE, byval index as long) as single
declare function FPDFText_GetCharBox(byval text_page as FPDF_TEXTPAGE, byval index as long, byval left as double ptr, byval right as double ptr, byval bottom as double ptr, byval top as double ptr) as FPDF_BOOL
declare function FPDFText_GetLooseCharBox(byval text_page as FPDF_TEXTPAGE, byval index as long, byval rect as FS_RECTF ptr) as FPDF_BOOL
declare function FPDFText_GetMatrix(byval text_page as FPDF_TEXTPAGE, byval index as long, byval matrix as FS_MATRIX ptr) as FPDF_BOOL
declare function FPDFText_GetCharOrigin(byval text_page as FPDF_TEXTPAGE, byval index as long, byval x as double ptr, byval y as double ptr) as FPDF_BOOL
declare function FPDFText_GetCharIndexAtPos(byval text_page as FPDF_TEXTPAGE, byval x as double, byval y as double, byval xTolerance as double, byval yTolerance as double) as long
declare function FPDFText_GetText(byval text_page as FPDF_TEXTPAGE, byval start_index as long, byval count as long, byval result as ushort ptr) as long
declare function FPDFText_CountRects(byval text_page as FPDF_TEXTPAGE, byval start_index as long, byval count as long) as long
declare function FPDFText_GetRect(byval text_page as FPDF_TEXTPAGE, byval rect_index as long, byval left as double ptr, byval top as double ptr, byval right as double ptr, byval bottom as double ptr) as FPDF_BOOL
declare function FPDFText_GetBoundedText(byval text_page as FPDF_TEXTPAGE, byval left as double, byval top as double, byval right as double, byval bottom as double, byval buffer as ushort ptr, byval buflen as long) as long

const FPDF_MATCHCASE = &h00000001
const FPDF_MATCHWHOLEWORD = &h00000002
const FPDF_CONSECUTIVE = &h00000004

declare function FPDFText_FindStart(byval text_page as FPDF_TEXTPAGE, byval findwhat as FPDF_WIDESTRING, byval flags as culong, byval start_index as long) as FPDF_SCHHANDLE
declare function FPDFText_FindNext(byval handle as FPDF_SCHHANDLE) as FPDF_BOOL
declare function FPDFText_FindPrev(byval handle as FPDF_SCHHANDLE) as FPDF_BOOL
declare function FPDFText_GetSchResultIndex(byval handle as FPDF_SCHHANDLE) as long
declare function FPDFText_GetSchCount(byval handle as FPDF_SCHHANDLE) as long
declare sub FPDFText_FindClose(byval handle as FPDF_SCHHANDLE)
declare function FPDFLink_LoadWebLinks(byval text_page as FPDF_TEXTPAGE) as FPDF_PAGELINK
declare function FPDFLink_CountWebLinks(byval link_page as FPDF_PAGELINK) as long
declare function FPDFLink_GetURL(byval link_page as FPDF_PAGELINK, byval link_index as long, byval buffer as ushort ptr, byval buflen as long) as long
declare function FPDFLink_CountRects(byval link_page as FPDF_PAGELINK, byval link_index as long) as long
declare function FPDFLink_GetRect(byval link_page as FPDF_PAGELINK, byval link_index as long, byval rect_index as long, byval left as double ptr, byval top as double ptr, byval right as double ptr, byval bottom as double ptr) as FPDF_BOOL
declare function FPDFLink_GetTextRange(byval link_page as FPDF_PAGELINK, byval link_index as long, byval start_char_index as long ptr, byval char_count as long ptr) as FPDF_BOOL
declare sub FPDFLink_CloseWebLinks(byval link_page as FPDF_PAGELINK)
#define PUBLIC_FPDF_THUMBNAIL_H_
declare function FPDFPage_GetDecodedThumbnailData(byval page as FPDF_PAGE, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDFPage_GetRawThumbnailData(byval page as FPDF_PAGE, byval buffer as any ptr, byval buflen as culong) as culong
declare function FPDFPage_GetThumbnailAsBitmap(byval page as FPDF_PAGE) as FPDF_BITMAP
#define PUBLIC_FPDF_TRANSFORMPAGE_H_
declare sub FPDFPage_SetMediaBox(byval page as FPDF_PAGE, byval left as single, byval bottom as single, byval right as single, byval top as single)
declare sub FPDFPage_SetCropBox(byval page as FPDF_PAGE, byval left as single, byval bottom as single, byval right as single, byval top as single)
declare sub FPDFPage_SetBleedBox(byval page as FPDF_PAGE, byval left as single, byval bottom as single, byval right as single, byval top as single)
declare sub FPDFPage_SetTrimBox(byval page as FPDF_PAGE, byval left as single, byval bottom as single, byval right as single, byval top as single)
declare sub FPDFPage_SetArtBox(byval page as FPDF_PAGE, byval left as single, byval bottom as single, byval right as single, byval top as single)
declare function FPDFPage_GetMediaBox(byval page as FPDF_PAGE, byval left as single ptr, byval bottom as single ptr, byval right as single ptr, byval top as single ptr) as FPDF_BOOL
declare function FPDFPage_GetCropBox(byval page as FPDF_PAGE, byval left as single ptr, byval bottom as single ptr, byval right as single ptr, byval top as single ptr) as FPDF_BOOL
declare function FPDFPage_GetBleedBox(byval page as FPDF_PAGE, byval left as single ptr, byval bottom as single ptr, byval right as single ptr, byval top as single ptr) as FPDF_BOOL
declare function FPDFPage_GetTrimBox(byval page as FPDF_PAGE, byval left as single ptr, byval bottom as single ptr, byval right as single ptr, byval top as single ptr) as FPDF_BOOL
declare function FPDFPage_GetArtBox(byval page as FPDF_PAGE, byval left as single ptr, byval bottom as single ptr, byval right as single ptr, byval top as single ptr) as FPDF_BOOL
declare function FPDFPage_TransFormWithClip(byval page as FPDF_PAGE, byval matrix as const FS_MATRIX ptr, byval clipRect as const FS_RECTF ptr) as FPDF_BOOL
declare sub FPDFPageObj_TransformClipPath(byval page_object as FPDF_PAGEOBJECT, byval a as double, byval b as double, byval c as double, byval d as double, byval e as double, byval f as double)
declare function FPDFPageObj_GetClipPath(byval page_object as FPDF_PAGEOBJECT) as FPDF_CLIPPATH
declare function FPDFClipPath_CountPaths(byval clip_path as FPDF_CLIPPATH) as long
declare function FPDFClipPath_CountPathSegments(byval clip_path as FPDF_CLIPPATH, byval path_index as long) as long
declare function FPDFClipPath_GetPathSegment(byval clip_path as FPDF_CLIPPATH, byval path_index as long, byval segment_index as long) as FPDF_PATHSEGMENT
declare function FPDF_CreateClipPath(byval left as single, byval bottom as single, byval right as single, byval top as single) as FPDF_CLIPPATH
declare sub FPDF_DestroyClipPath(byval clipPath as FPDF_CLIPPATH)
declare sub FPDFPage_InsertClipPath(byval page as FPDF_PAGE, byval clipPath as FPDF_CLIPPATH)

end extern

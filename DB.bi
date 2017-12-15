#include once "sqlite3.bi" 
#include once "list.bi" 
#include once "Dict.bi" 
#include once "VarBox.bi"

type TDB_ as TDB

type TDbStmt
	declare constructor(db as sqlite3 ptr)
	declare destructor()
	declare function prepare(query as const zstring ptr) as boolean
	declare sub bind(index as integer, value as integer)
	declare sub bind(index as integer, value as longint)
	declare sub bind(index as integer, value as double)
	declare sub bind(index as integer, value as const zstring ptr)
	declare sub bind(index as integer, value as const wstring ptr)
	declare sub bindNull(index as integer)
	declare sub clear_()
	declare function step_() as long
	declare function colCount() as integer
	declare function colName(index as integer) as const zstring ptr
	declare function colValue(index as integer) as const zstring ptr
	declare sub reset()
private:
	db				as sqlite3 ptr
	stmt			as sqlite3_stmt ptr = null
end type

type TRSetRow
	declare constructor()
	declare destructor()
	declare sub newColumn(name as const zstring ptr, value as const zstring ptr)
	declare operator [](index as const zstring ptr) as zstring ptr
	declare operator [](index as integer) as zstring ptr
private:
	columns			as TDict
	colList(any)	as zstring ptr
	colCnt			as integer
end type

type TRSet
	declare constructor()
	declare destructor()
	declare function newRow() as TRSetRow ptr
	declare function hasNext() as boolean
	declare sub next_()
	declare property row as TRSetRow ptr
	
private:
	rows			as TList		'' list of TRSetRow
	currRow			as TRSetRow ptr
end type

type TDb
	declare function open(fileName as const zstring ptr) as boolean
	declare function open() as boolean
	declare sub close()
	declare function getErrorMsg() as const zstring ptr
	declare function prepare(query as const zstring ptr) as TDbStmt ptr
	declare function format cdecl(fmt as string, ... /' of VarBox ptr '/) as string
	declare function exec(query as const zstring ptr) as TRSet ptr
	declare function exec(stmt as TDbStmt ptr) as TRSet ptr
	declare function execScalar(query as const zstring ptr) as zstring ptr
	declare sub execNonQuery(query as const zstring ptr) 
	declare sub execNonQuery(stmt as TDbStmt ptr)
	

private:
	instance 		as sqlite3 ptr 
	errMsg			as string
end type
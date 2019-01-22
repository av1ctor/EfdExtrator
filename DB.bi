#include once "sqlite3.bi" 
#include once "list.bi" 
#include once "Dict.bi" 
#include once "VarBox.bi"
#include once "Lua/lualib.bi"
#include once "Lua/lauxlib.bi"

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

type TDbColumn
	name			as zstring ptr
	value			as zstring ptr
end type

type TDataSetRow
	declare constructor(cols as integer = 0)
	declare destructor()
	declare sub newColumn(name as const zstring ptr, value as const zstring ptr)
	declare operator [](index as const zstring ptr) as zstring ptr
	declare operator [](index as integer) as zstring ptr
	cols(any)		as TDbColumn
	cnt				as integer
private:
	dict			as TDict
end type

type TDataSet
	declare constructor()
	declare destructor()
	declare function newRow(cols as integer = 0) as TDataSetRow ptr
	declare function hasNext() as boolean
	declare sub next_()
	declare property row as TDataSetRow ptr
	
	currRow			as TDataSetRow ptr
private:
	rows			as TList		'' list of TDataSetRow
end type

type TDb
	declare function open(fileName as const zstring ptr) as boolean
	declare function open() as boolean
	declare sub close()
	declare function getErrorMsg() as const zstring ptr
	declare function prepare(query as const zstring ptr) as TDbStmt ptr
	declare function format cdecl(fmt as string, ... /' of VarBox ptr '/) as string
	declare function exec(query as const zstring ptr) as TDataSet ptr
	declare function exec(stmt as TDbStmt ptr) as TDataSet ptr
	declare function execScalar(query as const zstring ptr) as zstring ptr
	declare function execNonQuery(query as const zstring ptr) as boolean
	declare function execNonQuery(stmt as TDbStmt ptr) as boolean
	declare static sub exportAPI(L as lua_State ptr)

private:
	instance 		as sqlite3 ptr 
	errMsg			as string
end type
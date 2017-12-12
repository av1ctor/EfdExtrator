#include once "sqlite3.bi" 
#include once "list.bi" 
#include once "hash.bi" 

type TRSetRow
	declare constructor()
	declare destructor()
	declare sub newColumn(name as const zstring ptr, value as const zstring ptr)
	declare operator [](index as const zstring ptr) as zstring ptr
	declare operator [](index as integer) as zstring ptr
private:
	columns			as THash
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
	declare sub close()
	declare function exec(query as const zstring ptr) as TRSet ptr
	declare function execScalar(query as const zstring ptr) as zstring ptr

private:
	instance 		as sqlite3 ptr 
	errMsg			as string
end type
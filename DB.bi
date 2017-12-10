#include once "sqlite3.bi" 
#include once "list.bi" 
#include once "hash.bi" 

type TRSetRow
	declare constructor()
	declare destructor()
	declare sub newColumn(name as zstring ptr, value as zstring ptr)
	declare operator [](index as string) as zstring ptr
private:
	columns		as THash
end type

type TRSet
	declare constructor()
	declare destructor()
	declare function newRow() as TRSetRow ptr
	declare function hasNext() as boolean
	declare sub next_()
	declare property row as TRSetRow ptr
	
private:
	rows		as TList		'' list of TRSetRow
	currRow		as TRSetRow ptr
end type

type TDb
	declare function open(fileName as string) as boolean
	declare sub close()
	declare function exec(query as string) as TRSet ptr

private:
	instance 		as sqlite3 ptr 
	errMsg			as string
end type
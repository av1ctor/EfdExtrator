'' Try/Catch implementation using sjlj for FreeBASIC
'' Copyright 2018 by Andre Victor (av1ctortv[@]gmail.com)
'' Licensed under GNU GPL-2.0-or-above

#include once "trycatch.bi"

extern "C"
	type __p_sig_fn_t as sub(byval as integer)
	declare function signal(byval as integer, byval as __p_sig_fn_t) as __p_sig_fn_t
end extern

type TryCatchCtx
	cur			as TryCatch ptr
	oldsigabrt  as __p_sig_fn_t
	oldsigsegv  as __p_sig_fn_t
	oldsigfpe	as __p_sig_fn_t
	oldsigill	as __p_sig_fn_t
	oldsigterm	as __p_sig_fn_t
	oldsigint	as __p_sig_fn_t
end type

	dim shared ctx as TryCatchCtx
	
private sub handler cdecl(byval sig as integer)
	signal(sig, @handler)
	var buf = @ctx.cur->buf
	if ctx.cur then
		ctx.cur = ctx.cur->old
	end if
	longjmp(buf, sig)
end sub

constructor TryCatch()
	this.old = ctx.cur
	ctx.cur = @this
end constructor

destructor TryCatch()
	if ctx.cur = @this then
		ctx.cur = this.old
	end if
end destructor

private sub gb_ctor () constructor
	ctx.oldsigabrt = signal(SIGABRT, @handler)
	ctx.oldsigsegv = signal(SIGSEGV, @handler)
	ctx.oldsigfpe = signal(SIGFPE, @handler)
	ctx.oldsigill = signal(SIGILL, @handler)
	ctx.oldsigterm = signal(SIGTERM, @handler)
	ctx.oldsigint = signal(SIGINT, @handler)
end sub

private sub gb_dtor () destructor
	signal(SIGABRT, ctx.oldsigabrt)
	signal(SIGSEGV, ctx.oldsigsegv)
	signal(SIGFPE, ctx.oldsigfpe)
	signal(SIGILL, ctx.oldsigill)
	signal(SIGTERM, ctx.oldsigterm)
	signal(SIGINT, ctx.oldsigint)
end sub


	

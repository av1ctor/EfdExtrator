#include once "crt/setjmp.bi"

#define	SIGINT		2	'' Interactive attention
#define	SIGILL		4	'' Illegal instruction
#define	SIGFPE		8	'' Floating point error
#define	SIGSEGV		11	'' Segmentation violation
#define	SIGTERM		15	'' Termination request
#define SIGBREAK	21	'' Control-break
#define	SIGABRT		22	'' Abnormal termination (abort)

#define __MERGEPARAM(a,b) a##b
#define TRY scope: dim as TryCatch __MERGEPARAM(__tryCatchCtx, __line__): select case setjmp(@__MERGEPARAM(__tryCatchCtx, __line__).buf): case 0:
#define CATCHSIG(x) case x:
#define CATCH case else:
#define ENDTRY end select: end scope

type TryCatch
public:
	declare constructor()
	declare destructor()
	buf			as jmp_buf
	old			as TryCatch ptr
end type


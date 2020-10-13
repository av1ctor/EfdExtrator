#define DdMmYyyy2Yyyy_Mm(s) (mid(s,1,4) + "-" + mid(s,5,2))

#define STR2CNPJ(s) (left(s,2) + "." + mid(s,3,3) + "." + mid(s,3+3,3) + "/" + mid(s,3+3+3,4) + "-" + right(s,2))

#define STR2CPF(s) (left(s,3) + "." + mid(s,4,3) + "." + mid(s,4+3,3) + "-" + right(s,2))

#define DBL2MONEYBR(d) (format(d,"#,#,0.00"))

#define MUNICIPIO2SIGLA(m) (iif(m >= 1100000 and m <= 5399999, ufCod2Sigla(m \ 100000), "EX"))

declare function csvDate2YYYYMMDD(s as zstring ptr) as string 
declare function ddMmYyyy2YyyyMmDd(s as const zstring ptr) as string
declare function yyyyMmDd2Datetime(s as const zstring ptr) as string 
declare function YyyyMmDd2DatetimeBR(s as const zstring ptr) as string 
declare function STR2IE(ie as string) as string
declare function tipoItem2Str(tipo as TipoItemId) as string
declare function dupstr(s as const zstring ptr) as zstring ptr
declare sub splitstr(Text As String, Delim As String = ",", Ret() As String)
declare function strreplace(byref text as string, byref a as string, byref b as string) as string
declare function UF_SIGLA2COD(s as zstring ptr) as integer
declare function loadstrings(fromFile as string, toArray() as string) as boolean
declare function latinToUtf16le(src as zstring ptr) as wstring ptr

extern as string ufCod2Sigla(11 to 53)
extern as TDict ufSigla2CodDict ptr
extern as string codSituacao2Str(0 to __TipoSituacao__LEN__-1)

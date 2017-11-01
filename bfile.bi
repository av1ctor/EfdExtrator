
const MAX_BUFFER_SIZE = 8192

#define NULL 0

type bfile
public:
   declare constructor ()
   declare function abrir(arquivo as string) as Boolean
   declare function criar(arquivo as string) as Boolean
   declare sub fechar()
   declare function tamanho() as longint
   declare function posicao() as longint
   declare function lerTudo() as string
   declare function ler(destino() as byte, lgt as integer) as integer
   declare function ler(destino as byte ptr, lgt as integer) as integer
   declare sub escrever(texto as string)
   declare sub escrever(src() as byte) 
   declare function temProximo() as Boolean
   declare property peek1() as Byte
   declare property char1() as Byte
   declare property char2() as string
   declare property char4() as string
   declare property char6() as string
   declare property char8() as string
   declare property char13() as string
   declare property char14() as string
   declare property char22() as string
   declare function nchar(caracteres as Integer, preenchimento as byte = asc("0")) as string
   declare function varchar(separador as Integer = asc("|")) as string
   declare function varint(separador as Integer = asc("|")) as longint
   declare function vardbl(separador as Integer = asc("|")) as double
   declare property dbl13_2() as double
   declare property dbl13_3() as double
   declare property dbl4_2() as double
   declare property int1() as integer
   declare property int2() as integer
   declare property int4() as integer
   declare property int6() as integer
   declare property int9() as integer
   declare property lng14() as longint
   declare function charCsv(separador as Integer = asc(","), qualificador as Integer = asc("""")) as string
   declare function intCsv(separador as Integer = asc(","), qualificador as Integer = asc("""")) as longint
   declare function dblCsv(separador as Integer = asc(","), qualificador as Integer = asc("""")) as double

private:
   fnum as integer = 0
   fpos as longint = 0
   blen as integer = 0
   bptr as zstring ptr = NULL
   buffer as zstring * MAX_BUFFER_SIZE+1
   
   res2 as string    = "00"
   res4 as string    = "0000"
   res5 as string    = "00000"
   res6 as string    = "000000"
   res8 as string    = "00000000"
   res9 as string    = "000000000"
   res13 as string   = "0000000000000"
   res14 as string   = "00000000000000"
   res22 as string   = "0000000000000000000000"

   declare sub init()
end type


#inclib "ssl_helper"
#inclib "crypto" 
#inclib "ssl" 
#inclib "kernel32" 
#inclib "user32"
#inclib "gdi32"
#inclib "advapi32"

extern "C++"

enum ALTNAME_ATTRIBUTES
	AN_ATT_CPF		= 0
	AN_ATT_CNPJ		= 1
	AN_ATT_EMAIL	= 2
end enum

type PKCS7 as any

type SSL_Helper
public:	
	declare constructor( )
	declare destructor( )
	declare function Load_P7K(fileName as const zstring ptr) as PKCS7 ptr
	declare function Load_P7K(buffer as ubyte ptr, lgt as integer) as PKCS7 ptr
	declare sub Free(p7 as PKCS7 ptr)
	declare function Get_CommonName(p7 as PKCS7 ptr) as zstring ptr
	declare function Get_AttributeFromAltName(p7 as PKCS7 ptr, attrib as ALTNAME_ATTRIBUTES) as zstring ptr
private:
	unused__ as byte
end type

end extern
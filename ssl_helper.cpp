// g++ ssl_helper.cpp -c 
// ar rcs libssl_helper.a ssl_helper.o
// ld myapp.o -lssl_helper -lcrypto -lssl -static -lkernel32 -luser32 -lgdi32 -ladvapi32

#include <openssl/bio.h>
#include <openssl/err.h>
#include <openssl/x509v3.h>
#include <string.h>

enum ALTNAME_ATTRIBUTES
{
	AN_ATT_CPF 		= 0,
	AN_ATT_CNPJ	 	= 1,
	AN_ATT_EMAIL	= 2
};

class SSL_Helper
{
public:	
	__cdecl SSL_Helper();
	
	__cdecl ~SSL_Helper();

	void __cdecl *Load_P7K(char const *fileName);

	void __cdecl *Load_P7K(unsigned char *buffer, int len);
	
	void __cdecl Free(void *p7);
	
	char __cdecl *Get_CommonName(void *p7);

	char __cdecl *Get_AttributeFromAltName(void *p7, ALTNAME_ATTRIBUTES attrib);
	
private:
	X509 *hGetTopCertFromP7K(PKCS7 *p7);

	char *hGetCommonNameFromSujectName(X509 *cert);
	
	char *hGetAttributeFromAltName(X509 *cert, ALTNAME_ATTRIBUTES attrib);

};

int cpf_nid = 0;
int cnpj_nid = 0;
int cpf_resp_nid = 0;

void hInitializeOpenSSL()
{
	if( cpf_nid == 0 ) 
	{
		OpenSSL_add_all_algorithms();
	
		cpf_nid = OBJ_create("2.16.76.1.3.1", "CPF", "usuarioCPF");
		cnpj_nid = OBJ_create("2.16.76.1.3.3", "CNPJ", "empresaCNPJ");
		cpf_resp_nid = OBJ_create("2.16.76.1.3.4", "CPFRESP", "responsavelCPF");
	}
}

__cdecl SSL_Helper::SSL_Helper()
{
	hInitializeOpenSSL();
}
	
__cdecl SSL_Helper::~SSL_Helper()
{
}

void __cdecl *SSL_Helper::Load_P7K(char const *fileName)
{
	BIO *in = BIO_new(BIO_s_file());

	BIO_read_filename(in, fileName);
	PKCS7 *p7 = d2i_PKCS7_bio(in, NULL);
	
	//BIO_free(in);
	
	return (void *)p7;
}

void __cdecl *SSL_Helper::Load_P7K(unsigned char *buffer, int len)
{
	BIO *in = BIO_new_mem_buf(buffer, len);
	PKCS7 *p7 = d2i_PKCS7_bio(in, NULL);
	
	//BIO_free(in);
	
	return (void *)p7;
}
	
void __cdecl SSL_Helper::Free(void *p7)
{
	PKCS7_free((PKCS7 *)p7);
}

char __cdecl *SSL_Helper::Get_CommonName(void *p7)
{
	return hGetCommonNameFromSujectName(hGetTopCertFromP7K((PKCS7 *)p7));
}

char __cdecl *SSL_Helper::Get_AttributeFromAltName(void *p7, ALTNAME_ATTRIBUTES attrib)
{
	return hGetAttributeFromAltName(hGetTopCertFromP7K((PKCS7 *)p7), attrib);
}

X509 *SSL_Helper::hGetTopCertFromP7K(PKCS7 *p7)
{
	int nid = OBJ_obj2nid(p7->type);
	STACK_OF(X509) *certs = NULL;
	if(nid == NID_pkcs7_signed) 
	{
		certs = p7->d.sign->cert;
	} 
	else if(nid == NID_pkcs7_signedAndEnveloped) 
	{
		certs = p7->d.signed_and_enveloped->cert;
	}
	
	for(int i = 0; certs && i < sk_X509_num(certs); i++ )
	{
		X509 *cert = sk_X509_value(certs, i);
		X509_check_purpose(cert, -1, 0);
		if( (cert->ex_kusage & X509v3_KU_DIGITAL_SIGNATURE) != 0 )
		{
			return cert;
		}
	}
	
	return NULL;
}

char *SSL_Helper::hGetCommonNameFromSujectName(X509 *cert)
{
	char *res = NULL;
	
	// common name
	ASN1_OBJECT* obj = OBJ_txt2obj("2.5.4.3", 0);

	X509_NAME* name = X509_get_subject_name(cert);
	
	int pos = -1;
	if((pos = X509_NAME_get_index_by_OBJ(name, obj, pos)) != -1) 
	{
		X509_NAME_ENTRY* name_entry = X509_NAME_get_entry(name, pos);
		char *entry = (char *)ASN1_STRING_data(X509_NAME_ENTRY_get_data(name_entry));

		int length = strlen(entry);
		res = (char *)malloc(length+1);
		strncpy(res, (char *)entry, length);
		res[length] = '\0';
	}
	
	ASN1_OBJECT_free(obj);
	
	return res;
}

char *SSL_Helper::hGetAttributeFromAltName(X509 *cert, ALTNAME_ATTRIBUTES attrib)
{
	char *res = NULL;
	GENERAL_NAMES* subjectAltNames = (GENERAL_NAMES*)X509_get_ext_d2i(cert, NID_subject_alt_name, NULL, NULL);
	
	for (int i = 0; (res == NULL) && (i < sk_GENERAL_NAME_num(subjectAltNames)); i++)
	{
		GENERAL_NAME* gen = sk_GENERAL_NAME_value(subjectAltNames, i);
		switch (gen->type)
		{
			case GEN_EMAIL:
			{
				if( attrib == AN_ATT_EMAIL )
				{
					ASN1_IA5STRING *asn1_str = gen->d.uniformResourceIdentifier;
					char *s = (char*)ASN1_STRING_data(asn1_str);
					int len = strlen(s);
					res = (char *)malloc(len+1);
					strncpy(res, s, len);
					res[len] = '\0';
				}
				break;
			}
			case GEN_OTHERNAME:
			{
				int nid = OBJ_obj2nid(gen->d.otherName->type_id);
				
				if( nid == cpf_nid )
				{
					if( attrib == AN_ATT_CPF )
					{
						/*
							Nas primeiras 8 (oito) posições, a data de nascimento da pessoa física titular do
							certificado, no formato ddmmaaaa; nas 11 (onze) posições subseqüentes, o número de
							inscrição no Cadastro de Pessoa Física (CPF) da pessoa física titular do certificado
						*/
						char *astr = (char*)ASN1_STRING_data(gen->d.otherName->value->value.asn1_string);			
						
						res = (char *)malloc(11+1);
						strncpy(res, &astr[8], 11);
						res[11] = '\0';
					}
				}
				else if( nid == cnpj_nid )
				{
					if( attrib == AN_ATT_CNPJ )
					{
						char *astr = (char*)ASN1_STRING_data(gen->d.otherName->value->value.asn1_string);			
					
						res = (char *)malloc(14+1);
						strncpy(res, &astr[0], 14);
						res[14] = '\0';
					}
				}
				else if( nid == cpf_resp_nid )
				{
					if( attrib == AN_ATT_CPF )
					{
						/*
							Nas primeiras 8 (oito) posições, a data de nascimento do responsável pela Pessoa
							Jurídica perante o CNPJ, no formato ddmmaaaa; nas 11 (onze) posições subseqüentes,
							o número de inscrição no Cadastro de Pessoas Físicas (CPF) do responsável pela
							Pessoa Jurídica perante o CNPJ; nas 11 (onze) posições subseqüentes o número de
							inscrição no NIS (PIS, PASEP ou CI) do responsável pela Pessoa Jurídica perante o
							CNPJ; nas 15 (quinze) posições subseqüentes, o número do Registro Geral (RG) do
							responsável pela Pessoa Jurídica perante o CNPJ; nas 6 (seis) posições subseqüentes,
							as siglas do órgão expedidor do RG e respectiva UF;
						*/
						char *astr = (char*)ASN1_STRING_data(gen->d.otherName->value->value.asn1_string);			
						
						res = (char *)malloc(11+1);
						strncpy(res, &astr[8], 11);
						res[11] = '\0';
					}
				}
				break;
			}
		}
	}
	
	GENERAL_NAMES_free(subjectAltNames);
	return res;
}

// g++ ssl_helper.cpp -c 
// ar rcs libssl_helper.a ssl_helper.o
// ld myapp.o -lssl_helper -lcrypto -lssl -static -lkernel32 -luser32 -lgdi32 -ladvapi32

#include <openssl/bio.h>
#include <openssl/err.h>
#include <openssl/x509v3.h>
#include <string.h>

enum ALTNAME_ATTRIBUTES
{
	AN_ATT_CPF,
	AN_ATT_CNPJ,
	AN_ATT_EMAIL
};

class SSL_Helper
{
public:	
	SSL_Helper();
	
	~SSL_Helper();

	void *Load_P7K(char const *fileName);

	void *Load_P7K(unsigned char *buffer, int len);
	
	void Free(void *p7);
	
	char *Get_CommonName(void *p7);

	char *Get_AttributeFromAltName(void *p7, ALTNAME_ATTRIBUTES attrib);
	
private:
	int cpf_nid;
	int cnpj_nid;
	int cpf_resp_nid;

	X509 *hGetTopCertFromP7K(PKCS7 *p7);

	char *hGetCommonNameFromSujectName(X509_NAME *name);
	
	char *hGetAttributeFromAltName(X509 *cert, ALTNAME_ATTRIBUTES attrib);

};

SSL_Helper::SSL_Helper()
{
	OpenSSL_add_all_algorithms();

	int cpf_nid = OBJ_create("2.16.76.1.3.1", "CPF", "usuarioCPF");
	int cnpj_nid = OBJ_create("2.16.76.1.3.3", "CNPJ", "empresaCNPJ");
	int cpf_resp_nid = OBJ_create("2.16.76.1.3.4", "CPFRESP", "responsavelCPF");
}
	
SSL_Helper::~SSL_Helper()
{
}

void *SSL_Helper::Load_P7K(char const *fileName)
{
	BIO *in = BIO_new(BIO_s_file());

	BIO_read_filename(in, fileName);
	PKCS7 *p7 = d2i_PKCS7_bio(in, NULL);
	
	BIO_free(in);
	
	return (void *)p7;
}

void *SSL_Helper::Load_P7K(unsigned char *buffer, int len)
{
	BIO *in = BIO_new(BIO_s_mem());

	BIO_read(in, buffer, len);
	PKCS7 *p7 = d2i_PKCS7_bio(in, NULL);
	
	BIO_free(in);
	
	return (void *)p7;
}
	
void SSL_Helper::Free(void *p7)
{
	PKCS7_free((PKCS7 *)p7);
}

char *SSL_Helper::Get_CommonName(void *p7)
{
	return hGetCommonNameFromSujectName(X509_get_subject_name(hGetTopCertFromP7K((PKCS7 *)p7)));
}

char *SSL_Helper::Get_AttributeFromAltName(void *p7, ALTNAME_ATTRIBUTES attrib)
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
	
	return sk_X509_value(certs, 0);
}

char *SSL_Helper::hGetCommonNameFromSujectName(X509_NAME *name)
{
	for (int i = 0; i < sk_X509_NAME_ENTRY_num(name->entries); i++) 
	{
		const X509_NAME_ENTRY *ne = sk_X509_NAME_ENTRY_value(name->entries, i);
		int nid = OBJ_obj2nid(ne->object);
		if( nid == NID_commonName )
		{
			int type = ne->value->type;
			int num = ne->value->length;
			unsigned char *q = ne->value->data;
			int gs_doit[4];
			
			if ((type == V_ASN1_GENERALSTRING) && ((num % 4) == 0)) 
			{
				gs_doit[0] = gs_doit[1] = gs_doit[2] = gs_doit[3] = 0;
				for (int j = 0; j < num; j++)
					if (q[j] != 0)
						gs_doit[j & 3] = 1;

				if (gs_doit[0] | gs_doit[1] | gs_doit[2])
					gs_doit[0] = gs_doit[1] = gs_doit[2] = gs_doit[3] = 1;
				else 
				{
					gs_doit[0] = gs_doit[1] = gs_doit[2] = 0;
					gs_doit[3] = 1;
				}
			} 
			else
				gs_doit[0] = gs_doit[1] = gs_doit[2] = gs_doit[3] = 1;
			
			int len = 0;
			for(int j = 0; j < num; j++) 
			{
				if (!gs_doit[j & 3])
					continue;			
				++len;
			}
			
			char *res, *p;
			res = p = (char *)malloc(len+1);
			
			for(int j = 0; j < num; j++) 
			{
				if (!gs_doit[j & 3])
					continue;			
				
				*(p++) = q[j];
			}
			
			*p = '\0';
			
			return res;
		}
	}
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
					res = (char *)malloc(strlen(s)+1);
					strncpy(res, s, strlen(s)+1);
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

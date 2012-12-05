#ifndef __XML_PROFILE_STRING_H__
#define __XML_PROFILE_STRING_H__ 1

#pragma once

#ifndef __msxml2_h__
#error this "xmlprofilestring.h" requires header file "msxml2.h" included first.
#endif // __RPCNDR_H_VERSION__


#ifndef XML_PROFILE_ROOT_ELEM
#define XML_PROFILE_ROOT_ELEM	L"XmlConfigProfile"
#endif  // XML_PROFILE_ROOT_ELEM

#ifndef XML_PROFILE_ROOT_NODE
#define XML_PROFILE_ROOT_NODE	L"<" XML_PROFILE_ROOT_ELEM L"></" XML_PROFILE_ROOT_ELEM L">"
#endif

#ifndef XML_PROFILE_VALUE
#define XML_PROFILE_VALUE		L"value"
#endif

EXTERN_C HRESULT WINAPI XmlCreateProfileDoc(LPCWSTR xmlFileName, IXMLDOMDocument2 ** ppXmlDoc);

EXTERN_C BOOL WINAPI XmlWriteSystemTime(IN IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection, LPCWSTR lpSubSection, LPCWSTR lpKeyName,
	IN const SYSTEMTIME *BeginTime);

EXTERN_C BOOL WINAPI
	XmlWriteProfileString (IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection, 
	LPCWSTR lpSubSection,
	LPCWSTR lpKeyName, 
	LPCWSTR lpString
	);

EXTERN_C BOOL WINAPI
	XmlWriteProfileInt (IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection, 
	LPCWSTR lpSubSection,
	LPCWSTR lpKeyName, 
	INT nValue
	);

EXTERN_C BOOL WINAPI
	XmlRemoveProfileString (IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection, 
	LPCWSTR lpSubSection,
	LPCWSTR lpKeyName
	);

EXTERN_C BOOL WINAPI 
	XmlWriteProfileSection(IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection,
	LPCWSTR lpSubSection,
	LPCWSTR lpString
	);

EXTERN_C DWORD WINAPI
	XmlGetProfileString(IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection,
	LPCWSTR lpSubSection,
	LPCWSTR lpKeyName,
	LPCWSTR lpDefault,
	LPWSTR lpReturnedString,
	DWORD cchMaxReturn
	);

EXTERN_C DWORD WINAPI
XmlGetProfileInt(IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection,
	LPCWSTR lpSubSection,
	LPCWSTR lpKeyName,
	INT nDefault
	);

EXTERN_C HRESULT WINAPI
XmlMergeDocument(IN IXMLDOMDocument2 * pSrcDoc1,
	IN IXMLDOMDocument2 * pSrcDoc2,
	OUT IXMLDOMDocument2 ** ppNewDoc
	);

EXTERN_C HRESULT WINAPI 
ConvertXmlEncodingUTF8(IXMLDOMDocument2 * spXmlDoc);


#endif	// __XML_PROFILE_STRING_H__

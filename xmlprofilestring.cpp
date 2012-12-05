#include <windows.h>

#include <atlbase.h>

#include <Shlwapi.h>
#pragma comment(lib, "Shlwapi.lib")

//#include <msxml.h>
//#define __USE_MSXML2_NAMESPACE__	1
#include <msxml2.h>
//#import "MSXML4.dll" no_smart_pointers raw_interfaces_only no_implementation no_namespace // rename_namespace(_T("MSXML"))

#include "xmlprofilestring.h"

#define MS_XML_DLL		_T("msxml2.dll")

#ifndef MS_XML_DOMDOC_CLASS
#define MS_XML_DOMDOC_CLASS DOMDocument
#endif	// MS_XML_DOMDOC_CLASS


#ifndef _countof
#define _countof(_Array) (sizeof(_Array) / sizeof((_Array)[0]))
#endif

#if defined(_MSC_VER)
#pragma comment(linker, "/EXPORT:XmlCreateProfileDoc=_XmlCreateProfileDoc@8") 
#pragma comment(linker, "/EXPORT:XmlWriteSystemTime=_XmlWriteSystemTime@20") 
#pragma comment(linker, "/EXPORT:XmlWriteProfileString=_XmlWriteProfileString@20") 
#pragma comment(linker, "/EXPORT:XmlWriteProfileInt=_XmlWriteProfileInt@20") 
#pragma comment(linker, "/EXPORT:XmlRemoveProfileString=_XmlRemoveProfileString@16") 
#pragma comment(linker, "/EXPORT:XmlWriteProfileSection=_XmlWriteProfileSection@16") 
#pragma comment(linker, "/EXPORT:XmlGetProfileString=_XmlGetProfileString@28") 
#pragma comment(linker, "/EXPORT:XmlGetProfileInt=_XmlGetProfileInt@20") 
#pragma comment(linker, "/EXPORT:XmlMergeDocument=_XmlMergeDocument@12") 
#pragma comment(linker, "/EXPORT:ConvertXmlEncodingUTF8=_ConvertXmlEncodingUTF8@4") 
#endif  // _MSC_VER


HRESULT GetOrCreateChildElement(IXMLDOMDocument2 * pDoc, 
	IXMLDOMNode * pParent, 
	LPCWSTR lpszName, 
	IXMLDOMNode ** ppOut);
HRESULT ClearOldAndCreateNewChildElement(IXMLDOMDocument2 * pDoc, 
	IXMLDOMNode * pParent, 
	LPCWSTR lpszName, 
	IXMLDOMNode ** ppOut);
HRESULT CreateChildElement(IXMLDOMDocument2 * pDoc, 
	IXMLDOMNode * pParent, 
	LPCWSTR lpszName, 
	IXMLDOMNode ** ppOut);

HRESULT WINAPI FormatXmlDoc(IXMLDOMDocument2 * pXmlDoc, IXMLDOMDocument2 ** ppFormattedDoc);

EXTERN_C HRESULT WINAPI XmlCreateProfileDoc(LPCWSTR xmlFileName, IXMLDOMDocument2 ** ppXmlDoc)
{
    HRESULT hr = E_FAIL;
    do 
    {
        CComPtr<IXMLDOMDocument2> spXMLDoc;
        hr = spXMLDoc.CoCreateInstance(__uuidof(MS_XML_DOMDOC_CLASS));
        if (FAILED(hr)) { break; }

        VARIANT_BOOL bSucc = VARIANT_FALSE;
        spXMLDoc->load(CComVariant(xmlFileName), &bSucc);
        if (VARIANT_FALSE != bSucc) {
            CComPtr<IXMLDOMElement> spXMLRootElem;
            hr = spXMLDoc->get_documentElement(&spXMLRootElem);
            if (FAILED(hr)) { break; }
            CComBSTR name;
            spXMLRootElem->get_nodeName(&name);
            if (!(name == XML_PROFILE_ROOT_ELEM)) {
                bSucc = VARIANT_FALSE;
            }
        }
        if (VARIANT_FALSE == bSucc) {
            hr = spXMLDoc->loadXML((BSTR) XML_PROFILE_ROOT_NODE, &bSucc);
            if (FAILED(hr)){ break; }
        }

        hr = spXMLDoc->QueryInterface(ppXmlDoc); 
        if (FAILED(hr)) {
            break;
        }
    
        hr = S_OK;
    } while (FALSE);
    return hr;
}

EXTERN_C BOOL WINAPI XmlWriteSystemTime(IN IXMLDOMDocument2 * pXmlDoc,
    LPCWSTR lpSection, LPCWSTR lpSubSection, LPCWSTR lpKeyName,
    IN const SYSTEMTIME *BeginTime)
{
    BOOL bRs = FALSE;
    do 
    {
        if ( pXmlDoc == NULL) { break; }
        DOUBLE dTime = 0.0;
        SystemTimeToVariantTime(const_cast<LPSYSTEMTIME>(BeginTime), &dTime);

        CComVariant vTime((DOUBLE)dTime, VT_DATE);
        vTime.ChangeType(VT_BSTR);

        if (FALSE == XmlWriteProfileString(pXmlDoc, lpSection,
            lpSubSection, lpKeyName, vTime.bstrVal))
        {
            break;
        }

        bRs = TRUE;
    } while (FALSE);
    return bRs;
}


EXTERN_C BOOL WINAPI
	XmlWriteProfileString (IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection, 
	LPCWSTR lpSubSection,
	LPCWSTR lpKeyName, 
	LPCWSTR lpString
	)
{
	HRESULT hr = E_FAIL;
	do 
	{
		if (NULL == lpSection || 0 == lstrlen(lpSection)) {
			break;
		}

		// Just to test the doc pointer contain string
		CComBSTR bstrData;
		hr = pXmlDoc->get_xml(&bstrData);
		if (FAILED(hr)) { break; }

		if (S_OK != hr) {
			VARIANT_BOOL bSucc2 = VARIANT_FALSE;
			hr = pXmlDoc->loadXML(CComBSTR( XML_PROFILE_ROOT_NODE ), &bSucc2);
			hr = E_FAIL;
			if (bSucc2 == VARIANT_FALSE) { break; }
		}

		CComPtr<IXMLDOMElement> spXMLRootElem;
		hr = pXmlDoc->get_documentElement(&spXMLRootElem);
		if (FAILED(hr)) { break; }

		CComPtr<IXMLDOMNode> spSection;
		hr = GetOrCreateChildElement(pXmlDoc, spXMLRootElem, lpSection, &spSection);
		if (FAILED(hr)) { break; }

		if (lpSubSection && lstrlen(lpSubSection))
		{
			CComPtr<IXMLDOMNode> spSubSection;
			hr = GetOrCreateChildElement(pXmlDoc, spSection, lpSubSection, &spSubSection);
			if (FAILED(hr)) { break; }

			if (lpKeyName)
			{
				CComPtr<IXMLDOMNode> spKey;
				hr = GetOrCreateChildElement(pXmlDoc, spSubSection, lpKeyName, &spKey);
				if (FAILED(hr)) { break; }

				if (lpString && lstrlenW(lpString))
				{
					CComPtr<IXMLDOMElement> spKeyName;  spKey->QueryInterface(&spKeyName);
					hr = spKeyName->setAttribute(CComBSTR(XML_PROFILE_VALUE), CComVariant(lpString));
				}
			} 
		}
		else
		{
			//////////////////////////////////
			// add option to section
			//////////////////////////////////

			if (lpKeyName)
			{
				CComPtr<IXMLDOMNode> spKey;
				hr = GetOrCreateChildElement(pXmlDoc, spSection, lpKeyName, &spKey);
				if (FAILED(hr)) { break; }

				if (lpString && lstrlenW(lpString))
				{
					CComPtr<IXMLDOMElement> spKeyName; spKey->QueryInterface(&spKeyName);
					hr = spKeyName->setAttribute(CComBSTR(XML_PROFILE_VALUE), CComVariant(lpString));
				}
			} 
		}

		CComPtr<IXMLDOMDocument2> spXmlFormattedDoc;
		do 
		{
			hr = FormatXmlDoc(pXmlDoc, &spXmlFormattedDoc);
			if (FAILED(hr)) { break; }

			//hr = ConvertXmlEncodingUTF8(spXmlFormattedDoc);
			//if (FAILED(hr)) { break; }
		} while (FALSE);

		if (SUCCEEDED(hr)) {
			bstrData.Empty();
			spXmlFormattedDoc->get_xml(&bstrData);
			VARIANT_BOOL bSucc = VARIANT_FALSE;
			pXmlDoc->loadXML(bstrData, &bSucc);
		}
	} while (FALSE);
	return SUCCEEDED(hr) ? TRUE : FALSE;
}

EXTERN_C BOOL WINAPI
	XmlWriteProfileInt (IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection, 
	LPCWSTR lpSubSection,
	LPCWSTR lpKeyName, 
	INT nValue
	)
{
	CComVariant vValue(nValue);
	vValue.ChangeType(VT_BSTR);
	return XmlWriteProfileString(pXmlDoc, lpSection, lpSubSection,
		lpKeyName, vValue.bstrVal);
}


EXTERN_C BOOL WINAPI
	XmlRemoveProfileString (IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection, 
	LPCWSTR lpSubSection,
	LPCWSTR lpKeyName
	)
{
	BOOL bsResult = FALSE;
	HRESULT hr = E_FAIL;
	CComPtr<IXMLDOMDocument2> spXMLDoc;
	ULARGE_INTEGER uliDummy = { 0 };
	LARGE_INTEGER liBegin = { 0 };

	do
	{
		if (NULL == lpSection || 0 == lstrlen(lpSection)) {
			break;
		}

		//Note: We assume that the stream data is Unicode String
		CComBSTR bstrData;

		hr = pXmlDoc->get_xml(&bstrData);
		if( 0 ==bstrData.Length() ) {
			VARIANT_BOOL bSucc2 = VARIANT_FALSE;
			hr = spXMLDoc->loadXML(CComBSTR( XML_PROFILE_ROOT_NODE ), &bSucc2);
			hr = E_FAIL;
			if (bSucc2 == VARIANT_FALSE) { break; }
		}
		bstrData.Empty();

		CComPtr<IXMLDOMElement> spXMLRootElem;
		hr = pXmlDoc->get_documentElement(&spXMLRootElem);
		if (FAILED(hr)){ break; }

		CComPtr<IXMLDOMNode> spSection;
		hr = GetOrCreateChildElement(pXmlDoc, spXMLRootElem, lpSection, &spSection);
		if (FAILED(hr)) { break; }

		// if lpSection is true, lpSubSection is true,
		// lpKeyName is true then delete lpKeyName.

		if ((lpSubSection && lstrlen(lpSubSection)) &&
			(lpKeyName && lstrlen(lpKeyName)) )
		{
			CComPtr<IXMLDOMNode> spSubSection;
			hr = GetOrCreateChildElement(pXmlDoc, spSection, lpSubSection, &spSubSection);

			CComPtr<IXMLDOMNode> spKey;
			hr = GetOrCreateChildElement(pXmlDoc, spSubSection, lpKeyName, &spKey);

			//delete the lpKeyName
			CComPtr<IXMLDOMNode> spFileName2 =spSubSection;
			CComPtr<IXMLDOMNode> spDummy;

			hr = spFileName2->removeChild(spKey, &spDummy);
			bsResult = TRUE;
			break;
		}

		//if lpSection is true, lpSubSection is Null
		//lpKeyName is true then delete the lpkeyName
		if ((!lpSubSection || 0==lstrlen(lpSubSection)) && 
			(lpKeyName && lstrlen(lpKeyName)))
		{
			//delete lpkeyName
			CComPtr<IXMLDOMNode> spKey;
			hr = GetOrCreateChildElement(pXmlDoc, spSection, lpKeyName, &spKey);

			//delete lpKeyName
			CComPtr<IXMLDOMNode> spFileName2 = spSection;
			CComPtr<IXMLDOMNode> spDummy;
			hr = spFileName2->removeChild( spKey, &spDummy);

			bsResult = TRUE;
			break;
		}

		//if lpSection is true, lpSubSection is true,
		//and lpKeyName is Null, then delete the lpSubSection
		if ((lpSubSection && lstrlen(lpSubSection)) &&
			(!lpKeyName || 0==lstrlen(lpKeyName)))
		{
			CComPtr<IXMLDOMNode> spSubSection;
			hr = GetOrCreateChildElement(spXMLDoc, spSection, 
				lpSubSection, &spSubSection);

			//delete the lpSubSection
			CComPtr<IXMLDOMNode> spSection2 = spSection;
			CComPtr<IXMLDOMNode> spDummy;
			hr = spSection2->removeChild(spSubSection, &spDummy);

			bsResult = TRUE;
			break;
		}

		//if lpSection is true, lpSubSection is Null, 
		//and lpKeyName is Null, then delete the lpSection
		if ((!lpSubSection || 0==lstrlen(lpSubSection)) &&
			(!lpKeyName || 0==lstrlen(lpKeyName)))
		{
			// delete the lpSection
			CComPtr<IXMLDOMNode> spRoot = spXMLRootElem;
			CComPtr<IXMLDOMNode> spDummy;
			hr = spRoot->removeChild( spSection, &spDummy);

			bsResult = TRUE;
			break;
		}

	} while (FALSE);

	if (bsResult)
	{
		CComPtr<IXMLDOMDocument2> spXmlFormattedDoc;
		do 
		{
			hr = FormatXmlDoc(pXmlDoc, &spXmlFormattedDoc);
			if (FAILED(hr)) { break; }

			//hr = ConvertXmlEncodingUTF8(spXmlFormattedDoc);
			//if (FAILED(hr)) { break; }
		} while (FALSE);

		if (SUCCEEDED(hr)) {
			CComBSTR bstrData;
			VARIANT_BOOL isSuccessful = VARIANT_FALSE;
			spXmlFormattedDoc->get_xml(&bstrData);
			pXmlDoc->loadXML( bstrData, &isSuccessful );
		}

		//spXMLDoc->save(CComVariant(L"c:\\aaad.xml"));
	}

	return bsResult;
}


EXTERN_C BOOL WINAPI 
	XmlWriteProfileSection(IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection,
	LPCWSTR lpSubSection,
	LPCWSTR lpString
	)
{
	HRESULT hr = E_FAIL;
	TCHAR * pIter = (TCHAR *) lpString;
	int nlen = 0;

	while ( nlen = lstrlen(pIter) ) 
	{
		TCHAR * pEquel = _tcschr(pIter, _T('='));
		if (NULL == pEquel) {
			pEquel = _tcschr(pIter, _T(':'));
		}
		//	CString strKey;
		TCHAR *	strKey=NULL;

		//	CString strValue;
		TCHAR * strValue = NULL;	
		unsigned int address=0;
		if (pEquel) {
			//strKey = CString(pIter).Left(pEquel-pIter);
			address = pEquel - pIter;
			strKey= (TCHAR *)malloc(sizeof(TCHAR) * (address + 1));
			memset( strKey, 0, sizeof(TCHAR) * (address + 1) );
			_tcsncpy( strKey, pIter,  address);

			pEquel++;
			strValue = pEquel;
		} else {
			strKey = pIter;
		}

		XmlWriteProfileString(pXmlDoc, lpSection, lpSubSection, strKey, strValue);

		if (strKey!= pIter)
		{
			free(strKey);
		}

		pIter += (nlen + 1);
	} while (FALSE);

	return SUCCEEDED(hr) ? TRUE : FALSE;
}


//
// load information from pXmlDoc (IXMLDOMDocument2) interface,
// and parse it to Section/subSec/key/value format
//

EXTERN_C DWORD WINAPI
	XmlGetProfileString(IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection,
	LPCWSTR lpSubSection,
	LPCWSTR lpKeyName,
	LPCWSTR lpDefault,
	LPWSTR lpReturnedString,
	DWORD cchMaxReturn
	)
{
	USES_CONVERSION;
	HRESULT hr = E_FAIL;
	CComVariant varValue;
	DWORD nRs = 0;

	do 
	{
		if (NULL == pXmlDoc) {
			break;
		}

		CComBSTR bstrData;
		pXmlDoc->get_xml(&bstrData);
		if ( 0 == bstrData.Length() )
		{
			VARIANT_BOOL bSucc2 = VARIANT_FALSE;
			hr = pXmlDoc->loadXML(XML_PROFILE_ROOT_NODE, &bSucc2);
			hr = E_FAIL;
			if (bSucc2 == VARIANT_FALSE) { break; }
		}

		CComPtr<IXMLDOMElement> spXMLRootElem;
		hr = pXmlDoc->get_documentElement(&spXMLRootElem);
		if (FAILED(hr)) { break; }

		if (NULL == lpSection || 0 == lstrlen(lpSection)) {
			hr = E_FAIL;
			break;
		}

		CComPtr<IXMLDOMNode> spSection;
		hr = GetOrCreateChildElement(pXmlDoc, spXMLRootElem, lpSection, &spSection);
		if (FAILED(hr)) { break; }

		if (lpSubSection && lstrlen(lpSubSection))
		{
			CComPtr<IXMLDOMNode> spSubSection;
			hr = GetOrCreateChildElement(pXmlDoc, spSection, lpSubSection, &spSubSection);
			if (FAILED(hr)) { break; }

			if (NULL == lpKeyName) { break; }
			CComPtr<IXMLDOMNode> spKey;
			hr = GetOrCreateChildElement(pXmlDoc, spSubSection, lpKeyName, &spKey);
			if (FAILED(hr)) { break; }

			CComPtr<IXMLDOMElement> spKeyName; spKey->QueryInterface(&spKeyName);

			hr = spKeyName->getAttribute(CComBSTR(XML_PROFILE_VALUE), &varValue);
		}
		else
		{
			CComPtr<IXMLDOMNode> spKey;
			hr = GetOrCreateChildElement(pXmlDoc, spSection, lpKeyName, &spKey);
			if (FAILED(hr)) { break; }
			CComPtr<IXMLDOMElement> spkeyName; spKey->QueryInterface(&spkeyName);
			hr = spkeyName->getAttribute( CComBSTR(XML_PROFILE_VALUE), &varValue);
		}

	} while (FALSE);

	if (S_OK != hr) {
		if (lpDefault) {
			varValue = lpDefault;
		}
	}

	varValue.ChangeType(VT_BSTR);

	if (lpReturnedString && cchMaxReturn)
	{
		lstrcpyn(lpReturnedString, OLE2CT(varValue.bstrVal), cchMaxReturn);
		nRs = lstrlen(lpReturnedString);
	}

	return nRs;
}

EXTERN_C DWORD WINAPI
	XmlGetProfileInt(IXMLDOMDocument2 * pXmlDoc,
	LPCWSTR lpSection,
	LPCWSTR lpSubSection,
	LPCWSTR lpKeyName,
	INT nDefault
	)
{
	TCHAR szOut[MAX_PATH] = { 0 };
	XmlGetProfileString(pXmlDoc, lpSection, lpSubSection, lpKeyName, 
		_itot(nDefault, szOut, _countof(szOut)), 
		szOut, _countof(szOut));


	TCHAR * stopstring = NULL;
	return (DWORD) _tcstol(szOut, &stopstring, 10);
}

EXTERN_C HRESULT WINAPI
	XmlMergeDocument(IN IXMLDOMDocument2 * pSrcDoc1,
	IN IXMLDOMDocument2 * pSrcDoc2,
	OUT IXMLDOMDocument2 ** ppNewDoc
	)
{
	HRESULT hr = E_FAIL;
	VARIANT_BOOL bSucc = VARIANT_FALSE;

	do 
	{
		if (NULL==pSrcDoc1 || NULL==pSrcDoc2 || NULL==ppNewDoc) {
			break;
		}

		CComPtr<IXMLDOMDocument2> spOutXmlDoc;

		{
			CComPtr<IXMLDOMNode> spXmlNode;
			hr = pSrcDoc1->cloneNode(VARIANT_TRUE, &spXmlNode);
			if (FAILED(hr)) { break; }

			hr = spXmlNode->QueryInterface(&spOutXmlDoc);
			if (FAILED(hr)) { break; }
		}

		// merge operating... 
		{
			long cbFirstLength = 0, cbSecondLength = 0, cbThirdLength = 0;
			INT32 cbfirstSize = 0, cbSecondSize = 0, cbThirdSize = 0;

			CComBSTR lpSection, lpSubSection, lpkeyName, lpString;			

			CComPtr<IXMLDOMNode> RootNode ;
			CComPtr<IXMLDOMNode> FirstLevelNode;
			CComPtr<IXMLDOMNode> SecondLevelNode;
			CComPtr<IXMLDOMNode> ThirdLevelNode;

			{
				// convert "pSrcDoc2" Document to IXMLDomNode interface
				CComPtr<IXMLDOMElement> spElm;
				hr = pSrcDoc2->get_documentElement(&spElm);
				if (S_OK != hr) { break; }
				RootNode = spElm;
			}
			hr = RootNode->get_firstChild(&FirstLevelNode);
			if (S_OK != hr) { break; }
			hr = FirstLevelNode->get_firstChild(&SecondLevelNode);

			while (SecondLevelNode == NULL)
			{
				CComPtr<IXMLDOMNode> FirstLevelNode2;
				CComPtr<IXMLDOMNode> SecondLevelNode2;

				hr =FirstLevelNode->get_nextSibling(&FirstLevelNode2);
				FirstLevelNode = FirstLevelNode2;
				if (FirstLevelNode == NULL)
				{
					break;
				}
				hr = FirstLevelNode->get_firstChild(&SecondLevelNode2);
				SecondLevelNode = SecondLevelNode2;
			}

			if (S_OK != hr) { break; }
			hr = SecondLevelNode->get_firstChild(&ThirdLevelNode);

			CComPtr<IXMLDOMNodeList> spXmlFirstList;
			hr = RootNode->get_childNodes(&spXmlFirstList);

			hr = spXmlFirstList->get_length(&cbFirstLength);


			CComPtr<IXMLDOMNodeList> spXmlSecondList;
			hr = FirstLevelNode->get_childNodes(&spXmlSecondList);

			hr= spXmlSecondList->get_length(&cbSecondLength);


			if (ThirdLevelNode != NULL)
			{
				//cbThirdLength = SecondLevelNode.childNodes.length;
				CComPtr<IXMLDOMNodeList> spXmlList;
				hr = SecondLevelNode->get_childNodes(&spXmlList);
				if (S_OK != hr) { break; }
				hr = spXmlList->get_length(&cbThirdLength);
				if (S_OK != hr) { break; }
			}

			hr = FirstLevelNode->get_baseName(&lpSection);


			// judge lpSubSection is exist ?
			//  if not exist 
			ThirdLevelNode.Release();
			hr = SecondLevelNode->get_firstChild(&ThirdLevelNode);

			if (ThirdLevelNode == NULL)
			{
				lpkeyName.Empty();
				hr = SecondLevelNode->get_baseName(&lpkeyName);
				if (FAILED(hr))	{ break; }

				CComPtr<IXMLDOMNamedNodeMap> spAttributeNodeMap;
				hr = SecondLevelNode->get_attributes(&spAttributeNodeMap);
				if (FAILED(hr))	{ break; }

				CComPtr<IXMLDOMNode> spNode;
				hr = spAttributeNodeMap->get_item(0, &spNode);
				if (FAILED(hr))	{ break; }

				CComBSTR tempComp;
				hr = spNode->get_nodeName(&tempComp);
				if (FAILED(hr)) { break; }

				if (tempComp == XML_PROFILE_VALUE)
				{
					CComVariant vValue;
					hr =spNode->get_nodeValue( &vValue );
					if (FAILED(hr))	{ break; }

					vValue.ChangeType(VT_BSTR);
					lpString = vValue.bstrVal;
				}
			}
			else
			{
				lpSubSection.Empty();
				hr =SecondLevelNode->get_baseName(&lpSubSection);
				if (FAILED(hr))	{ break; }

				CComPtr<IXMLDOMNode> spTempNode;
				hr = SecondLevelNode->get_firstChild( &spTempNode);
				if (FAILED(hr))	{ break; }

				lpkeyName.Empty();
				spTempNode->get_baseName(&lpkeyName);

				CComPtr<IXMLDOMNode> spNode;
				CComPtr<IXMLDOMNamedNodeMap> spMapTemp;
				CComBSTR tempComp;
				hr = ThirdLevelNode->get_attributes(&spMapTemp);
				if (FAILED(hr))	{ break; }
				hr = spMapTemp->get_item(0, &spNode);
				if (FAILED(hr))	{ break; }			
				hr = spNode->get_nodeName(&tempComp);
				if (FAILED(hr))	{ break; }

				if (tempComp == XML_PROFILE_VALUE)
				{
					CComVariant vValue;
					hr = spNode->get_nodeValue(&vValue);
					if (FAILED(hr))	{ break; }			
					vValue.ChangeType(VT_BSTR);
					lpString =vValue.bstrVal;
				}				
			}

			// first "for" loop
			for (cbfirstSize = 0; cbfirstSize < cbFirstLength; cbfirstSize++)
			{
				// second "for" loop 
				for (cbSecondSize = 0; cbSecondSize < cbSecondLength; cbSecondSize++)
				{
					if (0 == cbThirdLength)
					{
						lpSubSection.Empty();
						XmlWriteProfileString(spOutXmlDoc, lpSection, lpSubSection, lpkeyName, lpString);
					}

					// third "for" loop
					for (cbThirdSize = 0; cbThirdSize < cbThirdLength; cbThirdSize++)
					{
						XmlWriteProfileString(spOutXmlDoc, lpSection, lpSubSection, lpkeyName, lpString);

						lpString.Empty();

						CComPtr<IXMLDOMNode> spTemp;
						hr = ThirdLevelNode->get_nextSibling(&spTemp);

						if ( spTemp != NULL )
						{
							spTemp.Release();
							ThirdLevelNode->get_nextSibling(&spTemp);

							lpkeyName.Empty();
							spTemp->get_baseName(&lpkeyName);

							CComPtr<IXMLDOMNode> spNodeTemp;
							CComPtr<IXMLDOMNamedNodeMap> spMapTemp;
							CComBSTR tempComp;

							ThirdLevelNode->get_nextSibling(&spNodeTemp);
							spNodeTemp->get_attributes(&spMapTemp);

							spNodeTemp.Release();
							spMapTemp->get_item(0, &spNodeTemp);
							if (spNodeTemp)
							{
								spNodeTemp->get_nodeName( &tempComp);

								if (tempComp == XML_PROFILE_VALUE)
								{
									CComVariant vValue;
									spNodeTemp->get_nodeValue(&vValue);
									vValue.ChangeType(VT_BSTR);
									lpString = vValue.bstrVal;
								}
							}
							CComPtr<IXMLDOMNode> ThirdLevelNode2;
							ThirdLevelNode->get_nextSibling(&ThirdLevelNode2);
							ThirdLevelNode = ThirdLevelNode2;
						}
					} // third "for" loop 

					do
					{
						// 
						// find out the "second level"'s next sibling
						// 

						CComPtr<IXMLDOMNode> spTemp;
						CComPtr<IXMLDOMNodeList> spListTemp;
						CComPtr<IXMLDOMNamedNodeMap> spMapTemp;

						HRESULT hr00 = SecondLevelNode->get_nextSibling(&spTemp);
						HRESULT hr01 = E_FAIL;
						CComPtr<IXMLDOMNode> spTmp2;
						if (spTemp)
						{
							hr01 = spTemp->get_firstChild(&spTmp2);
							spTemp = spTmp2;
						}

						if( hr00==S_OK && spTemp == NULL)
						{
							CComVariant vValue;
							CComPtr<IXMLDOMNode> SecondLevelNode2;
							SecondLevelNode->get_nextSibling(&SecondLevelNode2);
							SecondLevelNode = SecondLevelNode2;
							SecondLevelNode->get_childNodes(&spListTemp);
							spListTemp->get_length(&cbThirdLength);

							lpkeyName.Empty();
							SecondLevelNode->get_baseName(&lpkeyName);
							SecondLevelNode->get_attributes(&spMapTemp);

							spTemp.Release();
							spMapTemp->get_item(0, &spTemp);

							lpString.Empty();
							if (spTemp)
							{
								spTemp->get_nodeValue(&vValue);
								vValue.ChangeType(VT_BSTR);
								lpString = vValue.bstrVal;
							}

							break;
						}

						spTemp.Release();
						hr00 = SecondLevelNode->get_nextSibling(&spTemp);
						if (spTemp)
						{
							CComPtr<IXMLDOMNode> spTmp3;
							hr01 = spTemp->get_firstChild(&spTmp3);
							spTemp = spTmp3;
						}

						if (hr00==S_OK && spTemp)
						{
							CComVariant vValue;
							CComPtr<IXMLDOMNode> SecondLevelNode2;
							SecondLevelNode->get_nextSibling(&SecondLevelNode2);
							SecondLevelNode = SecondLevelNode2;

							ThirdLevelNode.Release();
							SecondLevelNode->get_firstChild(&ThirdLevelNode);
							SecondLevelNode->get_childNodes(&spListTemp);
							spListTemp->get_length(&cbThirdLength);

							lpSubSection.Empty();
							SecondLevelNode->get_baseName(&lpSubSection);

							lpkeyName.Empty();
							ThirdLevelNode->get_baseName(&lpkeyName);
							ThirdLevelNode->get_attributes(&spMapTemp);

							spTemp.Release();
							spMapTemp->get_item(0, &spTemp);
							spTemp->get_nodeValue(&vValue);
							vValue.ChangeType(VT_BSTR);
							lpString = vValue.bstrVal;
						}
					} while (false);
				} // second for loop

				// 
				// the follow code purse:
				// find out the next sibling node
				// 

				CComPtr<IXMLDOMNode> spTemp;
				CComPtr<IXMLDOMNodeList> spListTemp;
				CComPtr<IXMLDOMNamedNodeMap> spMapTemp;
				hr = FirstLevelNode->get_nextSibling(&spTemp);
				if (spTemp)
				{
					CComPtr<IXMLDOMNode> FirstLevelNode2;
					FirstLevelNode->get_nextSibling(&FirstLevelNode2);
					FirstLevelNode = FirstLevelNode2;

					SecondLevelNode.Release();
					FirstLevelNode->get_firstChild(&SecondLevelNode);

					FirstLevelNode->get_childNodes(&spListTemp);
					spListTemp->get_length(&cbSecondLength);

					SecondLevelNode.Release();
					hr = FirstLevelNode->get_firstChild(&SecondLevelNode);
					if (SecondLevelNode == NULL)
					{
						continue;
						//throw new Exception("Don't allow this kind of node!!!\r\n");
					}

					spTemp.Release();
					SecondLevelNode->get_firstChild(&spTemp);
					if (spTemp == NULL)
					{
						CComVariant vValue;
						cbThirdLength = 0;
						lpSection.Empty();
						FirstLevelNode->get_baseName(&lpSection);

						spTemp.Release();
						FirstLevelNode->get_firstChild(&spTemp);

						lpkeyName.Empty();
						spTemp->get_baseName(&lpkeyName);

						spTemp.Release();
						FirstLevelNode->get_firstChild(&spTemp);
						spTemp->get_attributes(&spMapTemp);

						spTemp.Release();
						spMapTemp->get_item(0, &spTemp);
						spTemp->get_nodeValue(&vValue);
						vValue.ChangeType(VT_BSTR);
						lpString = vValue.bstrVal;
					}
					spTemp.Release();
					SecondLevelNode->get_firstChild(&spTemp);
					if (spTemp != NULL)
					{
						ThirdLevelNode.Release();
						SecondLevelNode->get_firstChild(&ThirdLevelNode);

						lpSection.Empty();
						FirstLevelNode->get_baseName(&lpSection);

						lpSubSection.Empty();
						SecondLevelNode->get_baseName(&lpSubSection);

						lpkeyName.Empty();
						ThirdLevelNode->get_baseName(&lpkeyName);

						ThirdLevelNode->get_attributes(&spMapTemp);

						spTemp.Release();
						spMapTemp->get_item(0, &spTemp);

						CComVariant vValue;
						spTemp->get_nodeValue(&vValue);
						vValue.ChangeType(VT_BSTR);
						lpString = vValue.bstrVal;

						spListTemp.Release();
						SecondLevelNode->get_childNodes(&spListTemp);
						spListTemp->get_length(&cbThirdLength);
					}
				}
			} // first "for" loop 
		}

		hr = spOutXmlDoc->QueryInterface(ppNewDoc); 
	} while (FALSE);
	return hr;
}


HRESULT GetOrCreateChildElement(IXMLDOMDocument2 * pDoc, 
	IXMLDOMNode * pParent, 
	LPCWSTR lpszName, 
	IXMLDOMNode ** ppOut)
{
	HRESULT hr = E_FAIL;
	do 
	{
		CComPtr<IXMLDOMElement> spParent;
		pParent->QueryInterface(&spParent);
		if (spParent == NULL) { break; }

		CComPtr<IXMLDOMNode> spChildNode;
		hr = spParent->selectSingleNode(CComBSTR(lpszName), &spChildNode);
		//if (FAILED(hr)) { break; }

		if (spChildNode==NULL) {
			hr = CreateChildElement(pDoc, pParent, lpszName, &spChildNode);
		}

		if (SUCCEEDED(hr)) {
			hr = spChildNode->QueryInterface(ppOut);
		}
	} while (FALSE);
	return hr;
}

HRESULT ClearOldAndCreateNewChildElement(IXMLDOMDocument2 * pDoc, 
	IXMLDOMNode * pParent, 
	LPCWSTR lpszName, 
	IXMLDOMNode ** ppOut)
{
	HRESULT hr = E_FAIL;
	do 
	{
		CComPtr<IXMLDOMElement> spParent; pParent->QueryInterface(&spParent);
		if (spParent == NULL) { break; }

		CComPtr<IXMLDOMNode> spChildNode;
		hr = spParent->selectSingleNode(CComBSTR(lpszName), &spChildNode);
		if (spChildNode)
		{
			CComPtr<IXMLDOMNode> spDummy;
			spParent->removeChild(spChildNode, &spDummy);
			spChildNode.Release();
		}

		hr = CreateChildElement(pDoc, pParent, lpszName, &spChildNode);
		if (FAILED(hr)) { ATLASSERT(FALSE); break; }

		hr = spChildNode->QueryInterface(ppOut);
	} while (FALSE);
	return hr;
}


HRESULT CreateChildElement(IXMLDOMDocument2 * pDoc, 
	IXMLDOMNode * pParent, 
	LPCWSTR lpszName, 
	IXMLDOMNode ** ppOut)
{
	HRESULT hr = E_FAIL;
	CComPtr<IXMLDOMElement> spTmp;
	hr = pDoc->createElement(CComBSTR(lpszName), &spTmp);
	if (SUCCEEDED(hr)) {
		CComPtr<IXMLDOMElement> spParent; pParent->QueryInterface(&spParent);
		if (spParent) { 
			CComPtr<IXMLDOMNode> spChildNode;
			hr = spParent->appendChild(spTmp, &spChildNode);
			if (SUCCEEDED(hr)) {
				hr = spChildNode->QueryInterface(ppOut);
			}
		}
	}
	return hr;
}

HRESULT WINAPI FormatXmlDoc(IXMLDOMDocument2 * pXmlDoc, 
	IXMLDOMDocument2 ** ppFormattedDoc)
{
	static WCHAR g_szStyle[] = {
		L"<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" 
		L"<xsl:stylesheet xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\" version=\"1.0\">\r\n" 
		L"  <xsl:output method=\"xml\" indent=\"yes\"/>\r\n" 
		L"  <xsl:template match=\"@* | node()\">\r\n" 
		L"      <xsl:copy>\r\n" 
		L"          <xsl:apply-templates select=\"@* | node()\"/>\r\n" 
		L"      </xsl:copy>\r\n" 
		L"  </xsl:template>\r\n" 
		L"</xsl:stylesheet>\r\n" 
	};

	HRESULT hr = E_FAIL;
	CComPtr<IXMLDOMDocument2> spXmlFormattedDoc;
	do 
	{
		if (pXmlDoc == NULL || ppFormattedDoc == NULL) {
			hr = E_INVALIDARG;
			break;
		}

		// Format the XML. This requires a style sheet
		CComPtr<IXMLDOMDocument2> spXmlStyle;
		hr = spXmlStyle.CoCreateInstance(__uuidof(MS_XML_DOMDOC_CLASS));
		if (FAILED(hr)) { break; }

		// We need to load the style sheet which will be used to indent the XMl properly.
		VARIANT_BOOL bSucc2 = VARIANT_FALSE;
		hr = spXmlStyle->loadXML(CComBSTR(g_szStyle), &bSucc2);
		if (FAILED(hr)) { break; }

		// Create the final document which will be indented properly
		hr = spXmlFormattedDoc.CoCreateInstance(__uuidof(MS_XML_DOMDOC_CLASS));
		if (FAILED(hr)) { break; }

		// Apply the transformation to format the final document
		{
			CComPtr<IDispatch> pDispatch(spXmlFormattedDoc);
			CComVariant vtOutObject(pDispatch);
			hr = pXmlDoc->transformNodeToObject(spXmlStyle, vtOutObject);
			if (FAILED(hr)) { break; }

			*ppFormattedDoc = spXmlFormattedDoc.Detach();
		}

	} while (FALSE);
	return hr;
}

EXTERN_C HRESULT WINAPI ConvertXmlEncodingUTF8(IXMLDOMDocument2 * spXmlDoc)
{
	HRESULT hr = E_FAIL;
	do 
	{
		//By default it is writing the encoding = UTF-16. Let us change the encoding to UTF-8
		CComPtr<IXMLDOMNode> spXMLFirstChild;

		// <?xml version="1.0" encoding="UTF-8"?>
		hr = spXmlDoc->get_firstChild(&spXMLFirstChild);
		if (FAILED(hr)) { break; }

		CComPtr<IXMLDOMNamedNodeMap> spXMLAttributeMap;
		// A map of the a attributes (version, encoding) values (1.0, UTF-8) pair
		hr = spXMLFirstChild->get_attributes(&spXMLAttributeMap);
		if (FAILED(hr)) { break; }

		CComPtr<IXMLDOMNode> spXMLEncodNode;
		hr = spXMLAttributeMap->getNamedItem(CComBSTR(L"encoding"), &spXMLEncodNode);
		if (FAILED(hr)) { break; }

		//encoding = UTF-8
		hr = spXMLEncodNode->put_nodeValue(CComVariant(L"UTF-8"));

	} while (FALSE);

	return hr;
}

/*
// =====================================================
// Get the Length of a IStream interface pointer
// =====================================================
BOOL GetStreamLength (IStream* pStream, ULARGE_INTEGER* puiLength)
{
	BOOL  bSuccess = FALSE;

	LARGE_INTEGER lMov = { 0 };
	ULARGE_INTEGER ulEnd,ulBegin;

	HRESULT hr = S_FALSE;

	do
	{
		if (NULL==pStream || NULL==puiLength) {
			break;
		}

		CComPtr<IStream> spStreamClone;
		hr = pStream->Clone (&spStreamClone);

		if (FAILED(hr)) {
			break;
		}

		hr  =  spStreamClone->Seek (lMov, STREAM_SEEK_END, &ulEnd);
		if (FAILED(hr))
			break;

		hr  =  spStreamClone->Seek (lMov, STREAM_SEEK_SET, &ulBegin);
		if (FAILED(hr))
			break;

		puiLength->QuadPart  =  ulEnd.QuadPart - ulBegin.QuadPart;

		bSuccess   =  TRUE;
	} while(FALSE);
	return bSuccess;
}


// =====================================================
// Create a steam in memory, 
// the final release will automatically free the hGlobal 
// =====================================================

BOOL CreateCompatileStream (SIZE_T dwSize, IStream** ppStream)
{
	BOOL  bSuccess  =  FALSE;
	HGLOBAL hGlobal = NULL;

	if (dwSize) {
		hGlobal = ::GlobalAlloc (GPTR, dwSize);
	}

	//Note: fDeleteOnRelease must be TRUE
	bSuccess = SUCCEEDED(::CreateStreamOnHGlobal(NULL, TRUE, ppStream));

	return bSuccess;
}

EXTERN_C BOOL WINAPI UnzipXmlToStream(IN const LPBYTE pZippedInfo, IN UINT cbSize,
	OUT IStream ** ppStrm)
{
	BOOL bResult = FALSE;
	BYTE * pUnZip = NULL;
	ULONG nUnzipLen = 0;

	do 
	{
		if (NULL==pZippedInfo || 0==cbSize || NULL==ppStrm) { 
			break;
		}

		nUnzipLen = (cbSize + 1) * 2;
		pUnZip = (BYTE *) malloc( nUnzipLen );
		if (NULL == pUnZip) { break; }

		bResult = ZipDecompressInfo(pZippedInfo, cbSize, 
			pUnZip, (unsigned int *)&nUnzipLen);
		if (FALSE == bResult) { break; }

		bResult = CreateCompatileStream(nUnzipLen, ppStrm);
		if (FALSE == bResult) { break; }

		if (FAILED((*ppStrm)->Write(pUnZip, nUnzipLen, &nUnzipLen))) {
			(*ppStrm)->Release();
			break;
		}

		bResult = TRUE;
	} while (FALSE);

	if (pUnZip) {
		free(pUnZip);
	}

	return bResult;
}

EXTERN_C BOOL WINAPI StreamToZippedXml(IN IStream * pStrm, 
	LPBYTE pZippedInfo, 
	IN OUT UINT * pcbSize)
{
	BOOL bResult = FALSE;

	unsigned char * pXmlData = NULL;
	ULONG pcbRead =0;

	HRESULT hr = E_FAIL;
	ULARGE_INTEGER streamLength= { 0 };

	do 
	{
		if (NULL==pStrm || NULL==pZippedInfo || 
			NULL==pcbSize || 0==(*pcbSize))
		{
			break;
		}

		bResult = GetStreamLength(pStrm, &streamLength);

		pXmlData = (unsigned char*) malloc(streamLength.QuadPart+4);
		memset(pXmlData, 0, streamLength.QuadPart+4);

		ULARGE_INTEGER uliDummy = { 0 };
		LARGE_INTEGER liBegin = { 0 };
		hr = pStrm->Seek (liBegin, STREAM_SEEK_SET, &uliDummy);
		if (FAILED(hr)) { break; }

		hr = pStrm->Read( pXmlData, streamLength.QuadPart, &pcbRead );
		if (FAILED(hr)) { break; }

		bResult = ZipCompressInfo((const unsigned char *)pXmlData, pcbRead, 
			pZippedInfo, pcbSize);
		if (FALSE == bResult) { break; }

		bResult =TRUE;
	} while (FALSE);

	// must free memory
	if (pXmlData) {
		free(pXmlData);
	}

	return bResult;
}
//*/

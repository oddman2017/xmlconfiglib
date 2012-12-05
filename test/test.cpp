// test.cpp : Defines the entry point for the console application.
//
#include <Windows.h>

#include <atlbase.h>

#include <msxml2.h>

#include "../xmlprofilestring.h"

void test(void)
{
	HRESULT hr = E_FAIL;

#define XML_FILE L"test.xml"
	
	CComPtr<IXMLDOMDocument2> spDoc;
	hr = XmlCreateProfileDoc(XML_FILE, &spDoc);
	if (SUCCEEDED(hr)) {
		XmlWriteProfileString(spDoc, L"made", L"gouride", L"key1", L"test string");
		ConvertXmlEncodingUTF8(spDoc);
		spDoc->save(CComVariant(XML_FILE));
	}
}


int main(int argc, char* argv[])
{

	CoInitialize(NULL);
	
	test();

	CoUninitialize();
	return 0;
}

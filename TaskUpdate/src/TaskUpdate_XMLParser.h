#pragma once

class CTU_XMLParser
{
public:
    CTU_XMLParser();
    virtual ~CTU_XMLParser();

public:
    HRESULT GetNodeValue(IXMLDOMNode* spXMLNode, CComBSTR nodeName, CString& nodeVal);
    HRESULT GetAttributeValue(IXMLDOMNode* spXMLNode, CComBSTR attributeName, CString& attrVal);

    HRESULT AddProcessingInstructionNode(IXMLDOMDocument* pDom, CComBSTR bstrTarget, CComBSTR bstrData);
    HRESULT AddRootElementNode(IXMLDOMDocument* pDom, CComBSTR bstrName, IXMLDOMNode** ppNodeOut);
    HRESULT AddChildElementNode(IXMLDOMDocument* pDom, IXMLDOMNode* pParent, CComBSTR bstrName, IXMLDOMNode** ppNodeOut);
    HRESULT AddAtributeToNode(IXMLDOMDocument* pDom, IXMLDOMNode* pNode, CComBSTR bstrName, CComBSTR bstrValue);
    HRESULT AddTextNode(IXMLDOMDocument* pDom, IXMLDOMNode* pParent, CComBSTR bstrNodeName, CString& strNodeValue, IXMLDOMNode** ppNodeOut);
};


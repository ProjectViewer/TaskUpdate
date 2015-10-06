#include "TaskUpdate_XMLParser.h"

CTU_XMLParser::CTU_XMLParser()
{
}
CTU_XMLParser::~CTU_XMLParser()
{
}

HRESULT CTU_XMLParser::GetNodeValue(IXMLDOMNode* spXMLNode, CComBSTR nodeName, CString& nodeVal)
{
    CComPtr<IXMLDOMNode> spXMLChildNode;
    CComVariant varValue(VT_EMPTY);
    HRESULT hr = spXMLNode->selectSingleNode(nodeName, &spXMLChildNode);
    if (hr != S_OK)
        return hr;

    hr = spXMLChildNode ? spXMLChildNode->get_nodeTypedValue(&varValue) : S_FALSE;
    if (hr != S_OK)
        return hr;

    nodeVal = varValue.bstrVal;
    nodeVal.Replace(_T("&#x0Dx0A"), _T("\r\n"));
    return hr;
}
HRESULT CTU_XMLParser::GetAttributeValue(IXMLDOMNode* spXMLNode, CComBSTR attributeName, CString& attrVal)
{
    CComQIPtr<IXMLDOMElement> spXMLChildElement;
    spXMLChildElement = spXMLNode;
    if (spXMLChildElement.p == NULL)
        return S_FALSE;

    CComVariant varValue(VT_EMPTY);
    HRESULT hr = spXMLChildElement->getAttribute(attributeName, &varValue);
    if (hr != S_OK)
        return hr;

    attrVal = varValue.bstrVal;
    return hr;
}
//---------------------------------------------------------------------------------------------------------------------
HRESULT CTU_XMLParser::AddTextNode(IXMLDOMDocument* pDom, IXMLDOMNode* pParent, CComBSTR bstrNodeName, CString& strNodeValue, IXMLDOMNode** ppNodeOut)
{
    CComPtr<IXMLDOMElement> spElement;
    CComPtr<IXMLDOMText> spTempNodeOut1;
    CComPtr<IXMLDOMNode> spTempNodeOut2;

    HRESULT hr = pDom->createElement(bstrNodeName, &spElement);
    if (hr != S_OK)
        return hr;

    strNodeValue.Replace(_T("\r\n"), _T("&#x0Dx0A"));
    hr = spElement->put_text(CComBSTR(strNodeValue));
    if (hr != S_OK)
        return hr;
    hr = pParent->appendChild(spElement, ppNodeOut);
    return hr;
}

HRESULT CTU_XMLParser::AddRootElementNode(IXMLDOMDocument* pDom, CComBSTR bstrName, IXMLDOMNode** ppNodeOut)
{
    CComPtr<IXMLDOMElement> spRoot;
    HRESULT hr = pDom->createElement(bstrName, &spRoot);
    if (hr != S_OK)
        return hr;
    hr = pDom->appendChild(spRoot, ppNodeOut);
    return hr;
}
HRESULT CTU_XMLParser::AddChildElementNode(IXMLDOMDocument* pDom, IXMLDOMNode* pParent, CComBSTR bstrName, IXMLDOMNode** ppNodeOut)
{
    CComPtr<IXMLDOMElement> spElement;
    HRESULT hr = pDom->createElement(bstrName, &spElement);
    if (hr != S_OK)
        return hr;
    hr = pParent->appendChild(spElement, ppNodeOut);
    return hr;
}
HRESULT CTU_XMLParser::AddAtributeToNode(IXMLDOMDocument* pDom, IXMLDOMNode* pNode, CComBSTR bstrName, CComBSTR bstrValue)
{
    CComPtr<IXMLDOMAttribute> spAttr;
    CComPtr<IXMLDOMAttribute> spAttrOut;
    HRESULT hr = pDom->createAttribute(bstrName, &spAttr);
    if (hr != S_OK)
        return hr;
    hr = spAttr->put_value(CComVariant(bstrValue));
    if (hr != S_OK)
        return hr;
    CComPtr<IXMLDOMNode> spTest;
    CComPtr<IXMLDOMElement> spTest2;
    spTest = pNode;
    spTest2 = spTest;
    hr = spTest2->setAttributeNode(spAttr, &spAttrOut);
    return hr;
}

HRESULT CTU_XMLParser::AddProcessingInstructionNode(IXMLDOMDocument* pDom, CComBSTR bstrTarget, CComBSTR bstrData)
{
    CComPtr<IXMLDOMProcessingInstruction> spPi;
    CComPtr<IXMLDOMNode> spPiOut;

    HRESULT hr = pDom->createProcessingInstruction(bstrTarget, bstrData, &spPi);
    if (hr != S_OK)
        return hr;

    hr = pDom->appendChild(spPi, &spPiOut);
    return hr;
}


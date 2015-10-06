#include "DocCustomProperties.h"
#include "PropertyStorage.h"
#include <sstream>

CDocCustomProperties::CDocCustomProperties(void)
{
}

CDocCustomProperties::~CDocCustomProperties(void)
{
}

BOOL CDocCustomProperties::LoadCustomProperties(const CString path)
{
    BOOL retVal = TRUE;
    CComPtr<IStorage> storage;

    HRESULT hr = ::StgOpenStorage(static_cast<CComBSTR>(path), NULL, STGM_DIRECT_SWMR | STGM_READ | STGM_SHARE_DENY_NONE, NULL, 0, &storage);
    if FAILED(hr)
    {
        retVal = FALSE;
    }
    else
    {
        m_PropertySetStorage = storage;

        if (!m_PropertySetStorage)
            retVal = FALSE;

        retVal = GetCustomDocProperties();
    }
    return retVal;
}

BOOL CDocCustomProperties::GetCustomDocProperties()
{
    CComPtr<IPropertyStorage> propertyStorage;
    CComPtr<IEnumSTATPROPSTG> enumPropertyStorage;

    ATLASSERT(m_PropertySetStorage);

    if (!m_PropertySetStorage)
        return FALSE;

    HRESULT hr = m_PropertySetStorage->Open(FMTID_UserDefinedProperties, STGM_READ | STGM_SHARE_EXCLUSIVE, &propertyStorage);

    if (FAILED(hr) || !m_PropertySetStorage)
        return FALSE;

    hr = propertyStorage->Enum(&enumPropertyStorage);
    if (FAILED(hr) || !enumPropertyStorage)
        return FALSE;
    // Enumerate properties.
    STATPROPSTG sps;
    ULONG fetched;
    PROPSPEC propSpec[1];
    PROPVARIANT propVar[1];
    while (enumPropertyStorage->Next(1, &sps, &fetched) == S_OK)
    {
        // Build a PROPSPEC for this property.
        ZeroMemory(&propSpec[0], sizeof(PROPSPEC));
        propSpec[0].ulKind = PRSPEC_PROPID;
        propSpec[0].propid = sps.propid;

        // Read this property.

        hr = propertyStorage->ReadMultiple(1, &propSpec[0], &propVar[0]);
        if (!FAILED(hr))
        {
            if (!CString(sps.lpwstrName).IsEmpty())
            {
                CustomProps m_tempCustomProp;
                m_tempCustomProp.name = sps.lpwstrName;
                m_tempCustomProp.value = GetPropVariant(&propVar[0]);
                m_docsCustomProperties.push_back(m_tempCustomProp);
            }
        }
    }

    return TRUE;
}

CString CDocCustomProperties::GetPropVariant(PROPVARIANT* pPropVar)
{
    CString retVal;

    // Don't iterate arrays, just inform as an array.
    if (pPropVar->vt & VT_ARRAY)
    {
        retVal = _T("(Array)");
        return retVal;
    }

    // Don't handle byref for simplicity, just inform byref.
    if (pPropVar->vt & VT_BYREF)
    {
        retVal = _T("(ByRef)");
        return retVal;
    }

    // Switch types.
    switch (pPropVar->vt)
    {
    case VT_EMPTY:
        retVal = _T("");
        break;
    case VT_NULL:
        retVal = _T("(VT_NULL)");
        break;
    case VT_BLOB:
        retVal = _T("(VT_BLOB)");
        break;
    case VT_BOOL:
        retVal.Format(_T("%s"), pPropVar->boolVal ? "TRUE/YES" : "FALSE/NO");
        break;
    case VT_I2: // 2-byte signed int.
        retVal.Format(_T("%d"), (int)pPropVar->iVal);
        break;
    case VT_I4: // 4-byte signed int.
        retVal.Format(_T("%d"), (int)pPropVar->lVal);
        break;
    case VT_R4: // 4-byte real.
        retVal.Format(_T("%.2lf"), (double)pPropVar->fltVal);
        break;
    case VT_R8: // 8-byte real.
        retVal.Format(_T("%.2lf"), (double)pPropVar->dblVal);
        break;
    case VT_BSTR: // OLE Automation string.
        retVal.Format(_T("%s"), static_cast<CString>(static_cast<CComBSTR>(pPropVar->bstrVal)));
        break;
    case VT_LPSTR: // Null-terminated string.
    {
        CString sTempString = CString(pPropVar->pszVal);
        retVal = sTempString;
    }
    break;
    case VT_FILETIME:
    {
        wchar_t* dayPre[] = {_T("Sun"), _T("Mon"), _T("Tue"), _T("Wed"), _T("Thu"), _T("Fri"), _T("Sat")};

        FILETIME lft;
        FileTimeToLocalFileTime(&pPropVar->filetime, &lft);

        SYSTEMTIME lst;
        FileTimeToSystemTime(&lft, &lst);
        CTime tTheTime = (CTime)lst;
        CString sTheTime = tTheTime.Format(_T("%a, %d %B, %Y, %H:%M:%S"));
        //For edit time
        CString sEditTime = tTheTime.Format(_T("%M"));
        //////////////////////////////////////////////////////////////////////////
        retVal = sTheTime;
    }
    break;
    case VT_CF: // Clipboard format.
        retVal = _T("(Clipboard format)\n");

        break;
    default: // Unhandled type, consult wtypes.h's VARENUM structure.
        retVal.Format(_T("(Unhandled type: 0x%08lx)"), pPropVar->vt);
        break;
    }

    return retVal;
}

CString CDocCustomProperties::GetCustomProperertyValueFor(CString strName)
{
    std::vector<CustomProps>::iterator itCustomProp;
    for (itCustomProp = m_docsCustomProperties.begin();
         itCustomProp != m_docsCustomProperties.end();
         itCustomProp++)
    {
        if (itCustomProp->name == strName)
            return itCustomProp->value;
    }
    return _T("");
}

BOOL CDocCustomProperties::SetCustomDocProperties(const CString path, bool isCollaborative, CString nGuid, int reportingMode)
{
    std::wstring temp = path;
    PropertyStorage propetyStorage(temp);
    const std::string tempGuid = propetyStorage.narrow(nGuid.GetBuffer());
    std::stringstream strMode;
    strMode << reportingMode;
    propetyStorage.setString(L"Collaboration ID", tempGuid);
    propetyStorage.setString(L"Reporting Mode", strMode.str());
    if (isCollaborative)
    {
        propetyStorage.setString(L"IsTUActive", "1");
    }
    else
    {
        propetyStorage.setString(L"IsTUActive", "0");
    }
    return true;
}

bool CDocCustomProperties::SetCustomDocPropertiesMapping(const CString path,  std::vector<std::string>& valueList)
{
    if (valueList.size() == 21)
    {
        std::wstring temp = path;
        PropertyStorage propetyStorage(temp);

        propetyStorage.setString(L"PercComplReport", valueList[0]);
        propetyStorage.setString(L"PercComplDiff", valueList[1]);
        propetyStorage.setString(L"PercComplGraph", valueList[2]);
        propetyStorage.setString(L"PercWorkComplReport", valueList[3]);
        propetyStorage.setString(L"PercWorkComplDiff", valueList[4]);
        propetyStorage.setString(L"PercWorkComplGraph", valueList[5]);
        propetyStorage.setString(L"ActWorkReport", valueList[6]);
        propetyStorage.setString(L"ActWorkDiff", valueList[7]);
        propetyStorage.setString(L"ActWorkGraph", valueList[8]);
		propetyStorage.setString(L"ActStartReport", valueList[9]);
		propetyStorage.setString(L"ActStartDiff", valueList[10]);
		propetyStorage.setString(L"ActStartGraph", valueList[11]);
		propetyStorage.setString(L"ActFinishReport", valueList[12]);
		propetyStorage.setString(L"ActFinishDiff", valueList[13]);
		propetyStorage.setString(L"ActFinishGraph", valueList[14]);
        propetyStorage.setString(L"ActDurationReport", valueList[15]);
        propetyStorage.setString(L"ActDurationDiff", valueList[16]);
        propetyStorage.setString(L"ActDurationGraph", valueList[17]);
        propetyStorage.setString(L"CustomText1", valueList[18]);
        propetyStorage.setString(L"CustomText2", valueList[19]);
        propetyStorage.setString(L"CustomText3", valueList[20]);
    }

    return true;
}

void CDocCustomProperties::SetCustomDocProperty(const CString path, const std::wstring strKey, std::string strVal)
{
    std::wstring temp = path;
    PropertyStorage propetyStorage(temp);
    propetyStorage.setString(strKey, strVal);
}
#pragma once
#include <vector>

struct CustomProps
{
    CString name;
    CString value;
};
class CDocCustomProperties
{
public:
    CDocCustomProperties(void);
    ~CDocCustomProperties(void);

private:
    std::vector<CustomProps> m_docsCustomProperties;
    CComQIPtr<IPropertySetStorage> m_PropertySetStorage;

public:
    BOOL LoadCustomProperties(const CString path);
    BOOL GetCustomDocProperties();
    CString GetPropVariant(PROPVARIANT* pPropVar);
    CString GetCustomProperertyValueFor(CString strName);
    BOOL SetCustomDocProperties(const CString path,  bool isCollaborative, CString nGuid, int reportingMode);
    bool SetCustomDocPropertiesMapping(const CString path,  std::vector<std::string>& valueList);
    void SetCustomDocProperty(const CString path, const std::wstring strKey, std::string strVal);
};

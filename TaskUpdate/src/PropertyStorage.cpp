#include "PropertyStorage.h"
#include <locale>
#include <iostream>
#include <string>
#include <sstream>

class PropertyStorage::Property
{
public:
    Property(const PropertyStorage* storage, const std::wstring& name);
    ~Property();

    bool setValue(const std::string& value);
    bool setValue(int value);

    std::string toString() const;
    int toInt() const;

    void reload();
private:
    bool setValue(PROPVARIANT& value);

    const PropertyStorage* m_storage;
    PROPSPEC m_property;
    PROPVARIANT m_value;
};

PropertyStorage::Property::Property(const PropertyStorage* storage, const std::wstring& name)
    : m_storage(storage)
{
    m_property.ulKind = PRSPEC_LPWSTR;
    m_property.lpwstr = const_cast<wchar_t*>(name.c_str());
    m_storage->m_pPropStg->ReadMultiple(1, &m_property, &m_value);
}

PropertyStorage::Property::~Property()
{
    PropVariantClear(&m_value);
}

std::string PropertyStorage::Property::toString() const
{
    std::string returnValue;

    if (m_value.vt == VT_LPSTR)
    {
        returnValue = m_value.pszVal;
    }

    return returnValue;
}

int PropertyStorage::Property::toInt() const
{
    int returnValue = 0;

    if (m_value.vt == VT_I4)
    {
        returnValue = m_value.intVal;
    }

    return returnValue;
}

bool PropertyStorage::Property::setValue(const std::string& value)
{
    PROPVARIANT propValue;
    propValue.vt = VT_LPSTR;
    propValue.pszVal = const_cast<char*>(value.c_str());

    return setValue(propValue);
}

bool PropertyStorage::Property::setValue(int value)
{
    PROPVARIANT propValue;
    propValue.vt = VT_I4;
    propValue.intVal = value;

    return setValue(propValue);
}

bool PropertyStorage::Property::setValue(PROPVARIANT& value)
{
    bool returnValue = false;

    if (returnValue = SUCCEEDED(m_storage->m_pPropStg->WriteMultiple(1, &m_property, &value, PID_FIRST_USABLE)))
    {
        reload();
    }

    return returnValue;
}

void PropertyStorage::Property::reload()
{
    PropVariantClear(&m_value);
    m_storage->m_pPropStg->ReadMultiple(1, &m_property, &m_value);
}

PropertyStorage::PropertyStorage()
    : m_pPropSetStg(0)
    , m_pPropStg(0)
    , m_bLoaded(false)
{
}

PropertyStorage::PropertyStorage(const std::wstring& file)
    : m_pPropSetStg(0)
    , m_pPropStg(0)
    , m_bLoaded(false)
{
    open(file);
}

PropertyStorage::~PropertyStorage(void)
{
    if (m_bLoaded)
    {
        close();
    }
}

bool PropertyStorage::open(const std::wstring& file)
{
    HRESULT hr = StgOpenStorageEx(file.c_str(),
                                  STGM_DIRECT_SWMR | STGM_READWRITE | STGM_SHARE_DENY_WRITE,
                                  STGFMT_ANY,
                                  NULL,
                                  NULL,
                                  NULL,
                                  IID_IPropertySetStorage,
                                  reinterpret_cast<void**>(&m_pPropSetStg));

    if (SUCCEEDED(hr))
    {
        hr = m_pPropSetStg->Create(FMTID_UserDefinedProperties, NULL, PROPSETFLAG_DEFAULT,
                                   STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, &m_pPropStg);
    }

    return (m_bLoaded = (hr == S_OK));
}

void PropertyStorage::close()
{
    if (m_pPropStg)
    {
        m_pPropStg->Release();
    }

    if (m_pPropSetStg)
    {
        m_pPropSetStg->Release();
    }
}

std::string PropertyStorage::readString(const std::wstring& key) const
{
    return Property(this, key).toString();
}

int PropertyStorage::readInt(const std::wstring& key) const
{
    return Property(this, key).toInt();
}

bool PropertyStorage::setString(const std::wstring& key, const std::string& value)
{
    return Property(this, key).setValue(value);
}

bool PropertyStorage::setInt(const std::wstring& key, int value)
{
    return Property(this, key).setValue(value);
}

std::string PropertyStorage::narrow(const std::wstring& str)
{
    std::ostringstream stm;
    const std::ctype<char>& ctfacet = std::use_facet< std::ctype<char> >(stm.getloc());
    for (size_t i = 0; i < str.size(); ++i)
        stm << ctfacet.narrow(static_cast<char>(str[i]), 0);
    return stm.str();
}
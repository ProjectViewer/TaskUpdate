#pragma once

class PropertyStorage
{
public:
	class Property;

	PropertyStorage();
	PropertyStorage(const std::wstring& file);
	~PropertyStorage(void);

	bool open(const std::wstring& file);
	void close();

	std::string readString(const std::wstring& key) const;
	int readInt(const std::wstring& key) const;

	bool setString(const std::wstring& key, const std::string& value);
    std::string narrow( const std::wstring& str );
	bool setInt(const std::wstring& key, int value);

private:
	IPropertySetStorage* m_pPropSetStg;
	IPropertyStorage* m_pPropStg;
	bool m_bLoaded;
};


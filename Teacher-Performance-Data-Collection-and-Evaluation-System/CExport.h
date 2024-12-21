// CExport.h
#pragma once
#include <vector>
#include <string>

class CExport
{
public:
    static void ExportToExcel(const std::vector<std::vector<std::wstring>>& data, const std::wstring& fileName);
    static void ExportToWord(const std::vector<std::vector<std::wstring>>& data, const std::wstring& fileName);
    static void ExportDataToFile(const std::vector<std::vector<std::wstring>>& data, const std::wstring& format);
};

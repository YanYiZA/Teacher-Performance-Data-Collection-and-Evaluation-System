// CExport.cpp
#include "pch.h"
#include "CExport.h"
#include <fstream>
#include <sstream>
#include <stdio.h>

void CExport::ExportToExcel(const std::vector<std::vector<std::wstring>>& data, const std::wstring& fileName)
{
    // 启动 Python 脚本导出数据到 Excel 文件
    std::wstring command = L"python export_data.py excel ";
    FILE* pipe = _wpopen(command.c_str(), L"w");

    if (pipe != nullptr)
    {
        for (const auto& row : data)
        {
            for (const auto& cell : row)
            {
                fprintf(pipe, "%s,", std::string(cell.begin(), cell.end()).c_str());
            }
            fprintf(pipe, "\n");
        }
        fclose(pipe);
    }
    else
    {
        AfxMessageBox(L"无法启动 Python 脚本！");
    }
}

void CExport::ExportToWord(const std::vector<std::vector<std::wstring>>& data, const std::wstring& fileName)
{
    // 启动 Python 脚本导出数据到 Word 文件
    std::wstring command = L"python export_data.py word ";
    FILE* pipe = _wpopen(command.c_str(), L"w");

    if (pipe != nullptr)
    {
        for (const auto& row : data)
        {
            for (const auto& cell : row)
            {
                fprintf(pipe, "%s,", std::string(cell.begin(), cell.end()).c_str());
            }
            fprintf(pipe, "\n");
        }
        fclose(pipe);
    }
    else
    {
        AfxMessageBox(L"无法启动 Python 脚本！");
    }
}

void CExport::ExportDataToFile(const std::vector<std::vector<std::wstring>>& data, const std::wstring& format)
{
    if (format == L"Excel")
    {
        ExportToExcel(data, L"output.xlsx");
    }
    else if (format == L"Word")
    {
        ExportToWord(data, L"output.docx");
    }
    else
    {
        AfxMessageBox(L"未知的文件格式！");
    }
}

// CExport.cpp
#include "pch.h"
#include "CExport.h"
#include <fstream>
#include <sstream>
#include <stdio.h>

void CExport::ExportToExcel(const std::vector<std::vector<std::wstring>>& data, const std::wstring& fileName)
{
    // ���� Python �ű��������ݵ� Excel �ļ�
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
        AfxMessageBox(L"�޷����� Python �ű���");
    }
}

void CExport::ExportToWord(const std::vector<std::vector<std::wstring>>& data, const std::wstring& fileName)
{
    // ���� Python �ű��������ݵ� Word �ļ�
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
        AfxMessageBox(L"�޷����� Python �ű���");
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
        AfxMessageBox(L"δ֪���ļ���ʽ��");
    }
}

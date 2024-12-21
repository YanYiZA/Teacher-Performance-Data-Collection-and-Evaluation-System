#include "pch.h"
#include "CExport.h"
#include <afx.h>
#include <afxdlgs.h>  // 引入 MFC 文件对话框相关头文件
#include <fstream>
#include <string>
#include <Windows.h>

CExport::CExport(CListCtrl* pListCtrl)
    : m_pListData(pListCtrl)  // 构造函数接受 CListCtrl 控件指针
{
}

// 实际导出数据到文件的函数
void CExport::ExportDataToFile()
{
    // 设置固定的文件路径，保存在项目根目录
    CString filePath = _T("data.txt");

    // 打开文件进行写入，使用 std::ofstream (写入文本文件)
    std::ofstream file_out(filePath, std::ios::out | std::ios::binary);

    if (!file_out.is_open())
    {
        AfxMessageBox(L"Failed to open the file for writing");
        return;
    }

    // 写入 UTF-8 BOM（字节顺序标记）
    unsigned char BOM[] = { 0xEF, 0xBB, 0xBF }; // UTF-8 BOM
    file_out.write(reinterpret_cast<char*>(BOM), sizeof(BOM));

    // 获取 CListCtrl 数据
    int nItemCount = m_pListData->GetItemCount();  // 获取项目数
    int nColumnCount = m_pListData->GetHeaderCtrl()->GetItemCount();  // 获取列数

    // 如果没有数据，弹出提示并返回
    if (nItemCount == 0 || nColumnCount == 0)
    {
        AfxMessageBox(L"No data to export");
        return;
    }

    // 开始导出数据
    for (int row = 0; row < nItemCount; ++row)
    {
        for (int col = 0; col < nColumnCount; ++col)
        {
            CString cellText = m_pListData->GetItemText(row, col);  // 获取每个单元格的文本

            // 使用 WideCharToMultiByte 转换 CString（宽字符）为 UTF-8 字符串
            int bufferSize = WideCharToMultiByte(CP_UTF8, 0, cellText, -1, nullptr, 0, nullptr, nullptr);
            char* utf8Buffer = new char[bufferSize];
            WideCharToMultiByte(CP_UTF8, 0, cellText, -1, utf8Buffer, bufferSize, nullptr, nullptr);

            // 将 UTF-8 字符串写入文件
            file_out.write(utf8Buffer, bufferSize - 1); // 去掉字符串末尾的 null 字符
            file_out << "\t";  // 使用制表符分隔单元格数据
            delete[] utf8Buffer;
        }
        file_out << std::endl;  // 每一行结束后换行
    }

    // 关闭文件
    file_out.close();

    // 提示导出完成
    AfxMessageBox(L"Data export completed successfully!");
}

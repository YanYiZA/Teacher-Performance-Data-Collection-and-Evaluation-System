#include "pch.h"
#include "CExport.h"
#include <afx.h>
#include <afxdlgs.h>
#include <fstream>

CExport::CExport(CListCtrl* pListCtrl)
    : m_pListData(pListCtrl)
{
}

void CExport::ExportDataToFile()
{
    CString filePath = _T("data.txt");
    std::wofstream file_out(filePath, std::ios::out);

    if (!file_out.is_open())
    {
        AfxMessageBox(L"无法打开文件进行写入");
        return;
    }

    file_out.imbue(std::locale("zh_CN.UTF-8"));  // 设置输出文件的编码为 UTF-8

    int nItemCount = m_pListData->GetItemCount();
    int nColumnCount = m_pListData->GetHeaderCtrl()->GetItemCount();

    if (nItemCount == 0 || nColumnCount == 0)
    {
        AfxMessageBox(L"没有数据可导出");
        return;
    }

    try
    {
        for (int row = 0; row < nItemCount; ++row)
        {
            for (int col = 0; col < nColumnCount; ++col)
            {
                CString cellText = m_pListData->GetItemText(row, col);
                file_out << (LPCTSTR)cellText;  // 直接输出宽字符文本
                if (col < nColumnCount - 1)
                    file_out << L"\t";  // 用制表符分隔
            }
            file_out << L"\n";  // 每一行结束换行
        }
    }
    catch (...)
    {
        AfxMessageBox(L"导出时发生未知错误");
        file_out.close();
        return;
    }

    file_out.close();
}

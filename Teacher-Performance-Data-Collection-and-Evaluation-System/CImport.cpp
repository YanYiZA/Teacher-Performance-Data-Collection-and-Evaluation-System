#include "pch.h"
#include "CImport.h"
#include <afxdlgs.h>
#include <msxml6.h> // 用于处理 XML 的库

// 构造函数
CImport::CImport(CListCtrl* pListCtrl)
    : m_pListCtrl(pListCtrl)
{
}

// 析构函数
CImport::~CImport()
{
}

// 批量导入文件
void CImport::ImportData(const std::vector<std::wstring>& filePaths)
{
    for (const auto& filePath : filePaths)
    {
        // 这里用简单的方式解析 docx 文件，实际使用时需要更复杂的处理
        // 假设这里是从文件中提取数据的地方
        std::wstring name = L"科比"; // 模拟数据
        int teaching = 694;
        int research = 92;
        int other = 50;
        int total = teaching + research + other;

        AddDataToList(name, teaching, research, other, total); // 将提取的数据添加到 List 控件
    }
}

// 将数据添加到 List 控件
void CImport::AddDataToList(const std::wstring& name, int teaching, int research, int other, int total)
{
    int index = m_pListCtrl->InsertItem(0, CString(name.c_str()));
    m_pListCtrl->SetItemText(index, 1, CString(std::to_wstring(teaching).c_str()));
    m_pListCtrl->SetItemText(index, 2, CString(std::to_wstring(research).c_str()));
    m_pListCtrl->SetItemText(index, 3, CString(std::to_wstring(other).c_str()));
    m_pListCtrl->SetItemText(index, 4, CString(std::to_wstring(total).c_str()));
}

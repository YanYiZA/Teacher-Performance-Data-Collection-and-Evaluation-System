#pragma once

#include <afxwin.h>
#include <vector>
#include <string>

// CImport 类负责从 DOCX 文件导入数据并填充到列表控件中
class CImport
{
public:
    // 构造函数与析构函数
    CImport(CListCtrl* pListCtrl);
    ~CImport();

    // 导入多个 DOCX 文件
    void ImportDocxFiles(const std::vector<std::wstring>& filePaths);

    // 向 ListCtrl 中添加数据
    void AddDataToList(const std::wstring& name, double teachingWork, double researchWork, double scientificWork, double otherWork);

private:
    // ListCtrl 控件的指针，保存导入数据
    CListCtrl* m_pListCtrl;

    // 辅助函数，用于初始化 ListCtrl 控件的列
    void InitializeListCtrl();
};

#pragma once
#include <afxwin.h>
#include <vector>
#include <string>

class CImport
{
public:
    CImport(CListCtrl* pListCtrl);
    ~CImport();

    void ImportData(const std::vector<std::wstring>& filePaths); // 批量导入文件的接口

private:
    CListCtrl* m_pListCtrl; // 用于操作 List 控件
    void AddDataToList(const std::wstring& name, int teaching, int research, int other, int total); // 添加数据到List控件
};

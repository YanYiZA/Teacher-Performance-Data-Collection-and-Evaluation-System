#pragma once
#include <afxwin.h>  // 包含 MFC 相关头文件
#include <string>
#include <vector>

// CExport 类负责将 CListCtrl 控件中的数据导出到文件
class CExport
{
public:
    // 构造函数，接受 CListCtrl 控件指针
    CExport(CListCtrl* pListCtrl);

    // 导出数据到文件，文件路径由用户选择
    void ExportDataToFile();
    CListCtrl* m_pListData;

private:
};

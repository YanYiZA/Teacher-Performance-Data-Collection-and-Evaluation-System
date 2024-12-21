#pragma once

#include <afxwin.h>
#include <vector>
#include <string>
#include "TeacherPerformance.h"
#include "Teacher-Performance-Data-Collection-and-Evaluation-SystemDlg.h"

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
    void AddDataToList(const std::wstring& name, int teachingWork, int researchWork, int scientificWork, int otherWork);
    // 向数据列表添加新数据
    void AddPerformanceData(const TeacherPerformance& data);

private:
    // ListCtrl 控件的指针，保存导入数据
    CListCtrl* m_pListCtrl;
    std::vector<TeacherPerformance> m_dataList;
    // 辅助函数，用于初始化 ListCtrl 控件的列
    void InitializeListCtrl();
};

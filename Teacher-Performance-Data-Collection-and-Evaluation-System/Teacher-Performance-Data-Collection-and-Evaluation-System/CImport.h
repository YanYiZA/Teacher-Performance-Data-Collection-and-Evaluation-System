#pragma once

#include <afxwin.h>
#include <vector>
#include <string>
#include "TeacherPerformance.h"
#include "Teacher-Performance-Data-Collection-and-Evaluation-SystemDlg.h"

// CImport �ฺ��� DOCX �ļ��������ݲ���䵽�б�ؼ���
class CImport
{
public:
    // ���캯������������
    CImport(CListCtrl* pListCtrl);
    ~CImport();

    // ������ DOCX �ļ�
    void ImportDocxFiles(const std::vector<std::wstring>& filePaths);
    // �� ListCtrl ���������
    void AddDataToList(const std::wstring& name, int teachingWork, int researchWork, int scientificWork, int otherWork);
    // �������б����������
    void AddPerformanceData(const TeacherPerformance& data);

private:
    // ListCtrl �ؼ���ָ�룬���浼������
    CListCtrl* m_pListCtrl;
    std::vector<TeacherPerformance> m_dataList;
    // �������������ڳ�ʼ�� ListCtrl �ؼ�����
    void InitializeListCtrl();
};

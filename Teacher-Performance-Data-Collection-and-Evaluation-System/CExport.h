#pragma once
#include <afxwin.h>  // ���� MFC ���ͷ�ļ�
#include <string>
#include <vector>

// CExport �ฺ�� CListCtrl �ؼ��е����ݵ������ļ�
class CExport
{
public:
    // ���캯�������� CListCtrl �ؼ�ָ��
    CExport(CListCtrl* pListCtrl);

    // �������ݵ��ļ����ļ�·�����û�ѡ��
    void ExportDataToFile();
    CListCtrl* m_pListData;

private:
};

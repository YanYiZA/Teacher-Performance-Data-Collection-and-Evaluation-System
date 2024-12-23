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
        AfxMessageBox(L"�޷����ļ�����д��");
        return;
    }

    file_out.imbue(std::locale("zh_CN.UTF-8"));  // ��������ļ��ı���Ϊ UTF-8

    int nItemCount = m_pListData->GetItemCount();
    int nColumnCount = m_pListData->GetHeaderCtrl()->GetItemCount();

    if (nItemCount == 0 || nColumnCount == 0)
    {
        AfxMessageBox(L"û�����ݿɵ���");
        return;
    }

    try
    {
        for (int row = 0; row < nItemCount; ++row)
        {
            for (int col = 0; col < nColumnCount; ++col)
            {
                CString cellText = m_pListData->GetItemText(row, col);
                file_out << (LPCTSTR)cellText;  // ֱ��������ַ��ı�
                if (col < nColumnCount - 1)
                    file_out << L"\t";  // ���Ʊ���ָ�
            }
            file_out << L"\n";  // ÿһ�н�������
        }
    }
    catch (...)
    {
        AfxMessageBox(L"����ʱ����δ֪����");
        file_out.close();
        return;
    }

    file_out.close();
}

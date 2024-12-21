#include "pch.h"
#include "CExport.h"
#include <afx.h>
#include <afxdlgs.h>  // ���� MFC �ļ��Ի������ͷ�ļ�
#include <fstream>
#include <string>
#include <Windows.h>

CExport::CExport(CListCtrl* pListCtrl)
    : m_pListData(pListCtrl)  // ���캯������ CListCtrl �ؼ�ָ��
{
}

// ʵ�ʵ������ݵ��ļ��ĺ���
void CExport::ExportDataToFile()
{
    // ���ù̶����ļ�·������������Ŀ��Ŀ¼
    CString filePath = _T("data.txt");

    // ���ļ�����д�룬ʹ�� std::ofstream (д���ı��ļ�)
    std::ofstream file_out(filePath, std::ios::out | std::ios::binary);

    if (!file_out.is_open())
    {
        AfxMessageBox(L"Failed to open the file for writing");
        return;
    }

    // д�� UTF-8 BOM���ֽ�˳���ǣ�
    unsigned char BOM[] = { 0xEF, 0xBB, 0xBF }; // UTF-8 BOM
    file_out.write(reinterpret_cast<char*>(BOM), sizeof(BOM));

    // ��ȡ CListCtrl ����
    int nItemCount = m_pListData->GetItemCount();  // ��ȡ��Ŀ��
    int nColumnCount = m_pListData->GetHeaderCtrl()->GetItemCount();  // ��ȡ����

    // ���û�����ݣ�������ʾ������
    if (nItemCount == 0 || nColumnCount == 0)
    {
        AfxMessageBox(L"No data to export");
        return;
    }

    // ��ʼ��������
    for (int row = 0; row < nItemCount; ++row)
    {
        for (int col = 0; col < nColumnCount; ++col)
        {
            CString cellText = m_pListData->GetItemText(row, col);  // ��ȡÿ����Ԫ����ı�

            // ʹ�� WideCharToMultiByte ת�� CString�����ַ���Ϊ UTF-8 �ַ���
            int bufferSize = WideCharToMultiByte(CP_UTF8, 0, cellText, -1, nullptr, 0, nullptr, nullptr);
            char* utf8Buffer = new char[bufferSize];
            WideCharToMultiByte(CP_UTF8, 0, cellText, -1, utf8Buffer, bufferSize, nullptr, nullptr);

            // �� UTF-8 �ַ���д���ļ�
            file_out.write(utf8Buffer, bufferSize - 1); // ȥ���ַ���ĩβ�� null �ַ�
            file_out << "\t";  // ʹ���Ʊ���ָ���Ԫ������
            delete[] utf8Buffer;
        }
        file_out << std::endl;  // ÿһ�н�������
    }

    // �ر��ļ�
    file_out.close();

    // ��ʾ�������
    AfxMessageBox(L"Data export completed successfully!");
}

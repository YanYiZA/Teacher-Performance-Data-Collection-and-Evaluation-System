#include "pch.h"
#include "CImport.h"
#include <Windows.h>
#include <string>
#include <vector>
#include <iostream>
#include <fstream>
#include <sstream>
#include <commdlg.h>

// ���캯��
CImport::CImport(CListCtrl* pListCtrl)
    : m_pListCtrl(pListCtrl)
{
    if (m_pListCtrl == nullptr) {
        throw std::invalid_argument("pListCtrl ����Ϊ nullptr");
    }
}

// ��������
CImport::~CImport()
{
}


#include <locale>
#include <codecvt>

#include <stdexcept>

void CImport::ImportDocxFiles(const std::vector<std::wstring>& filePaths)
{
    try {
        for (const auto& filePath : filePaths)
        {
            std::wstring command = L"python parse_docx.pyw \"" + filePath + L"\"";

            // ��������
            FILE* pipe = _wpopen(command.c_str(), L"r");
            if (pipe == nullptr) {
                AfxMessageBox(L"�޷����� Python �ű���");
                return;
            }

            // ��ȡ Python �ű������
            char buffer[128];
            std::string result;
            while (fgets(buffer, sizeof(buffer), pipe) != nullptr) {
                result += buffer;
            }
            fclose(pipe);

            // ����ȡ�����ַ���ת��Ϊ���ַ�
            std::wstring wideResult;
            try {
                std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
                wideResult = converter.from_bytes(result);
            }
            catch (const std::exception& e) {
                CString errorMsg;
                errorMsg.Format(L"�ַ�ת��ʧ�ܣ�%s", CString(e.what()));
                AfxMessageBox(errorMsg);
                continue;
            }

            // �������ص����ݲ���䵽�б���
            std::wstringstream ss(wideResult);
            std::wstring name;
            int teachingWork = 0, researchWork = 0, scientificWork = 0, otherWork = 0;
            std::wstring line;

            while (std::getline(ss, line))
            {
                std::wstringstream lineStream(line);
                std::wstring currentName, workType, performance;

                std::getline(lineStream, currentName, L',');
                std::getline(lineStream, workType, L',');
                std::getline(lineStream, performance, L',');

                // �����ǰ�е�������Ϊ�գ����������
                if (!currentName.empty()) {
                    name = currentName;
                }

                // ���ݹ������ͶԼ�Ч���л���
                double performanceValue = 0.0;
                try {
                    performanceValue = std::stod(performance);
                }
                catch (...) {
                    performanceValue = 0.0; // ���ת��ʧ�ܣ�����Ч����Ϊ 0
                }

                if (workType == L"��ѧ����") {
                    teachingWork += performanceValue;
                }
                else if (workType == L"���й���") {
                    researchWork += performanceValue;
                }
                else if (workType == L"���й���") {
                    scientificWork += performanceValue;
                }
                else if (workType == L"��������") {
                    otherWork += performanceValue;
                }
            }

            // ��������ӵ� ListCtrl ��
            AddDataToList(name, teachingWork, researchWork, scientificWork, otherWork);
        }
    }
    catch (const std::exception& e) {
        CString errorMsg;
        errorMsg.Format(L"�����쳣: %s", CString(e.what()));
        AfxMessageBox(errorMsg);
    }
    catch (...) {
        AfxMessageBox(L"������һ��δ֪�Ĵ���");
    }
}




// ��������ӵ� List �ؼ�
void CImport::AddDataToList(const std::wstring& name, int teachingWork, int researchWork, int scientificWork, int otherWork)
{
    int index = m_pListCtrl->InsertItem(0, CString(name.c_str()));


    auto formatDouble = [](double value) -> CString {
        // �� double תΪ�ַ�����������λС��
        std::wostringstream oss;
        oss.precision(2);
        oss << std::fixed << value;
        return CString(oss.str().c_str());
        };

    // ����Ч���ݻ��ܲ���ʾ
    m_pListCtrl->SetItemText(index, 1, formatDouble(teachingWork));
    m_pListCtrl->SetItemText(index, 2, formatDouble(researchWork));
    m_pListCtrl->SetItemText(index, 3, formatDouble(scientificWork));
    m_pListCtrl->SetItemText(index, 4, formatDouble(otherWork));
    // �����ܼ�Ч�������
    double totalPerformance = teachingWork + researchWork + scientificWork + otherWork;
    m_pListCtrl->SetItemText(index, 5, formatDouble(totalPerformance));
    m_pListCtrl->SetItemText(index, 6, L"");

    // ��������ӵ�m_dataList
    TeacherPerformance data;
    data.name = CString(name.c_str());
    data.teachingPerformance = teachingWork;
    data.teachPerformance = teachingWork;
    data.researchPerformance = researchWork;
    data.otherPerformance = otherWork;
    data.totalPerformance = totalPerformance;
    data.rank = 0;  // ������ʼ��Ϊ0

    // ��ӵ�m_dataList
    m_dataList.push_back(data);

}
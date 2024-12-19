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
        // ��ʼ������
        for (const auto& filePath : filePaths)
        {
            // ʹ�� pythonw �����нű����������Ա��ⵯ�������д���
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
            std::wstring line;
            std::wstringstream ss(wideResult);
            std::wstring name;
            double teachingWork = 0, researchWork = 0, scientificWork = 0, otherWork = 0;

            while (std::getline(ss, line))
            {
                std::wstringstream lineStream(line);
                std::wstring currentName, workType, performance;

                // �ö��Ų������
                std::getline(lineStream, currentName, L',');
                std::getline(lineStream, workType, L',');
                std::getline(lineStream, performance, L',');

                // �����ǰ�е�������Ϊ�գ����������
                if (!currentName.empty()) {
                    name = currentName;
                }

                // ���ݹ������ͶԼ�Ч���л���
                if (workType == L"��ѧ����") {
                    try {
                        teachingWork += std::stod(performance);
                    }
                    catch (const std::invalid_argument&) {
                        teachingWork += 0; // ���ת��ʧ�ܣ�����Ч����Ϊ 0
                    }
                    catch (const std::out_of_range&) {
                        teachingWork += 0; // ���ת��ʧ�ܣ�����Ч����Ϊ 0
                    }
                }
                else if (workType == L"���й���") {
                    try {
                        researchWork += std::stod(performance);
                    }
                    catch (const std::invalid_argument&) {
                        researchWork += 0;
                    }
                    catch (const std::out_of_range&) {
                        researchWork += 0;
                    }
                }
                else if (workType == L"���й���") {
                    try {
                        scientificWork += std::stod(performance);
                    }
                    catch (const std::invalid_argument&) {
                        scientificWork += 0;
                    }
                    catch (const std::out_of_range&) {
                        scientificWork += 0;
                    }
                }
                else if (workType == L"��������") {
                    try {
                        otherWork += std::stod(performance);
                    }
                    catch (const std::invalid_argument&) {
                        otherWork += 0;
                    }
                    catch (const std::out_of_range&) {
                        otherWork += 0;
                    }
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
void CImport::AddDataToList(const std::wstring& name, double teachingWork, double researchWork, double scientificWork, double otherWork)
{
    int index = m_pListCtrl->InsertItem(0, CString(name.c_str()));

    // ��ʽ��Ϊ��λС��������ӵ��ؼ���
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

    // ����ܼ�Ч��������ʱΪ0���ͼ�Ч������Ҳ��0��
    m_pListCtrl->SetItemText(index, 5, L"");
    m_pListCtrl->SetItemText(index, 6, L"");
}
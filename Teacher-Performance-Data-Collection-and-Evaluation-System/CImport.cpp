#include "pch.h"
#include "CImport.h"
#include <Windows.h>
#include <string>
#include <vector>
#include <iostream>
#include <fstream>
#include <sstream>
#include <commdlg.h>

// 构造函数
CImport::CImport(CListCtrl* pListCtrl)
    : m_pListCtrl(pListCtrl)
{
    if (m_pListCtrl == nullptr) {
        throw std::invalid_argument("pListCtrl 不能为 nullptr");
    }
}

// 析构函数
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

            // 创建进程
            FILE* pipe = _wpopen(command.c_str(), L"r");
            if (pipe == nullptr) {
                AfxMessageBox(L"无法运行 Python 脚本！");
                return;
            }

            // 读取 Python 脚本的输出
            char buffer[128];
            std::string result;
            while (fgets(buffer, sizeof(buffer), pipe) != nullptr) {
                result += buffer;
            }
            fclose(pipe);

            // 将读取到的字符串转换为宽字符
            std::wstring wideResult;
            try {
                std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
                wideResult = converter.from_bytes(result);
            }
            catch (const std::exception& e) {
                CString errorMsg;
                errorMsg.Format(L"字符转换失败：%s", CString(e.what()));
                AfxMessageBox(errorMsg);
                continue;
            }

            // 解析返回的数据并填充到列表中
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

                // 如果当前行的姓名不为空，则更新姓名
                if (!currentName.empty()) {
                    name = currentName;
                }

                // 根据工作类型对绩效进行汇总
                double performanceValue = 0.0;
                try {
                    performanceValue = std::stod(performance);
                }
                catch (...) {
                    performanceValue = 0.0; // 如果转换失败，将绩效设置为 0
                }

                if (workType == L"教学工作") {
                    teachingWork += performanceValue;
                }
                else if (workType == L"教研工作") {
                    researchWork += performanceValue;
                }
                else if (workType == L"科研工作") {
                    scientificWork += performanceValue;
                }
                else if (workType == L"其它工作") {
                    otherWork += performanceValue;
                }
            }

            // 将数据添加到 ListCtrl 中
            AddDataToList(name, teachingWork, researchWork, scientificWork, otherWork);
        }
    }
    catch (const std::exception& e) {
        CString errorMsg;
        errorMsg.Format(L"捕获到异常: %s", CString(e.what()));
        AfxMessageBox(errorMsg);
    }
    catch (...) {
        AfxMessageBox(L"发生了一个未知的错误！");
    }
}




// 将数据添加到 List 控件
void CImport::AddDataToList(const std::wstring& name, int teachingWork, int researchWork, int scientificWork, int otherWork)
{
    int index = m_pListCtrl->InsertItem(0, CString(name.c_str()));


    auto formatDouble = [](double value) -> CString {
        // 将 double 转为字符串并保留两位小数
        std::wostringstream oss;
        oss.precision(2);
        oss << std::fixed << value;
        return CString(oss.str().c_str());
        };

    // 将绩效数据汇总并显示
    m_pListCtrl->SetItemText(index, 1, formatDouble(teachingWork));
    m_pListCtrl->SetItemText(index, 2, formatDouble(researchWork));
    m_pListCtrl->SetItemText(index, 3, formatDouble(scientificWork));
    m_pListCtrl->SetItemText(index, 4, formatDouble(otherWork));
    // 计算总绩效，并添加
    double totalPerformance = teachingWork + researchWork + scientificWork + otherWork;
    m_pListCtrl->SetItemText(index, 5, formatDouble(totalPerformance));
    m_pListCtrl->SetItemText(index, 6, L"");

    // 将数据添加到m_dataList
    TeacherPerformance data;
    data.name = CString(name.c_str());
    data.teachingPerformance = teachingWork;
    data.teachPerformance = teachingWork;
    data.researchPerformance = researchWork;
    data.otherPerformance = otherWork;
    data.totalPerformance = totalPerformance;
    data.rank = 0;  // 排名初始化为0

    // 添加到m_dataList
    m_dataList.push_back(data);

}
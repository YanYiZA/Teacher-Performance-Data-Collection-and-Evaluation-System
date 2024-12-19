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
        // 开始处理函数
        for (const auto& filePath : filePaths)
        {
            // 调用 Python 脚本并获取返回的文本数据
            std::wstring command = L"python parse_docx.py \"" + filePath + L"\"";
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
            std::wstring line;
            std::wstringstream ss(wideResult);
            std::wstring name;
            double teachingWork = 0, researchWork = 0, scientificWork = 0, otherWork = 0;

            while (std::getline(ss, line))
            {
                std::wstringstream lineStream(line);
                std::wstring currentName, workType, performance;

                // 用逗号拆分数据
                std::getline(lineStream, currentName, L',');
                std::getline(lineStream, workType, L',');
                std::getline(lineStream, performance, L',');

                // 如果当前行的姓名不为空，则更新姓名
                if (!currentName.empty()) {
                    name = currentName;
                }

                // 根据工作类型对绩效进行汇总
                if (workType == L"教学工作") {
                    try {
                        teachingWork += std::stod(performance);
                    }
                    catch (const std::invalid_argument&) {
                        // 如果转换失败，将绩效设置为 0
                        teachingWork += 0;
                    }
                    catch (const std::out_of_range&) {
                        // 如果转换失败，将绩效设置为 0
                        teachingWork += 0;
                    }
                }
                else if (workType == L"教研工作") {
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
                else if (workType == L"科研工作") {
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
                else if (workType == L"其它工作") {
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
void CImport::AddDataToList(const std::wstring& name, double teachingWork, double researchWork, double scientificWork, double otherWork)
{
    int index = m_pListCtrl->InsertItem(0, CString(name.c_str()));

    // 将绩效数据汇总并显示
    std::wstring teachingWorkStr = std::to_wstring(teachingWork);
    m_pListCtrl->SetItemText(index, 1, CString(teachingWorkStr.c_str()));

    m_pListCtrl->SetItemText(index, 2, CString(std::to_wstring(researchWork).c_str()));
    m_pListCtrl->SetItemText(index, 3, CString(std::to_wstring(scientificWork).c_str()));
    m_pListCtrl->SetItemText(index, 4, CString(std::to_wstring(otherWork).c_str()));

    // 填充总绩效（这里暂时为0）和绩效排名（也填0）
    m_pListCtrl->SetItemText(index, 5, L"0");
    m_pListCtrl->SetItemText(index, 6, L"0");
}
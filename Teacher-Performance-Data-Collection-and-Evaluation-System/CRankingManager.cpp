#include "pch.h"
#include "CRankingManager.h"
#include <algorithm> // 如果使用了 std::sort 等算法库

CRankingManager::CRankingManager(CListCtrl* pListCtrl)
    : m_pListCtrl(pListCtrl)
{
}
/*
void CRankingManager::UpdateRanking(std::vector<TeacherPerformance>& dataList)
{
}
*/

void CRankingManager::UpdateRanking(std::vector<TeacherPerformance>& dataList)
{
    if (!m_pListCtrl)
    {
        AfxMessageBox(_T("m_pListCtrl is NULL!"));
        return;
    }
    int rowCount = m_pListCtrl->GetItemCount();
    if (rowCount <= 0 || dataList.empty())
    {
        AfxMessageBox(_T("列表中没有数据或数据列表为空！"));
        return;
    }
    std::vector<std::pair<int, int>> performance;
    
    for (int i = 0; i < rowCount; ++i)
    {
        CString totalScore = m_pListCtrl->GetItemText(i, 5); // 第6列是总绩效
        int score = _ttoi(totalScore);
        TRACE(_T("Row %d TotalScore: %s -> %d\n"), i, totalScore, score);
        performance.emplace_back(i, score);
    }
    
    //排序
    std::sort(performance.begin(), performance.end(), [](const auto& a, const auto& b) {
        return a.second > b.second; // 降序
        });
    
    for (size_t i = 0; i < performance.size(); ++i)
    {
        
        int rowIndex = performance[i].first;
        CString rankStr;
        rankStr.Format(_T("%d"), static_cast<int>(i + 1));
        // 更新 ListCtrl 中的排名列
        m_pListCtrl->SetItemText(rowIndex, 6, rankStr);

        // 同步更新到 m_dataList
        dataList[rowIndex].rank = static_cast<int>(i + 1);

        TRACE(_T("Updating rank for Row %d -> Rank %d\n"), rowIndex, i + 1); // 调试输出
    }
    //AfxMessageBox(_T("排名更新完成！"));
}
#include "pch.h"
#include "CRankingManager.h"
#include <algorithm> // ���ʹ���� std::sort ���㷨��

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
        AfxMessageBox(_T("�б���û�����ݻ������б�Ϊ�գ�"));
        return;
    }
    std::vector<std::pair<int, int>> performance;
    
    for (int i = 0; i < rowCount; ++i)
    {
        CString totalScore = m_pListCtrl->GetItemText(i, 5); // ��6�����ܼ�Ч
        int score = _ttoi(totalScore);
        TRACE(_T("Row %d TotalScore: %s -> %d\n"), i, totalScore, score);
        performance.emplace_back(i, score);
    }
    
    //����
    std::sort(performance.begin(), performance.end(), [](const auto& a, const auto& b) {
        return a.second > b.second; // ����
        });
    
    for (size_t i = 0; i < performance.size(); ++i)
    {
        
        int rowIndex = performance[i].first;
        CString rankStr;
        rankStr.Format(_T("%d"), static_cast<int>(i + 1));
        // ���� ListCtrl �е�������
        m_pListCtrl->SetItemText(rowIndex, 6, rankStr);

        // ͬ�����µ� m_dataList
        dataList[rowIndex].rank = static_cast<int>(i + 1);

        TRACE(_T("Updating rank for Row %d -> Rank %d\n"), rowIndex, i + 1); // �������
    }
    //AfxMessageBox(_T("����������ɣ�"));
}
#pragma once
#include <afxcmn.h>
#include <vector>

// 在 CRankingManager.h 或公共头文件中
// 在 CRankingManager.h 或公共头文件中
struct TeacherPerformance
{
    CString name;
    int teachingPerformance;
    int researchPerformance;
    int otherPerformance;
    int totalPerformance;
    int rank;
};

// 排名管理类
class CRankingManager
{
public:
    // 构造函数，接收指向 CListCtrl 的指针
    CRankingManager(CListCtrl* pListCtrl);

    void UpdateRanking();

    // 更新排名
    void UpdateRanking(std::vector<TeacherPerformance>& dataList);

private:
    CListCtrl* m_pListCtrl; // 用于操作主对话框中的 CListCtrl
};


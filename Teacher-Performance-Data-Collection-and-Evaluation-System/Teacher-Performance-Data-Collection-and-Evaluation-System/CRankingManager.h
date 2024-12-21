#pragma once
#include <afxcmn.h>
#include <vector>
#include "TeacherPerformance.h"

// 排名管理类
class CRankingManager
{
public:
    // 构造函数，接收指向 CListCtrl 的指针
    CRankingManager(CListCtrl* pListCtrl);

    

    // 更新排名
    void UpdateRanking(std::vector<TeacherPerformance>& dataList);

private:
    CListCtrl* m_pListCtrl; // 用于操作主对话框中的 CListCtrl
};


#pragma once
#include <afxcmn.h>
#include <vector>
#include "TeacherPerformance.h"

// ����������
class CRankingManager
{
public:
    // ���캯��������ָ�� CListCtrl ��ָ��
    CRankingManager(CListCtrl* pListCtrl);

    

    // ��������
    void UpdateRanking(std::vector<TeacherPerformance>& dataList);

private:
    CListCtrl* m_pListCtrl; // ���ڲ������Ի����е� CListCtrl
};



// Teacher-Performance-Data-Collection-and-Evaluation-SystemDlg.h: 头文件
//

#pragma once
#include "CRankingManager.h"

// CTeacherPerformanceDataCollectionandEvaluationSystemDlg 对话框
class CTeacherPerformanceDataCollectionandEvaluationSystemDlg : public CDialogEx
{
// 构造
public:
	CTeacherPerformanceDataCollectionandEvaluationSystemDlg(CWnd* pParent = nullptr);	// 标准构造函数
	void UpdateRankings();//更新排名函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_TEACHERPERFORMANCEDATACOLLECTIONANDEVALUATIONSYSTEM_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	CListCtrl m_listData;
	afx_msg void OnBnClickedButtonAdd();
//	CEdit m_editName;
//	CEdit m_editTeaching;
	CEdit m_editResearch;
	CEdit m_editOther;
	CEdit m_editTotal;
	CEdit m_editRank;
	afx_msg void OnBnClickedButtonSort();

	~CTeacherPerformanceDataCollectionandEvaluationSystemDlg(); // 析构函数

private:
	// 添加数据列表和数据结构
	// 教师绩效数据结构
	struct TeacherPerformance
	{
		CString name;               // 姓名
		int teachingPerformance;    // 教学绩效
		int teachPerformance;       // 教研绩效
		int researchPerformance;    // 科研绩效
		int otherPerformance;       // 其他绩效
		int totalPerformance;       // 总绩效
		int rank;                   // 排名
	};
	std::vector<TeacherPerformance> m_dataList; // 数据存储容器
	void UpdateListCtrl(); // 更新列表控件函数
	CRankingManager* m_pRankingManager; // 添加排名管理器指针
	

public:
	afx_msg void OnBnClickedButtonDelete();
	afx_msg void OnBnClickedButtonUpdate();
	CEdit m_editName;
	CEdit m_editTeaching;
	CEdit m_editTeach;
	afx_msg void OnBnClickedButtonTotal();
	afx_msg void OnBnClickedButtonSearch();
	CEdit m_queryName;
};

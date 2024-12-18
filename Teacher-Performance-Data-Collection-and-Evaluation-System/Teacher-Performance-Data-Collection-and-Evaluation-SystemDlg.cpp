
// Teacher-Performance-Data-Collection-and-Evaluation-SystemDlg.cpp: 实现文件
//

#include "pch.h"
#include "framework.h"
#include "Teacher-Performance-Data-Collection-and-Evaluation-System.h"
#include "Teacher-Performance-Data-Collection-and-Evaluation-SystemDlg.h"
#include "CRankingManager.h"
#include "afxdialogex.h"
#include <algorithm>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CTeacherPerformanceDataCollectionandEvaluationSystemDlg 对话框



CTeacherPerformanceDataCollectionandEvaluationSystemDlg::CTeacherPerformanceDataCollectionandEvaluationSystemDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_TEACHERPERFORMANCEDATACOLLECTIONANDEVALUATIONSYSTEM_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_DATA, m_listData);
	//  DDX_Control(pDX, IDC_EDIT3, m_editName);
	//  DDX_Control(pDX, IDC_EDIT4, m_editTeaching);
	DDX_Control(pDX, IDC_EDIT5, m_editResearch);
	DDX_Control(pDX, IDC_EDIT6, m_editOther);
	DDX_Control(pDX, IDC_EDIT7, m_editTotal);
	DDX_Control(pDX, IDC_EDIT8, m_editRank);
	DDX_Control(pDX, IDC_EDIT2, m_editName);
	DDX_Control(pDX, IDC_EDIT3, m_editTeaching);
	DDX_Control(pDX, IDC_EDIT4, m_editTeach);
	DDX_Control(pDX, IDC_EDIT11, m_queryName);
}

BEGIN_MESSAGE_MAP(CTeacherPerformanceDataCollectionandEvaluationSystemDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON_ADD, &CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonAdd)
	ON_BN_CLICKED(IDC_BUTTON_SORT, &CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonSort)
	ON_BN_CLICKED(IDC_BUTTON_DELETE, &CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonDelete)
	ON_BN_CLICKED(IDC_BUTTON_UPDATE, &CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonUpdate)
	//ON_EN_CHANGE(IDC_EDIT10, &CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnEnChangeEdit10)
	ON_BN_CLICKED(IDC_BUTTON_TOTAL, &CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonTotal)
	ON_BN_CLICKED(IDC_BUTTON_SEARCH, &CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonSearch)
END_MESSAGE_MAP()


// CTeacherPerformanceDataCollectionandEvaluationSystemDlg 消息处理程序

BOOL CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	CListCtrl& listCtrl = m_listData; // 假设 m_listData 是你绑定的 CListCtrl 对象
	listCtrl.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES); // 开启全行选择和网格线

	// 设置列头
	CStringArray columns;
	columns.Add(_T("姓名"));
	columns.Add(_T("教学工作"));
	columns.Add(_T("教研工作"));
	columns.Add(_T("科研工作"));
	columns.Add(_T("其他工作"));
	columns.Add(_T("总绩效"));
	columns.Add(_T("绩效排名"));

	for (int i = 0; i < columns.GetCount(); ++i)
	{
		listCtrl.InsertColumn(i, columns[i], LVCFMT_LEFT, 100); // 每列宽度 100
	}

	// 初始化排名管理器
	m_pRankingManager = new CRankingManager(&m_listData);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


//导入文件
void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
}

//增加数据
void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonAdd()
{
	// TODO: 在此添加控件通知处理程序代码
	 // 获取输入框内容
	CString name, teachingStr, teachStr,researchStr, otherStr;
	m_editName.GetWindowText(name);
	m_editTeaching.GetWindowText(teachingStr);
	m_editTeach.GetWindowText(teachStr);
	m_editResearch.GetWindowText(researchStr);
	m_editOther.GetWindowText(otherStr);

	// 验证输入是否为空
	if (name.IsEmpty() || teachingStr.IsEmpty() || teachStr.IsEmpty() || researchStr.IsEmpty() || otherStr.IsEmpty())
	{
		MessageBox(_T("请填写所有必要信息！"), _T("错误"), MB_ICONERROR);
		return;
	}

	// 转换绩效为整数
	int teaching = _ttoi(teachingStr);
	int teach = _ttoi(teachStr);
	int research = _ttoi(researchStr);
	int other = _ttoi(otherStr);
	int total = teaching + teach + research + other; // 计算总绩效

	// 添加到数据列表
	TeacherPerformance newPerformance = { name, teaching,teach, research, other, total, 0 }; // 排名初始为 0
	m_dataList.push_back(newPerformance);

	// 更新列表控件
	UpdateListCtrl();

	// 清空输入框
	m_editName.SetWindowText(_T(""));
	m_editTeaching.SetWindowText(_T(""));
	m_editTeach.SetWindowText(_T(""));
	m_editResearch.SetWindowText(_T(""));
	m_editOther.SetWindowText(_T(""));

	MessageBox(_T("数据添加成功！"), _T("提示"), MB_ICONINFORMATION);
	// 更新排名
	m_pRankingManager->UpdateRanking();
}

//更新排名
void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonSort()
{
	if (m_pRankingManager)
	{
		AfxMessageBox(_T("更新排名开始！")); // 检查是否调用
		m_pRankingManager->UpdateRanking();
		AfxMessageBox(_T("更新排名结束！")); // 检查是否调用完成
	}
	else
	{
		AfxMessageBox(_T("RankingManager 未初始化！"));
	}
}

//删除数据
void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonDelete()
{
	// TODO: 在此添加控件通知处理程序代码
	POSITION pos = m_listData.GetFirstSelectedItemPosition();
	if (pos == NULL)
	{
		MessageBox(_T("请先选择要删除的数据！"), _T("错误"), MB_ICONERROR);
		return;
	}

	int selectedIndex = m_listData.GetNextSelectedItem(pos);

	// 从数据列表中删除对应数据
	m_dataList.erase(m_dataList.begin() + selectedIndex);

	// 更新列表控件
	UpdateListCtrl();

	MessageBox(_T("数据删除成功！"), _T("提示"), MB_ICONINFORMATION);

	// 更新排名
	m_pRankingManager->UpdateRanking();
}

//修改
void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonUpdate()
{
	// TODO: 在此添加控件通知处理程序代码
	POSITION pos = m_listData.GetFirstSelectedItemPosition();
	if (pos == NULL)
	{
		MessageBox(_T("请先选择要修改的数据！"), _T("错误"), MB_ICONERROR);
		return;
	}

	int selectedIndex = m_listData.GetNextSelectedItem(pos);

	// 获取输入框内容
	CString name, teachingStr,teachStr, researchStr, otherStr;
	m_editName.GetWindowText(name);
	m_editTeaching.GetWindowText(teachingStr);
	m_editTeach.GetWindowText(teachStr);
	m_editResearch.GetWindowText(researchStr);
	m_editOther.GetWindowText(otherStr);

	// 验证输入是否为空
	if (name.IsEmpty() || teachingStr.IsEmpty() || teachStr.IsEmpty() || researchStr.IsEmpty() || otherStr.IsEmpty())
	{
		MessageBox(_T("请填写所有必要信息！"), _T("错误"), MB_ICONERROR);
		return;
	}

	// 转换绩效为整数
	int teaching = _ttoi(teachingStr);
	int teach = _ttoi(teachStr);
	int research = _ttoi(researchStr);
	int other = _ttoi(otherStr);
	int total = teaching + research + other; // 计算总绩效

	// 修改数据列表中对应的项
	m_dataList[selectedIndex] = { name, teaching,teach, research, other, total, 0 }; // 排名初始为 0

	// 更新列表控件
	UpdateListCtrl();

	// 清空输入框
	m_editName.SetWindowText(_T(""));
	m_editTeaching.SetWindowText(_T(""));
	m_editTeach.SetWindowText(_T(""));
	m_editResearch.SetWindowText(_T(""));
	m_editOther.SetWindowText(_T(""));

	MessageBox(_T("数据修改成功！"), _T("提示"), MB_ICONINFORMATION);

	// 更新排名
	m_pRankingManager->UpdateRanking();

}

/*
void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnEnChangeEdit10()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}
*/

//计算总分
void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonTotal()
{
	// TODO: 在此添加控件通知处理程序代码
	// 检查列表是否为空
	if (m_listData.GetItemCount() == 0)
	{
		AfxMessageBox(_T("列表中没有数据！"));
		return;
	}
	// 累加列表中所有的总绩效分
	int totalPerformanceSum = 0;
	for (int i = 0; i < m_listData.GetItemCount(); ++i)
	{
		// 获取总绩效列的文本
		CString strTotalPerformance = m_listData.GetItemText(i, 6); // 列索引 6 为“总绩效”
		totalPerformanceSum += _ttoi(strTotalPerformance); // 转换为整数累加
	}
}

//查询数据
void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::OnBnClickedButtonSearch()
{
	// TODO: 在此添加控件通知处理程序代码
	// 获取查询的姓名
	CString queryName;
	m_queryName.GetWindowText(queryName);

	// 遍历数据列表，查找匹配的记录
	for (size_t i = 0; i < m_dataList.size(); ++i)
	{
		if (m_dataList[i].name.CompareNoCase(queryName) == 0) // 比较姓名，忽略大小写
		{
			// 高亮显示对应行
			m_listData.SetItemState((int)i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
			m_listData.EnsureVisible((int)i, FALSE); // 确保高亮行可见

			// 填充 Edit 控件
			CString strTeaching,strTeach, strResearch, strOther, strTotal, strRank;
			strTeaching.Format(_T("%d"), m_dataList[i].teachingPerformance);
			strTeach.Format(_T("%d"), m_dataList[i].teachPerformance);
			strResearch.Format(_T("%d"), m_dataList[i].researchPerformance);
			strOther.Format(_T("%d"), m_dataList[i].otherPerformance);
			strTotal.Format(_T("%d"), m_dataList[i].totalPerformance);
			strRank.Format(_T("%d"), m_dataList[i].rank);

			m_editName.SetWindowText(m_dataList[i].name);
			m_editTeaching.SetWindowText(strTeaching);
			m_editTeach.SetWindowText(strTeach);
			m_editResearch.SetWindowText(strResearch);
			m_editOther.SetWindowText(strOther);
			m_editTotal.SetWindowText(strTotal);
			m_editRank.SetWindowText(strRank);

			return; // 找到后直接退出
		}
	}

	// 如果没有找到，弹出消息框提示
	AfxMessageBox(_T("未找到匹配的记录！"));
}

CTeacherPerformanceDataCollectionandEvaluationSystemDlg::~CTeacherPerformanceDataCollectionandEvaluationSystemDlg()
{
	// 如果有需要清理的资源（例如动态分配的内存），可以在这里处理
	if (m_pRankingManager)
	{
		delete m_pRankingManager;
		m_pRankingManager = nullptr;
	}
}

void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::UpdateListCtrl()
{
	// 清空列表控件
	m_listData.DeleteAllItems();

	// 遍历数据列表并填充控件
	for (size_t i = 0; i < m_dataList.size(); ++i)
	{
		const TeacherPerformance& data = m_dataList[i];

		int index = m_listData.InsertItem((int)i, data.name); // 插入姓名

		CString teaching, teach,research, other, total, rank;
		teaching.Format(_T("%d"), data.teachingPerformance);
		teach.Format(_T("%d"), data.teachPerformance);
		research.Format(_T("%d"), data.researchPerformance);
		other.Format(_T("%d"), data.otherPerformance);
		total.Format(_T("%d"), data.totalPerformance);
		rank.Format(_T("%d"), data.rank);

		m_listData.SetItemText(index, 1, teaching);
		m_listData.SetItemText(index, 2, teach);
		m_listData.SetItemText(index, 3, research);
		m_listData.SetItemText(index, 4, other);
		m_listData.SetItemText(index, 5, total);
		m_listData.SetItemText(index, 6, rank);
	}
}

//排名计算
void CTeacherPerformanceDataCollectionandEvaluationSystemDlg::UpdateRankings()
{
	// 使用临时数组保存排序后的数据
	std::vector<TeacherPerformance> sortedData = m_dataList;

	// 按总绩效从大到小排序
	std::sort(sortedData.begin(), sortedData.end(), [](const TeacherPerformance& a, const TeacherPerformance& b) {
		return a.totalPerformance > b.totalPerformance;
		});

	// 根据排序更新排名
	for (size_t i = 0; i < sortedData.size(); ++i)
	{
		sortedData[i].rank = static_cast<int>(i + 1); // 排名从1开始
	}

	// 更新原始数据列表中的排名
	for (auto& item : m_dataList)
	{
		for (const auto& sortedItem : sortedData)
		{
			if (item.name == sortedItem.name) // 按姓名匹配
			{
				item.rank = sortedItem.rank; // 更新排名
				break;
			}
		}
	}
	AfxMessageBox(_T("Reached TRACE Section"));
	// 更新列表控件显示
	UpdateListCtrl();
}
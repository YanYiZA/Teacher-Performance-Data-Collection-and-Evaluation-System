#ifndef TEACHERPERFORMANCE_H
#define TEACHERPERFORMANCE_H

#include <afxstr.h> // CString 的头文件
#include <vector>   // 标准容器

struct TeacherPerformance
{
    CString name;               // 姓名
    int teachingPerformance; // 教学绩效
    int teachPerformance;    // 教研绩效
    int researchPerformance; // 科研绩效
    int otherPerformance;    // 其他绩效
    int totalPerformance;    // 总绩效
    int rank;                   // 排名

    // 构造函数，初始化成员变量
    TeacherPerformance()
        : teachingPerformance(0),
        teachPerformance(0),
        researchPerformance(0),
        otherPerformance(0),
        totalPerformance(0),
        rank(0) {}

    // 带参数的构造函数
    TeacherPerformance(const CString& n, int teaching, int teach, int research, int other, int total, int r)
        : name(n),
        teachingPerformance(teaching),
        teachPerformance(teach),
        researchPerformance(research),
        otherPerformance(other),
        totalPerformance(total),
        rank(r) {
    }
};

#endif // TEACHERPERFORMANCE_H
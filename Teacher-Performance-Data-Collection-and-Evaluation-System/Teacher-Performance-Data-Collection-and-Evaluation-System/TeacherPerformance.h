#ifndef TEACHERPERFORMANCE_H
#define TEACHERPERFORMANCE_H

#include <afxstr.h> // CString ��ͷ�ļ�
#include <vector>   // ��׼����

struct TeacherPerformance
{
    CString name;               // ����
    int teachingPerformance; // ��ѧ��Ч
    int teachPerformance;    // ���м�Ч
    int researchPerformance; // ���м�Ч
    int otherPerformance;    // ������Ч
    int totalPerformance;    // �ܼ�Ч
    int rank;                   // ����

    // ���캯������ʼ����Ա����
    TeacherPerformance()
        : teachingPerformance(0),
        teachPerformance(0),
        researchPerformance(0),
        otherPerformance(0),
        totalPerformance(0),
        rank(0) {}

    // �������Ĺ��캯��
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
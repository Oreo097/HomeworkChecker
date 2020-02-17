from openpyxl import Workbook
m_grade=[['小明', '小华', '小鹏', '小艾', '小明', '小明', '小鹏'], [70, 20, 20, 20, 0, 0, 0], [['A', 'B', 'B', 'C', 'D', 'D', 'C'], ['B', 'B', 'B', 'B', 'B', 'B', 'B'], ['C', 'B', 'B', 'B', 'B', 'B', 'B'], ['A', 'A', 'A', 'B', 'B', 'B', 'C'], ['-', 2, '-', '-', '-', '-', '-'], ['-', '-', 1, '-', '-', '-', '-'], [1, '-', '-', '-', '-', '-', '-']]]
pointer=0
name_check_time=0
student_del=[]
while name_check_time<=len(m_grade):
    student=[]
    student.append(m_grade[0][pointer])
    student.append(pointer)
    grade_check_time=0
    check_pointer=pointer+1
    a=len(m_grade[0])
    while grade_check_time<(len(m_grade[0])-pointer-1):
        if student[0]==m_grade[0][check_pointer]:
            if student[1]<m_grade[1][check_pointer]:
                m_grade[1][pointer]=m_grade[1][check_pointer]
                m_grade[2][pointer]=m_grade[2][check_pointer]
                student_del.append(check_pointer)
            else:
                student_del.append(check_pointer)
        grade_check_time+=1
        check_pointer+=1
    name_check_time+=1
    pointer+=1
    time=0
    for delete in student_del:
        del m_grade[0][delete-time]
        del m_grade[1][delete-time]
        del m_grade[2][delete-time]
        time+=1
    student_del=[]
print(m_grade)
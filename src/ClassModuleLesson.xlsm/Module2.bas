Attribute VB_Name = "Module2"
'#################################################
'�N���X���W���[�����p
'#################################################
Option Explicit

Private Enum Info
    eName = 1
    eAge = 2
    eMobile = 3
End Enum

'==================================================
'���k�S���̏����Ǘ�����(�z����g�p����)
'==================================================
Private Sub UseArray3()
    Dim students(1 To 10, 1 To 3) As Variant
    
    '1�l��
    students(1, Info.eName) = "�H���@�����Y"
    students(1, Info.eAge) = 17
    students(1, Info.eMobile) = "XXX-XXXX-XXXX"
    
    '2�l�ڈȍ~_10�l
    
    Dim i As Long, j As Long
    For i = LBound(students) To UBound(students)
        For j = LBound(students, 2) To UBound(students, 2)
            Debug.Print students(i, j)
        Next
    Next
End Sub

'==================================================
'�N���X���W���[���Ő��k�S���̏����Ǘ�����c���L�q����
'==================================================
Private Sub UseClassModule2()
    Dim vStudents(1 To 10) As Variant
    Dim vStudent As student
    
    'Student�N���X�̃I�u�W�F�N�g�𐶐�����
    Set vStudent = New student
    
    '���k�̏���ݒ肷��
    vStudent.Name = "�H���@�����Y"
    vStudent.Age = 17
    vStudent.Mobile = "xxx-xxxx-xxxx"
        
    Set vStudents(1) = vStudent
    Debug.Print vStudents(1).Name
    Debug.Print vStudents(1).Age
    Debug.Print vStudents(1).Mobile
End Sub

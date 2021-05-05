Attribute VB_Name = "Module2"
'#################################################
'クラスモジュール応用
'#################################################
Option Explicit

Private Enum Info
    eName = 1
    eAge = 2
    eMobile = 3
End Enum

'==================================================
'生徒全員の情報を管理する(配列を使用する)
'==================================================
Private Sub UseArray3()
    Dim students(1 To 10, 1 To 3) As Variant
    
    '1人目
    students(1, Info.eName) = "羽生　健太郎"
    students(1, Info.eAge) = 17
    students(1, Info.eMobile) = "XXX-XXXX-XXXX"
    
    '2人目以降_10人
    
    Dim i As Long, j As Long
    For i = LBound(students) To UBound(students)
        For j = LBound(students, 2) To UBound(students, 2)
            Debug.Print students(i, j)
        Next
    Next
End Sub

'==================================================
'クラスモジュールで生徒全員の情報を管理する…を記述する
'==================================================
Private Sub UseClassModule2()
    Dim vStudents(1 To 10) As Variant
    Dim vStudent As student
    
    'Studentクラスのオブジェクトを生成する
    Set vStudent = New student
    
    '生徒の情報を設定する
    vStudent.Name = "羽生　健太郎"
    vStudent.Age = 17
    vStudent.Mobile = "xxx-xxxx-xxxx"
        
    Set vStudents(1) = vStudent
    Debug.Print vStudents(1).Name
    Debug.Print vStudents(1).Age
    Debug.Print vStudents(1).Mobile
End Sub

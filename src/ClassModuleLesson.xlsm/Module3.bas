Attribute VB_Name = "Module3"
'#################################################
'Studentsクラスを利用するコード
'#################################################
Option Explicit

'==================================================
'Studentsクラスを利用する
'==================================================
Private Sub UseClassModule()
    
    '表のデータを保持する変数
    Dim TableRange As Range
    Dim TableValue As Variant
    
    '表のデータを取得する
    With ThisWorkbook.Worksheets("Sheet1").Range("A3").CurrentRegion
        Set TableRange = .Resize(.Rows.Count - 1).Offset(1)
    End With
    
    '表のデータを配列に取得する
    TableValue = TableRange.Value
    
    '生徒全員の情報を保持するオブジェクトを代入する変数
    Dim oStudents As students
    
    Set oStudents = New students
    
    '表のデータを元にstudentsオブジェクトに値を設定する
    Dim i As Long
    For i = LBound(TableValue) To UBound(TableValue)
        oStudents.add TableValue(i, 1), TableValue(i, 2), TableValue(i, 3), TableValue(i, 4)
    Next
    
    'IDが「A0003」の生徒を検索する
    Dim vIndex As Variant
    
    vIndex = oStudents.SearchItemIndex("A0003")
    If vIndex = False Then
        MsgBox "指定したIDは見つかりません", vbInformation
    Else
        '書き換えるコード
        TableValue(vIndex, 3) = 17 '17歳に変更する
        'Debug.Print oStudents.Item(vIndex).Name
    End If
    '修正したデータを反映する
    TableRange.Value = TableValue
    
    '最初の生徒の氏名をイミディエイトウインドウに表示する
    Debug.Print oStudents.Item(1).Name
    Set oStudents = Nothing
    
End Sub

'==================================================
'デフォルトプロパティの設定ができる(一度エクスポートしてデフォルトプロパティーに指定したいプロパティーに「Attribute Value.VB_UserMemID = 0」を追記する
'==================================================

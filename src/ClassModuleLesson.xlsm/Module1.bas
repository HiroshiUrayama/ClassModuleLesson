Attribute VB_Name = "Module1"
'#################################################
'クラスモジュール
'#################################################
Option Explicit

'オブジェクトが破棄されるタイミング
'参照カウント方式です｡オブジェクトへの参照数が0になると､すぐに破棄されます｡
'そのとき、Terminateイベントが呼び出されます（※１）。

'オブジェクトがどこからも参照されなくなった場合に破棄される。
'下記のコードは、Data変数をNothingしているが、
'要素オブジェクト自体は破棄されずに残る(with文がDataを参照していて、参照カウントが1残っているから)

'https://thom.hateblo.jp/entry/2015/12/20/135035
Private Sub CollecitonSample()
    Dim Data As Collection
    Set Data = New Collection
    
    With Data
        .add "A"
        .add "B"
        'ここでNothingを入れても参照は残る
        'With文がDataを参照し続けているから
        Set Data = Nothing
    
        '1番目の要素をイミディエイトウインドウに表示する
        Debug.Print .Item(1)
        Debug.Print .Item(2)
    End With
    
End Sub

'エラーになる
Private Sub test()
    Debug.Print Name
End Sub

'普通の標準モジュールに記載する方法
Private Sub UseArray()
    Dim student(1 To 3) As Variant
    
    '生徒の情報を配列に代入する
    student(1) = "羽生　健太郎"
    student(2) = 29
    student(3) = "xxx-xxxx-xxxx"
    
    Dim i As Long
    For i = LBound(student) To UBound(student)
        Debug.Print student(i)
    Next
End Sub

'==================================================
'クラスモジュールを使う
'==================================================
Private Sub UseClassModule()

    'Studentクラスのオブジェクトを代入する変数
    Dim vStudent As student
    
    'Studentクラスのオブジェクトを生成する
    Set vStudent = New student
    
    '生徒の情報を設定する
    vStudent.Name = "羽生健太郎"
    vStudent.Age = 17
    vStudent.Mobile = "xxx-xxxxx-xxx"
    
    '生徒の情報を出力する
    Debug.Print vStudent.Name
    Debug.Print vStudent.Age
    Debug.Print vStudent.Mobile
End Sub

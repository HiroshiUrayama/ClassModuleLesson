VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "students"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'生徒の情報を保持する変数
'配列みたいに要素数をいちいち増やす必要がないので生徒の数が変わってもコードを変える必要がなく便利
Private mItems As Collection

'生徒の情報にアクセスするための変数
'生徒のIDを管理するために利用している
'文字列をインデックスの代わりに使用して要素を取り出すことができる／キーになるIDが文字列の場合、とても有効
Private mItemDictionary As Object

'変数の初期化(コンストラクタ)
Private Sub Class_Initialize()
    Set mItems = New Collection
    Set mItemDictionary = CreateObject("Scripting.Dictionary")
End Sub

'変数の破棄をする(デストラクタ)
Private Sub Class_Terminate()
    Set mItems = Nothing
    Set mItemDictionary = Nothing
End Sub

'生徒の情報を追加する
Public Sub add(ByVal vID As String, ByVal vName As String, ByVal vAge As Long, ByVal vMobile As String)
    Dim vStudent As student
    
    'DictionaryオブジェクトのキーにIDを追加する
    mItemDictionary.add Key:=vID, Item:=mItems.Count + 1
    
    '生徒を表すStudentクラスのオブジェクトを生成する
    Set vStudent = New student
    mItems.add vStudent
    
    '生徒の情報を追加する
    With vStudent
        .ID = vID
        .Name = vName
        .Age = vAge
        .Mobile = vMobile
    End With
    
    Set vStudent = Nothing
End Sub

'生徒の情報を取得する(読み取り専用：Letがない)
Public Property Get Item(ByVal Index As Long) As student
    Set Item = mItems.Item(Index)
End Property

'生徒の情報を検索する
Public Function SearchItemIndex(ByVal vID As String) As Variant
    SearchItemIndex = False
    If mItemDictionary.exists(vID) Then
        SearchItemIndex = mItemDictionary.Item(vID)
    End If
End Function


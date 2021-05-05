VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#################################################
'studentクラス
'#################################################
Option Explicit

Private mID As String
Private mName As String
Private mAge As Long
Private mMobile As String

Public Property Let ID(ByVal vID As String)
    mID = vID
End Property

Public Property Get ID() As String
    ID = mID
End Property

Public Property Let Name(ByVal vName As String)
    mName = vName
End Property

Public Property Get Name() As String
    Name = mName
End Property

'年齢を設定する
Public Property Let Age(ByVal vAge As Long)
        
    '設定できる年齢は15歳から18歳まで
    If vAge < 15 Or vAge > 18 Then
        vAge = -1
    End If
        
    mAge = vAge
End Property

'年齢を返す
Public Property Get Age() As Long
    Age = mAge
End Property

'電話番号を設定する
Public Property Let Mobile(ByVal vMobile As String)
    mMobile = vMobile
End Property

'電話番号を返す
Public Property Get Mobile() As String
    Mobile = mMobile
End Property

'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'Propertyプロシージャ
'★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
'プロパティを設定するプロシージャ
'Letが値を設定する時に使う
'Getは値を取得する時に使／こっちだけあるときは読み取り専用
'Functionプロシージャをクラスモジュールに記述すれば値を返すメソッドが、Subプロシージャを記述すれば、
'値を返さないメソッドを作ることができる。

'クラスモジュールを使う
'---------------------------------------------------

'Let Procedure
'---------------------------------------------------
'定義するプロシージャの引数を少なくとも1つは指定しなくてはならない。
'指定した最後の引数は、Property Letステートメントで定義されたプロシージャが呼び出された時に、
'プロパティに代入する実際の値を格納します≒最後の引数には、Optinalキーワードを指定して省略可能にすることはできない。


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

'���k�̏���ێ�����ϐ�
'�z��݂����ɗv�f���������������₷�K�v���Ȃ��̂Ő��k�̐����ς���Ă��R�[�h��ς���K�v���Ȃ��֗�
Private mItems As Collection

'���k�̏��ɃA�N�Z�X���邽�߂̕ϐ�
'���k��ID���Ǘ����邽�߂ɗ��p���Ă���
'��������C���f�b�N�X�̑���Ɏg�p���ėv�f�����o�����Ƃ��ł���^�L�[�ɂȂ�ID��������̏ꍇ�A�ƂĂ��L��
Private mItemDictionary As Object

'�ϐ��̏�����(�R���X�g���N�^)
Private Sub Class_Initialize()
    Set mItems = New Collection
    Set mItemDictionary = CreateObject("Scripting.Dictionary")
End Sub

'�ϐ��̔j��������(�f�X�g���N�^)
Private Sub Class_Terminate()
    Set mItems = Nothing
    Set mItemDictionary = Nothing
End Sub

'���k�̏���ǉ�����
Public Sub add(ByVal vID As String, ByVal vName As String, ByVal vAge As Long, ByVal vMobile As String)
    Dim vStudent As student
    
    'Dictionary�I�u�W�F�N�g�̃L�[��ID��ǉ�����
    mItemDictionary.add Key:=vID, Item:=mItems.Count + 1
    
    '���k��\��Student�N���X�̃I�u�W�F�N�g�𐶐�����
    Set vStudent = New student
    mItems.add vStudent
    
    '���k�̏���ǉ�����
    With vStudent
        .ID = vID
        .Name = vName
        .Age = vAge
        .Mobile = vMobile
    End With
    
    Set vStudent = Nothing
End Sub

'���k�̏����擾����(�ǂݎ���p�FLet���Ȃ�)
Public Property Get Item(ByVal Index As Long) As student
    Set Item = mItems.Item(Index)
End Property

'���k�̏�����������
Public Function SearchItemIndex(ByVal vID As String) As Variant
    SearchItemIndex = False
    If mItemDictionary.exists(vID) Then
        SearchItemIndex = mItemDictionary.Item(vID)
    End If
End Function

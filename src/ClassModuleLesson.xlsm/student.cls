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
'student�N���X
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

'�N���ݒ肷��
Public Property Let Age(ByVal vAge As Long)
        
    '�ݒ�ł���N���15�΂���18�΂܂�
    If vAge < 15 Or vAge > 18 Then
        vAge = -1
    End If
        
    mAge = vAge
End Property

'�N���Ԃ�
Public Property Get Age() As Long
    Age = mAge
End Property

'�d�b�ԍ���ݒ肷��
Public Property Let Mobile(ByVal vMobile As String)
    mMobile = vMobile
End Property

'�d�b�ԍ���Ԃ�
Public Property Get Mobile() As String
    Mobile = mMobile
End Property

'������������������������������������������������������������
'Property�v���V�[�W��
'������������������������������������������������������������
'�v���p�e�B��ݒ肷��v���V�[�W��
'Let���l��ݒ肷�鎞�Ɏg��
'Get�͒l���擾���鎞�Ɏg�^��������������Ƃ��͓ǂݎ���p
'Function�v���V�[�W�����N���X���W���[���ɋL�q����Βl��Ԃ����\�b�h���ASub�v���V�[�W�����L�q����΁A
'�l��Ԃ��Ȃ����\�b�h����邱�Ƃ��ł���B

'�N���X���W���[�����g��
'---------------------------------------------------

'Let Procedure
'---------------------------------------------------
'��`����v���V�[�W���̈��������Ȃ��Ƃ�1�͎w�肵�Ȃ��Ă͂Ȃ�Ȃ��B
'�w�肵���Ō�̈����́AProperty Let�X�e�[�g�����g�Œ�`���ꂽ�v���V�[�W�����Ăяo���ꂽ���ɁA
'�v���p�e�B�ɑ��������ۂ̒l���i�[���܂����Ō�̈����ɂ́AOptinal�L�[���[�h���w�肵�ďȗ��\�ɂ��邱�Ƃ͂ł��Ȃ��B

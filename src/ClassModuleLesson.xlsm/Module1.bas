Attribute VB_Name = "Module1"
'#################################################
'�N���X���W���[��
'#################################################
Option Explicit

'�I�u�W�F�N�g���j�������^�C�~���O
'�Q�ƃJ�E���g�����ł���I�u�W�F�N�g�ւ̎Q�Ɛ���0�ɂȂ�Ƥ�����ɔj������܂��
'���̂Ƃ��ATerminate�C�x���g���Ăяo����܂��i���P�j�B

'�I�u�W�F�N�g���ǂ�������Q�Ƃ���Ȃ��Ȃ����ꍇ�ɔj�������B
'���L�̃R�[�h�́AData�ϐ���Nothing���Ă��邪�A
'�v�f�I�u�W�F�N�g���͔̂j�����ꂸ�Ɏc��(with����Data���Q�Ƃ��Ă��āA�Q�ƃJ�E���g��1�c���Ă��邩��)

'https://thom.hateblo.jp/entry/2015/12/20/135035
Private Sub CollecitonSample()
    Dim Data As Collection
    Set Data = New Collection
    
    With Data
        .add "A"
        .add "B"
        '������Nothing�����Ă��Q�Ƃ͎c��
        'With����Data���Q�Ƃ������Ă��邩��
        Set Data = Nothing
    
        '1�Ԗڂ̗v�f���C�~�f�B�G�C�g�E�C���h�E�ɕ\������
        Debug.Print .Item(1)
        Debug.Print .Item(2)
    End With
    
End Sub

'�G���[�ɂȂ�
Private Sub test()
    Debug.Print Name
End Sub

'���ʂ̕W�����W���[���ɋL�ڂ�����@
Private Sub UseArray()
    Dim student(1 To 3) As Variant
    
    '���k�̏���z��ɑ������
    student(1) = "�H���@�����Y"
    student(2) = 29
    student(3) = "xxx-xxxx-xxxx"
    
    Dim i As Long
    For i = LBound(student) To UBound(student)
        Debug.Print student(i)
    Next
End Sub

'==================================================
'�N���X���W���[�����g��
'==================================================
Private Sub UseClassModule()

    'Student�N���X�̃I�u�W�F�N�g��������ϐ�
    Dim vStudent As student
    
    'Student�N���X�̃I�u�W�F�N�g�𐶐�����
    Set vStudent = New student
    
    '���k�̏���ݒ肷��
    vStudent.Name = "�H�������Y"
    vStudent.Age = 17
    vStudent.Mobile = "xxx-xxxxx-xxx"
    
    '���k�̏����o�͂���
    Debug.Print vStudent.Name
    Debug.Print vStudent.Age
    Debug.Print vStudent.Mobile
End Sub

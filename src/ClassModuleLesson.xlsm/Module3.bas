Attribute VB_Name = "Module3"
'#################################################
'Students�N���X�𗘗p����R�[�h
'#################################################
Option Explicit

'==================================================
'Students�N���X�𗘗p����
'==================================================
Private Sub UseClassModule()
    
    '�\�̃f�[�^��ێ�����ϐ�
    Dim TableRange As Range
    Dim TableValue As Variant
    
    '�\�̃f�[�^���擾����
    With ThisWorkbook.Worksheets("Sheet1").Range("A3").CurrentRegion
        Set TableRange = .Resize(.Rows.Count - 1).Offset(1)
    End With
    
    '�\�̃f�[�^��z��Ɏ擾����
    TableValue = TableRange.Value
    
    '���k�S���̏���ێ�����I�u�W�F�N�g��������ϐ�
    Dim oStudents As students
    
    Set oStudents = New students
    
    '�\�̃f�[�^������students�I�u�W�F�N�g�ɒl��ݒ肷��
    Dim i As Long
    For i = LBound(TableValue) To UBound(TableValue)
        oStudents.add TableValue(i, 1), TableValue(i, 2), TableValue(i, 3), TableValue(i, 4)
    Next
    
    'ID���uA0003�v�̐��k����������
    Dim vIndex As Variant
    
    vIndex = oStudents.SearchItemIndex("A0003")
    If vIndex = False Then
        MsgBox "�w�肵��ID�͌�����܂���", vbInformation
    Else
        '����������R�[�h
        TableValue(vIndex, 3) = 17 '17�΂ɕύX����
        'Debug.Print oStudents.Item(vIndex).Name
    End If
    '�C�������f�[�^�𔽉f����
    TableRange.Value = TableValue
    
    '�ŏ��̐��k�̎������C�~�f�B�G�C�g�E�C���h�E�ɕ\������
    Debug.Print oStudents.Item(1).Name
    Set oStudents = Nothing
    
End Sub

'==================================================
'�f�t�H���g�v���p�e�B�̐ݒ肪�ł���(��x�G�N�X�|�[�g���ăf�t�H���g�v���p�e�B�[�Ɏw�肵�����v���p�e�B�[�ɁuAttribute Value.VB_UserMemID = 0�v��ǋL����
'==================================================

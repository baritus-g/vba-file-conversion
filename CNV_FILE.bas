'���̓t�@�C�����o�̓t�@�C���֒u�����ďo��
'�����P�F���̓t�@�C���p�X
'�����Q�F�o�̓t�@�C���t�@�C���p�X
'�����R�F�u���O������
'�����S�F�u���㕶����
Private Sub CNV_FILE(ByVal s1 As String, ByVal s2 As String, hi1 As String, ho1 As String)
    
    Dim n1
    Dim n2
    Dim v

    '// �t�@�C���ԍ��擾
    n1 = FreeFile
    
    '// �V�[�P���V�������̓��[�h�Ńt�@�C�����J��
    Open s1 For Input As #n1
    
    '// �t�@�C���ԍ��擾
    n2 = FreeFile

    '// �V�[�P���V�����o�̓��[�h�Ńt�@�C�����J��
    Open s2 For Output As #n2

    '// ���̓t�@�C����EOF�܂Ń��[�v
    Do Until EOF(n1)
        '// ���̓t�@�C���̍s��ǂݍ���
        Line Input #n1, v
        '// �o�̓t�@�C���̃f�[�^����������
        Print #n2, v
    Loop
    
    '// �t�@�C�������
    Close #n1
    '// �t�@�C�������
    Close #n2

End Sub

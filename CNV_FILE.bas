'���̓t�@�C������o�̓t�@�C���֕������u�����ďo��
'�����P�F���̓t�@�C���p�X
'�����Q�F�o�̓t�@�C���t�@�C���p�X
'�����R�F�u���O������
'�����S�F�u���㕶����
Public Sub CNV_FILE(ByVal si As String, ByVal so As String, hi1 As String, ho1 As String)
    
    Dim ni
    Dim no
    Dim v

    '// �t�@�C���ԍ��擾
    ni = FreeFile
    
    '// �V�[�P���V�������̓��[�h�Ńt�@�C�����J��
    Open si For Input As #ni
    
    '// �t�@�C���ԍ��擾
    no = FreeFile

    '// �V�[�P���V�����o�̓��[�h�Ńt�@�C�����J��
    Open so For Output As #no

    '// ���̓t�@�C����EOF�܂Ń��[�v
    Do Until EOF(ni)
        '// ���̓t�@�C���̍s��ǂݍ���
        Line Input #ni, v
        '�������u��
        v = Replace(v, hi1, ho1)
        '// �o�̓t�@�C���̃f�[�^����������
        Print #no, v
    Loop
    
    '// �t�@�C�������
    Close #ni
    '// �t�@�C�������
    Close #no

End Sub

���̓t�@�C������o�̓t�@�C���֕������u�����ďo�͂���֐�
'�����P�F���̓t�@�C���p�X
'�����Q�F�o�̓t�@�C���p�X
'�����R�F�u���O������
'�����S�F�u���㕶����
Private Sub CNV_TXTFILE(ByVal path_i  As String, ByVal path_o As String, search_str As String, replace_str As String)
    
    Dim ni as integer
    Dim no as integer
    Dim v  as string

    '// �t�@�C���ԍ��擾�i���̓t�@�C���j
    ni = FreeFile
    
    '// �t�@�C�����J���i���̓��[�h�j
    Open path_i  For Input As #ni
    
    '// �t�@�C���ԍ��擾�i�o�̓t�@�C���j
    no = FreeFile

    '// �t�@�C�����J���i�o�̓��[�h�j
    Open path_o For Output As #no

    '// ���̓t�@�C���̏I�[�܂Ń��[�v
    Do Until EOF(ni)
        '// ���̓t�@�C���̓ǂݍ��݁i�P�s�j
        Line Input #ni, v
        '// �ǂݍ��񂾕������u��
        v = Replace(v, search_str, replace_str)
        '// �o�̓t�@�C�����������݁i�P�s�F�u���ρj
        Print #no, v
    Loop
    
    '// �t�@�C�������i���̓t�@�C���j
    Close #ni
    '// �t�@�C�������i�o�̓t�@�C���j
    Close #no

End Sub

'���̓u�b�N����o�̓u�b�N�֕������u�����ďo��
'�����P�F���̓t�@�C���p�X
'�����Q�F�o�̓t�@�C���t�@�C���p�X
'�����R�F�u���O������
'�����S�F�u���㕶����

Public Sub XLS_IN_OUT(ByVal si As String, ByVal so As String, hi1 As String, ho1 As String)

	Dim wb As Workbook
	Dim ws_io As Worksheet
	Dim spShape As Shape

	'���͌��̃t�@�C�����J���iExcel�u�b�N�j
	Workbooks.Open Filename:=si
	Set wb = ActiveWorkbook

	'�J�����u�b�N���̑S�V�[�g�����[�v
	For Each ws_io In wb.Worksheets
		'���[�N�V�[�g�̕����u���ixlPart�F�Z���̕�����v�^xlWhole�F�Z���̊��S��v�j
		ws_io.UsedRange.Replace What:=hi1, Replacement:=ho1, LookAt:=xlPart

		'�V�[�g���̃I�[�g�V�F�C�v���P�ΏۂƂ���i�Ō�̃I�[�g�V�F�C�v�܂ŌJ��Ԃ��j
		For Each spShape In ws_io.Shapes
			'�G���[�͖������đ��s
		    On Error Resume Next
			'�}�`�i�V�F�C�v�j�̕����u��
		    spShape.TextFrame.Characters.Text = Replace(spShape.TextFrame.Characters.Text, hi1, ho1)
		    On Error GoTo 0
		Next spShape
	Next ws_io

	'�o�͐�փt�@�C����ۑ��iExcel�u�b�N�j
	wb.SaveAs Filename:=so

	'�t�@�C�������iExcel�u�b�N�j
	wb.Close

End Sub

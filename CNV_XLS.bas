'���̓u�b�N����o�̓u�b�N�֕������u�����ďo��
'�����P�F����Excel�u�b�N�p�X
'�����Q�F�o��Excel�u�b�N�p�X
'�����R�F�u���O������
'�����S�F�u���㕶����
Private Sub XLS_IN_OUT(ByVal path_i As String, ByVal path_o As String, search_str As String, replace_str As String)

	Dim wb As Workbook
	Dim ws_io As Worksheet
	Dim spShape As Shape

	'Excel�u�b�N���J��
	Workbooks.Open Filename:=path_i
	Set wb = ActiveWorkbook

	'�J�����u�b�N���̑S�V�[�g�����[�v
	For Each ws_io In wb.Worksheets
		'���[�N�V�[�g�̕����u���ixlPart�F�Z���̕�����v�j
		ws_io.UsedRange.Replace What:=search_str, Replacement:=replace_str, LookAt:=xlPart

		'�}�`�i�V�F�C�v�j�̕����u��
		For Each spShape In ws_io.Shapes
			'�e�L�X�g�����}�`�����f
			If spShape.TextFrame2.HasText Then
			    spShape.TextFrame.Characters.Text = Replace(spShape.TextFrame.Characters.Text, search_str, replace_str)
			end if
		Next spShape
	Next ws_io

	'�o�͐��Excel�u�b�N��ۑ�
	wb.SaveAs Filename:=path_o

	'Excel�u�b�N�����
	wb.Close

End Sub

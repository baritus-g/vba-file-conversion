'入力ブックから出力ブックへ文字列を置換して出力
'引数１：入力Excelブックパス
'引数２：出力Excelブックパス
'引数３：置換前文字列
'引数４：置換後文字列
Private Sub XLS_IN_OUT(ByVal path_i As String, ByVal path_o As String, search_str As String, replace_str As String)

	Dim wb As Workbook
	Dim ws_io As Worksheet
	Dim spShape As Shape

	'Excelブックを開く
	Workbooks.Open Filename:=path_i
	Set wb = ActiveWorkbook

	'開いたブック内の全シート分ループ
	For Each ws_io In wb.Worksheets
		'ワークシートの文字置換（xlPart：セルの部分一致）
		ws_io.UsedRange.Replace What:=search_str, Replacement:=replace_str, LookAt:=xlPart

		'図形（シェイプ）の文字置換
		For Each spShape In ws_io.Shapes
			'テキストを持つ図形か判断
			If spShape.TextFrame2.HasText Then
			    spShape.TextFrame.Characters.Text = Replace(spShape.TextFrame.Characters.Text, search_str, replace_str)
			end if
		Next spShape
	Next ws_io

	'出力先へExcelブックを保存
	wb.SaveAs Filename:=path_o

	'Excelブックを閉じる
	wb.Close

End Sub

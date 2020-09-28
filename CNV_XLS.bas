'入力ブックから出力ブックへ文字列を置換して出力
'引数１：入力ファイルパス
'引数２：出力ファイルファイルパス
'引数３：置換前文字列
'引数４：置換後文字列

Public Sub XLS_IN_OUT(ByVal si As String, ByVal so As String, hi1 As String, ho1 As String)

	Dim wb As Workbook
	Dim ws_io As Worksheet
	Dim spShape As Shape

	'入力元のファイルを開く（Excelブック）
	Workbooks.Open Filename:=si
	Set wb = ActiveWorkbook

	'開いたブック内の全シート分ループ
	For Each ws_io In wb.Worksheets
		'ワークシートの文字置換（xlPart：セルの部分一致／xlWhole：セルの完全一致）
		ws_io.UsedRange.Replace What:=hi1, Replacement:=ho1, LookAt:=xlPart

		'シート内のオートシェイプを１つ対象とする（最後のオートシェイプまで繰り返す）
		For Each spShape In ws_io.Shapes
			'エラーは無視して続行
		    On Error Resume Next
			'図形（シェイプ）の文字置換
		    spShape.TextFrame.Characters.Text = Replace(spShape.TextFrame.Characters.Text, hi1, ho1)
		    On Error GoTo 0
		Next spShape
	Next ws_io

	'出力先へファイルを保存（Excelブック）
	wb.SaveAs Filename:=so

	'ファイルを閉じる（Excelブック）
	wb.Close

End Sub

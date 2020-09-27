'入力ファイルを出力ファイルへ置換して出力
'引数１：入力ファイルパス
'引数２：出力ファイルファイルパス
'引数３：置換前文字列
'引数４：置換後文字列
Private Sub CNV_FILE(ByVal s1 As String, ByVal s2 As String, hi1 As String, ho1 As String)
    
    Dim n1
    Dim n2
    Dim v

    '// ファイル番号取得
    n1 = FreeFile
    
    '// シーケンシャル入力モードでファイルを開く
    Open s1 For Input As #n1
    
    '// ファイル番号取得
    n2 = FreeFile

    '// シーケンシャル出力モードでファイルを開く
    Open s2 For Output As #n2

    '// 入力ファイルのEOFまでループ
    Do Until EOF(n1)
        '// 入力ファイルの行を読み込み
        Line Input #n1, v
        '// 出力ファイルのデータを書き込み
        Print #n2, v
    Loop
    
    '// ファイルを閉じる
    Close #n1
    '// ファイルを閉じる
    Close #n2

End Sub

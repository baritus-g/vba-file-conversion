'入力ファイルから出力ファイルへ文字列を置換して出力
'引数１：入力ファイルパス
'引数２：出力ファイルファイルパス
'引数３：置換前文字列
'引数４：置換後文字列
Public Sub CNV_FILE(ByVal si As String, ByVal so As String, hi1 As String, ho1 As String)
    
    Dim ni
    Dim no
    Dim v

    '// ファイル番号取得
    ni = FreeFile
    
    '// シーケンシャル入力モードでファイルを開く
    Open si For Input As #ni
    
    '// ファイル番号取得
    no = FreeFile

    '// シーケンシャル出力モードでファイルを開く
    Open so For Output As #no

    '// 入力ファイルのEOFまでループ
    Do Until EOF(ni)
        '// 入力ファイルの行を読み込み
        Line Input #ni, v
        '文字列を置換
        v = Replace(v, hi1, ho1)
        '// 出力ファイルのデータを書き込み
        Print #no, v
    Loop
    
    '// ファイルを閉じる
    Close #ni
    '// ファイルを閉じる
    Close #no

End Sub

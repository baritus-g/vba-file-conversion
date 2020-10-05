入力ファイルから出力ファイルへ文字列を置換して出力する関数
'引数１：入力ファイルパス
'引数２：出力ファイルパス
'引数３：置換前文字列
'引数４：置換後文字列
Private Sub CNV_TXTFILE(ByVal path_i  As String, ByVal path_o As String, search_str As String, replace_str As String)
    
    Dim ni as integer
    Dim no as integer
    Dim v  as string

    '// ファイル番号取得（入力ファイル）
    ni = FreeFile
    
    '// ファイルを開く（入力モード）
    Open path_i  For Input As #ni
    
    '// ファイル番号取得（出力ファイル）
    no = FreeFile

    '// ファイルを開く（出力モード）
    Open path_o For Output As #no

    '// 入力ファイルの終端までループ
    Do Until EOF(ni)
        '// 入力ファイルの読み込み（１行）
        Line Input #ni, v
        '// 読み込んだ文字列を置換
        v = Replace(v, search_str, replace_str)
        '// 出力ファイルを書き込み（１行：置換済）
        Print #no, v
    Loop
    
    '// ファイルを閉じる（入力ファイル）
    Close #ni
    '// ファイルを閉じる（出力ファイル）
    Close #no

End Sub

Attribute VB_Name = "RegExp"
Option Explicit

'*****************************************************************************
'[概要] 正規表現テスト
'[引数] 文字列，正規表現パターン
'[戻値] True:一致あり
'*****************************************************************************
Public Function RegExpTest(ByVal 文字列 As String, ByVal パターン As String) As Boolean
Attribute RegExpTest.VB_Description = "正規表現の条件を満たす場合Trueを返します"
    With CreateObject("VBScript.RegExp")
        .Pattern = パターン
        RegExpTest = .Test(文字列)
    End With
End Function

'*****************************************************************************
'[概要] 正規表現置換
'[引数] 文字列，正規表現パターン，置換文字列
'[戻値] 置換後の文字列
'*****************************************************************************
Public Function RegExpReplace(ByVal 文字列 As String, ByVal パターン As String, ByVal 置換文字列 As String) As String
Attribute RegExpReplace.VB_Description = "正規表現の置換を行います"
    With CreateObject("VBScript.RegExp")
        .Pattern = パターン
        .Global = True
        RegExpReplace = .Replace(文字列, 置換文字列)
    End With
End Function

'*****************************************************************************
'[概要] 正規表現マッチ文字列取得
'[引数] 文字列，正規表現パターン，何番目のサブマッチを取得するか
'[戻値] マッチ文字列
'*****************************************************************************
Public Function RegExpMatch(ByVal 文字列 As String, ByVal パターン As String, Optional ByVal サブマッチ番号 As Long = 0) As Variant
Attribute RegExpMatch.VB_Description = "正規表現に一致した文字を返します\n第3引数:サブマッチ文字を取得したい時のみ、そのIndexを指定します"
    With CreateObject("VBScript.RegExp")
        .Pattern = パターン
        If .Test(文字列) Then
            Dim objMatches
            Set objMatches = .Execute(文字列)
            If サブマッチ番号 > 0 Then
                RegExpMatch = objMatches(0).SubMatches(サブマッチ番号 - 1)
            Else
                RegExpMatch = objMatches(0).Value
            End If
        Else
            'マッチしない場合 #N/Aエラーを返す
            RegExpMatch = CVErr(xlErrNA)
        End If
    End With
End Function

'*****************************************************************************
'[概要] 正規表現マッチ文字列取得
'[引数] 文字列，正規表現パターン，配列の方向(0=行方向:1=列方向)
'[戻値] マッチ文字列の配列
'*****************************************************************************
Public Function RegExpMatches(ByVal 文字列 As String, ByVal パターン As String, Optional ByVal 方向 As Long = 0) As Variant
Attribute RegExpMatches.VB_Description = "正規表現に一致した箇所すべてを配列で返します\n配列数式形式(Ctrl+Shift+Enter)で取り出してください\n第3引数:配列の向きを指定します ※0=行方向(省略時) or 1=列方向"
    With CreateObject("VBScript.RegExp")
        .Pattern = パターン
        .Global = True
        If .Test(文字列) Then
            Dim objMatches
            Set objMatches = .Execute(文字列)
            ReDim Result(0 To objMatches.Count - 1)
            Dim i As Long
            For i = 0 To objMatches.Count - 1
                Result(i) = objMatches(i).Value
            Next
            If 方向 = 0 Then
                RegExpMatches = WorksheetFunction.Transpose(Result)
            Else
                RegExpMatches = Result
            End If
        Else
            'マッチしない場合 #N/Aエラーを返す
            RegExpMatches = CVErr(xlErrNA)
        End If
    End With
End Function


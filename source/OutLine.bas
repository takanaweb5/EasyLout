Attribute VB_Name = "OutLine"
Option Explicit

'*****************************************************************************
'[概要] 直前の段落番号の次の段落番号を取得する
'[引数] なし
'[戻値] 例：（ア）→（イ）
'*****************************************************************************
Public Function OutLineNext() As Variant
    Dim i As Long
    Dim Value As Variant
    Application.Volatile 'これがないと再計算されない

    '同一列を1行ずつ遡り、値の設定されたセルを検索
    For i = Application.ThisCell.Row - 1 To 1 Step -1
        Value = Application.ThisCell.EntireColumn.Rows(i).Value
        If Value <> "" Then
            '直前の段落番号から次の段落番号を取得
            If VarType(Value) = vbDouble Then
                OutLineNext = CDbl(GetNext(Value))
            Else
                OutLineNext = GetNext(Value)
            End If
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[概要] 段落番号の次の段落番号を取得する
'[引数] 直前の段落番号　例：（ア）
'[戻値] 例：（ア）→（イ）
'*****************************************************************************
Private Function GetNext(ByVal strOutLine As String) As String
    '左端の文字
    Dim strL As String
    strL = Left(strOutLine, 1)
    If InStr(1, "（(", strL) = 0 Then
        strL = ""
    End If

    '右端の文字
    Dim strR As String
    strR = Right(strOutLine, 1)
    If InStr(1, ".)．）", strR) = 0 Then
        strR = ""
    End If

    '両端の文字を削除
    Dim strNum As String
    strNum = Mid(strOutLine, Len(strL) + 1, Len(strOutLine) - Len(strL & strR))

    '整数以外の時で1文字でない時
    If Not IsNumeric(strNum) And Len(strNum) > 1 Then
        GetNext = GetNum(strOutLine)
        Exit Function
    End If
    If InStr(1, strNum, "－") > 0 Or InStr(1, strNum, "，") > 0 Or InStr(1, strNum, "．") > 0 Or InStr(1, strNum, "　") > 0 Or _
       InStr(1, strNum, "-") > 0 Or InStr(1, strNum, ",") > 0 Or InStr(1, strNum, ".") > 0 Or InStr(1, strNum, " ") > 0 Then
        GetNext = GetNum(strOutLine)
        Exit Function
    End If

    '全角の かな と カナ の時は イ の次は ィ、カ の次は ガ となるため
    '半角ｶﾅで次の文字を取得して全角に戻す
    Dim blnHiragana As Boolean
    Dim blnWide As Boolean

    '全角ひらがなの時、カタカナに変換
    If StrConv(strNum, vbKatakana) <> strNum Then
        blnHiragana = True
        strNum = StrConv(strNum, vbKatakana)
    End If

    '全角の数字・カタカナの時は半角に変換
    If StrConv(strNum, vbNarrow) <> strNum Then
        blnWide = True
        strNum = StrConv(strNum, vbNarrow)
    End If

    '次の値
    If IsNumeric(strNum) Then
        strNum = CLng(strNum) + 1
    Else
        strNum = Chr(Asc(strNum) + 1)
    End If

    '全角の時は全角に戻す
    If blnWide Then
        strNum = StrConv(strNum, vbWide)
    End If

    'ひらがなの時はひらがなに戻す
    If blnHiragana Then
        strNum = StrConv(strNum, vbHiragana)
    End If

    '両端の文字を連結する
    GetNext = strL & strNum & strR
End Function

'*****************************************************************************
'[概要] 段落番号の次の段落番号を取得する(連番部分が算用数字の時のみ)
'[引数] 例：第１章
'[戻値] 例：第１章 → 第２章
'*****************************************************************************
Private Function GetNum(ByVal strOutLine As String) As String
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True '複数箇所の一致に対応
    objRegExp.Pattern = "[0-9]+|[０-９]+" '全角または半角の数値を含むとき
    If Not objRegExp.Test(strOutLine) Then
        GetNum = strOutLine
        Exit Function
    End If

    Dim strL As String
    Dim strR As String
    Dim strNum As String
    Dim objMatches As Object
    Set objMatches = objRegExp.Execute(strOutLine)

    '算用数字の箇所が複数の時は、一番右側の個所を対象とする
    With objMatches(objMatches.Count - 1)
        strL = Left(strOutLine, .FirstIndex)
        strR = Mid(strOutLine, .FirstIndex + .Length + 1)
        strNum = .Value
    End With

    Dim lngNum As Long
    Dim blnWide As Boolean

    '全角数字の時は半角に変換
    If StrConv(strNum, vbNarrow) <> strNum Then
        blnWide = True
        lngNum = CLng(StrConv(strNum, vbNarrow))
    Else
        lngNum = CLng(strNum)
    End If

    strNum = CStr(lngNum + 1)

    '全角数字の時は全角に戻す
    If blnWide Then
        strNum = StrConv(strNum, vbWide)
    End If

    '両端の文字を連結する
    GetNum = strL & strNum & strR
End Function


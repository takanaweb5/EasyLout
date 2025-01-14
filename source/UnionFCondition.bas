Attribute VB_Name = "UnionFCondition"
Option Explicit
Option Private Module

Private Type TSaveInfo
    NewAppliesTo As Range   '再設定させるセル範囲
    Delete       As Boolean 'True:削除対象
End Type

'*****************************************************************************
'[概要] アクティブシートの条件付き書式を統合する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub MergeFormatConditions()
    Call MergeSameFormatConditions(ActiveSheet)
End Sub

'*****************************************************************************
'[概要] ワークシート内の条件付き書式を統合する
'[引数] 対象のワークシート
'[戻値] なし
'*****************************************************************************
Private Sub MergeSameFormatConditions(ByRef objWorksheet As Worksheet)
    Dim FConditions As FormatConditions
    Set FConditions = objWorksheet.Cells.FormatConditions
    If FConditions.Count = 0 Then
        Exit Sub
    End If
    
    ReDim SaveArray(1 To FConditions.Count) As TSaveInfo
    Dim i As Long
    Dim j As Long
    
    '条件付き書式を後方からLOOPし、統合出来るかどうかの情報をSaveArrayに設定
    For i = FConditions.Count To 1 Step -1
        For j = 1 To i - 1
            If IsSameFormatCondition(FConditions(i), FConditions(j)) Then
                '(i)と(j)が等しければ、後方の(i)を削除して、前方の(j)に統合
                If SaveArray(j).NewAppliesTo Is Nothing Then
                    Set SaveArray(j).NewAppliesTo = Application.Union(FConditions(i).AppliesTo, FConditions(j).AppliesTo)
                Else
                    Set SaveArray(j).NewAppliesTo = Application.Union(FConditions(i).AppliesTo, SaveArray(j).NewAppliesTo)
                End If
                SaveArray(i).Delete = True
            End If
        Next
    Next
    
    '条件付き書式を後方から削除し、前方の条件付き書式に統合
    For i = FConditions.Count To 1 Step -1
        If SaveArray(i).Delete = True Then
            Call FConditions(i).Delete
        Else
            If Not (SaveArray(i).NewAppliesTo Is Nothing) Then
                '条件付き書式の統合
                Call FConditions(i).ModifyAppliesToRange(SaveArray(i).NewAppliesTo)
            End If
        End If
    Next

    'A1,A2,A3 → A1:A3 のように領域を整理
    Dim objWk As Range
    'FormatConditionsを圧縮したため再設定
    Set FConditions = objWorksheet.Cells.FormatConditions
    For i = FConditions.Count To 1 Step -1
        With FConditions(i)
            If .AppliesTo.Areas.Count > 1 Then
                Set objWk = .AppliesTo
                'A1,A2,A3 → A1:A3 のように領域を整理
                Set objWk = Application.Intersect(objWk, objWk)
                '領域が左上から並ぶようにソートする 例:E:F,B:B → B:B,E:F
                Set objWk = SortAreas(objWk)
                Call .ModifyAppliesToRange(objWk)
            End If
        End With
    Next
End Sub

'*****************************************************************************
'[概要] 条件および書式が一致するか判定
'[引数] 比較対象のFormatConditionオブジェクト
'[戻値] True:一致
'*****************************************************************************
Private Function IsSameFormatCondition(ByRef F1 As Object, ByRef F2 As Object) As Boolean
    IsSameFormatCondition = False
    If Not (TypeOf F1 Is FormatCondition) Then
        Exit Function
    End If
    If Not (TypeOf F2 Is FormatCondition) Then
        Exit Function
    End If

'    Select Case F1.Type
'        'セルの値、数式、文字列、期間 のみ判定対象とする　→　FormatConditionはすべて対象にする
'        Case xlCellValue, xlExpression, xlTextString, xlTimePeriod
'        Case Else
'            Exit Function
'    End Select
    
    If IsSameCondition(F1, F2) Then
        IsSameFormatCondition = IsSameFormat(F1, F2)
    End If
End Function

'*****************************************************************************
'[概要] 条件が一致するか判定
'[引数] 比較対象のFormatConditionオブジェクト
'[戻値] True:一致
'*****************************************************************************
Private Function IsSameCondition(ByRef F1 As FormatCondition, ByRef F2 As FormatCondition) As Boolean
    Dim Operator(1 To 2)      As String '次の値に等しい、次の値の間etc
    Dim TextOperator(1 To 2)  As String 'Type=xlTextStringの時、次の値を含む、次の値で始まるetc
    Dim Text(1 To 2)          As String 'Type=xlTextStringの時の文字列
    Dim Formula1_R1C1(1 To 2) As String '数式をR1C1タイプで設定
    Dim Formula2_R1C1(1 To 2) As String '数式をR1C1タイプで設定
    
    'タイプによっては直接判定すると例外となる項目があるため例外を抑制して変数に設定
    On Error Resume Next
    With F1
        Operator(1) = .Operator
        TextOperator(1) = .TextOperator
        Text(1) = .Text
        Formula1_R1C1(1) = ConvertToR1C1Formula(.Formula1, GetTopLeftCell(.AppliesTo))
        Formula2_R1C1(1) = ConvertToR1C1Formula(.Formula2, GetTopLeftCell(.AppliesTo))
    End With
    With F2
        Operator(2) = .Operator
        TextOperator(2) = .TextOperator
        Text(2) = .Text
        Formula1_R1C1(2) = ConvertToR1C1Formula(.Formula1, GetTopLeftCell(.AppliesTo))
        Formula2_R1C1(2) = ConvertToR1C1Formula(.Formula2, GetTopLeftCell(.AppliesTo))
    End With
    On Error GoTo 0
    
    IsSameCondition = (F1.Type = F2.Type) _
                  And (Operator(1) = Operator(2)) _
                  And (TextOperator(1) = TextOperator(2)) _
                  And (Text(1) = Text(2)) _
                  And (Formula1_R1C1(1) = Formula1_R1C1(2)) _
                  And (Formula2_R1C1(1) = Formula2_R1C1(2))
End Function

'*****************************************************************************
'[概要] 一番左上のセルを取得する
'[引数] 条件付き書式の適用範囲
'[戻値] 一番左上のセル
'*****************************************************************************
Private Function GetTopLeftCell(ByRef objRange As Range) As Range
    Dim objArea As Range
    Dim lngRow As Long
    Dim lngCol As Long
    
    '最大値を初期設定
    lngRow = Rows.Count
    lngCol = Columns.Count
    
    For Each objArea In objRange.Areas
        With objArea.Cells(1) '領域ごとの一番左上のセル
            lngRow = WorksheetFunction.Min(lngRow, .Row)
            lngCol = WorksheetFunction.Min(lngCol, .Column)
        End With
    Next
    Set GetTopLeftCell = objRange.Worksheet.Cells(lngRow, lngCol)
End Function

'*****************************************************************************
'[概要] 書式が一致するか判定
'[引数] 比較対象のFormatConditionオブジェクト
'[戻値] True:一致
'*****************************************************************************
Private Function IsSameFormat(ByRef F1 As FormatCondition, ByRef F2 As FormatCondition) As Boolean
    Dim FontBold(1 To 2)      As String 'フォント太字
    Dim FontColor(1 To 2)     As String 'フォント色
    Dim InteriorColor(1 To 2) As String '塗りつぶし色
    Dim NumberFormat(1 To 2)  As String '値の表示形式 例：#,##0
    
    '場合によっては直接判定すると例外となる項目があることを考慮して例外を抑制し変数に設定
    On Error Resume Next
    With F1
        FontBold(1) = .Font.Bold
        FontColor(1) = .Font.Color
        InteriorColor(1) = .Interior.Color
        NumberFormat(1) = .NumberFormat
    End With
    With F2
        FontBold(2) = .Font.Bold
        FontColor(2) = .Font.Color
        InteriorColor(2) = .Interior.Color
        NumberFormat(2) = .NumberFormat
    End With
    On Error GoTo 0
    
    IsSameFormat = (FontBold(1) = FontBold(2)) _
               And (FontColor(1) = FontColor(2)) _
               And (InteriorColor(1) = InteriorColor(2)) _
               And (NumberFormat(1) = NumberFormat(2))
End Function

'*****************************************************************************
'[概要] 領域が左上から並ぶようにソートする 例:E:F,B:B → B:B,E:F
'[引数] 条件付き書式の適用範囲
'[戻値] ソート後の適用範囲
'*****************************************************************************
Private Function SortAreas(ByRef objRange As Range) As Range
    ReDim SortArray(1 To objRange.Areas.Count) As Currency
    Dim i As Long
    Dim j As Long
    
    'Sort対象の配列を作成
    For i = 1 To objRange.Areas.Count
        With objRange.Areas(i)
            '上5桁は列番号、中7桁は行番号、下4桁はIndex
            SortArray(i) = CCur(Format(.Column, "00000") & _
                                Format(.Row, "0000000") & _
                                Format(i, "0000"))
        End With
    Next
    
    'Sort
    Dim Swap As Currency
    For i = objRange.Areas.Count To 1 Step -1
        For j = 1 To i - 1
            If SortArray(j) > SortArray(j + 1) Then
                Swap = SortArray(j)
                SortArray(j) = SortArray(j + 1)
                SortArray(j + 1) = Swap
            End If
        Next j
    Next i
    
    '結果設定
    j = Right(SortArray(1), 4) 'Index=下4桁
    Set SortAreas = objRange.Areas(j)
    For i = 2 To UBound(SortArray)
        j = Right(SortArray(i), 4) 'Index=下4桁
        Set SortAreas = Application.Union(SortAreas, objRange.Areas(j))
    Next
End Function

'*****************************************************************************
'[概要] A1タイプのセル関数をR1C1タイプに変換する
'[引数] 変換前のセル関数、一番左上のセル
'[戻値] 例：A1 → RC
'*****************************************************************************
Private Function ConvertToR1C1Formula(ByVal strFormula As String, ByRef objCell As Range) As String
    ConvertToR1C1Formula = ConvertToR1C1FormulaSub(Application.ConvertFormula(strFormula, xlA1, xlR1C1, , objCell))
End Function

'*****************************************************************************
'[概要] R1C1タイプで相対パスがマイナスの時、プラスに変換する
'       セル関数が、A1=A1048575 と A2=A1 が同じ条件なのに同じ条件と判定されないため
'[引数] 変換前のセル関数
'[戻値] 例：R[-1]C[-1] → R[1048575]C[16383]
'*****************************************************************************
Private Function ConvertToR1C1FormulaSub(ByVal strFormula As String) As String
    Dim reg As Object
    Dim matches As Object
    Dim result As String
    result = strFormula
    
    ' 正規表現オブジェクトの作成
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True    ' 全ての一致を検索
    reg.IgnoreCase = False ' 大文字小文字を区別する
    reg.Pattern = "([RC])\[-([0-9]+)\]"
    If Not reg.Test(strFormula) Then
        ConvertToR1C1FormulaSub = result
        Exit Function
    End If
    
    Set matches = reg.Execute(strFormula)
    
    Dim i As Long
    Dim newValue As Long
    For i = matches.Count - 1 To 0 Step -1
        With matches(i)
            If .SubMatches(0) = "R" Then
                newValue = Rows.Count - .SubMatches(1)
            Else
                newValue = Columns.Count - .SubMatches(1)
            End If
            result = Left(result, .FirstIndex + 2) & newValue & Mid(result, .FirstIndex + .Length)
        End With
    Next
    ConvertToR1C1FormulaSub = result
End Function

'*****************************************************************************
'[概要] Debug用のセル関数
'[引数] objCell:条件付き書式の設定されたセル、n:FormatConditionsの何番目？
'       InfoNo:個別の情報を表示したい時、s(i)のIndexを設定
'[戻値] 例：Type:1 Operator:4 TextOperator:# Text:# Formula1:=0 Formula2:#  Formula1:=0 Formula2:# AppliesTo:A1:A20
'*****************************************************************************
Public Function GetFConditionInfo(objCell As Range, ByVal n As Long, Optional ByVal InfoNo As Long = 0) As String
    Dim objFCondition As Object
    Set objFCondition = objCell.FormatConditions(n)
        
    Dim s(1 To 12)
    Dim i As Long
    For i = 1 To UBound(s)
        s(i) = "#" 'エラーの時
    Next
    
    On Error Resume Next
    With objFCondition
        s(1) = .Type
        s(2) = .Priority
        s(3) = TypeName(objFCondition)
        s(4) = .Operator
        s(5) = .TextOperator
        s(6) = .Text
        s(7) = .Formula1
        s(8) = .Formula2
        s(9) = Application.ConvertFormula(.Formula1, xlA1, xlR1C1, , GetTopLeftCell(.AppliesTo))
        s(10) = Application.ConvertFormula(.Formula2, xlA1, xlR1C1, , GetTopLeftCell(.AppliesTo))
        s(11) = .AppliesTo.AddressLocal(False, False)
        s(12) = GetTopLeftCell(.AppliesTo).AddressLocal(False, False)
    End With
    On Error GoTo 0
    
    If InfoNo > 0 Then
        GetFConditionInfo = s(InfoNo)
    Else
        Dim strMsg As String
        strMsg = "Type:{1} Priority:{2} TypeName:{3} Operator:{4} TextOperator:{5} Text:{6} Formula1:{7} Formula2:{8}  Formula1:{9} Formula2:{10} AppliesTo:{11} TopLeftCell:{12}"
        For i = 1 To UBound(s)
            strMsg = Replace(strMsg, "{" & i & "}", s(i))
        Next
        GetFConditionInfo = strMsg
    End If
End Function




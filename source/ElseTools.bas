Attribute VB_Name = "ElseTools"
Option Explicit
Option Private Module

'罫線情報
Public Type TBorder
    ColorIndex As Long
    LineStyle  As Long
    Weight     As Long
    Color      As Long
End Type

'移動可能かどうかの情報
Private Type TPlacement
    Placement As Byte
    Top       As Double
    Height    As Double
    Left      As Double
    Width     As Double
End Type

'入力支援フォームの情報
Private Type TEditInfo
    Top      As Long
    Left     As Long
    Width    As Long
    Height   As Long
    FontSize As Long
    Zoomed   As Boolean
    WordWarp As Boolean
End Type

Public Const C_CheckErrMsg = 514

'呼び出し先のフォームがロードされているかどうか
Public blnFormLoad As Boolean

Private clsUndoObject  As CUndoObject  'Undo情報
Private lngProcessId As Long   'ヘルプのプロセスID
Private hHelp        As LongPtr   'ヘルプのハンドル
Private udtPlacement() As TPlacement

'*****************************************************************************
'[ 関数名 ]　OpenHelp
'[ 概  要 ]　ヘルプファイルを開く
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub OpenHelp()
    Call OpenHelpPage("Introduction.htm")
End Sub
'*****************************************************************************
'[ 関数名 ]　OpenHelpPage
'[ 概  要 ]　ヘルプファイルの特定のページを開く
'[ 引  数 ]　ページ情報
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub OpenHelpPage(ByVal Bookmark As String)
On Error GoTo ErrHandle
    Dim strHelpPath As String
    Dim strMsg As String
    
    strHelpPath = ThisWorkbook.Path & "\" & "EasyLout.chm"
    If Dir(strHelpPath) = "" Then
        strMsg = "ヘルプファイルが見つかりません。" & vbLf
        strMsg = strMsg & "EasyLout.chmファイルをEasyLout.xlaファイルと同じフォルダにコピーして下さい。"
        Call MsgBox(strMsg, vbExclamation)
        Exit Sub
    End If
    
    'Helpがすべに起動しているかどうか判定
    If hHelp <> 0 Then
        Dim lngExitCode As Long
        Call GetExitCodeProcess(hHelp, lngExitCode)
        If lngExitCode = STILL_ACTIVE Then
            Call AppActivate(lngProcessId)
            Exit Sub
        End If
    End If
    
    'ヘルプファイルをオープンする
    lngProcessId = Shell("hh.exe " & strHelpPath & "::/_RESOURCE/" & Bookmark, vbNormalFocus)
    
    'プロセスのハンドルを取得する
    hHelp = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE Or PROCESS_QUERY_INFORMATION, 0&, lngProcessId)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] オンラインヘルプを開く
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub OpenOnlineHelp()
    Const URL = "http://takana.web5.jp/EasyLout?ref=ap"
    With CreateObject("Wscript.Shell")
        Call .Run(URL)
    End With
End Sub

'*****************************************************************************
'[概要] バージョンを表示する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ShowVersion()
    Call MsgBox("かんたんレイアウト" & vbLf & "  Ver 4.0")
End Sub

'*****************************************************************************
'[ 関数名 ]　MergeCell
'[ 概  要 ]　セルを結合し、値を空白と改行でつないで入力する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MergeCell()
On Error GoTo ErrHandle
    Dim objRange     As Range
    Dim strSelection As String
    Dim lngCalculation As Long
    
    '**************************************
    '初期処理
    '**************************************
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    strSelection = Selection.Address(0, 0)
    lngCalculation = Application.Calculation
    
    '***********************************************
    '選択エリアが１つで結合セルなら対象外
    '***********************************************
    With Range(strSelection)
        If .Areas.Count = 1 And IsOnlyCell(.Cells) Then
            Exit Sub
        End If
    End With
    
    '***********************************************
    '重複領域のチェック
    '***********************************************
    If CheckDupRange(Range(strSelection)) = True Then
        Call MsgBox("選択されている領域に重複があります", vbExclamation)
        Exit Sub
    End If
    
    '**************************************
    '数式のセルの存在チェック
    '**************************************
    Dim objFormulaCell  As Range
    Dim objFormulaCells As Range
    
    'エリアの数だけループ
    For Each objRange In Range(strSelection).Areas
        If WorksheetFunction.CountA(objRange) > 1 Then
            On Error Resume Next
            Set objFormulaCell = objRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            Set objFormulaCells = UnionRange(objFormulaCells, objFormulaCell)
        End If
    Next objRange
    If Not (objFormulaCells Is Nothing) Then
        Call objFormulaCells.Select
        If MsgBox("数式は値に変換されます。よろしいですか？", vbOKCancel + vbQuestion) = vbCancel Then
            Exit Sub
        Else
            Call Range(strSelection).Select
        End If
    End If
    
    '**************************************
    '実行
    '**************************************
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False 'コメントが複数セルにある時の対応
    Application.Calculation = xlManual
    'アンドゥ用に元の状態を保存する
    Call SaveUndoInfo(E_MergeCell, strSelection)
    
    'エリアの数だけループ
    For Each objRange In Range(strSelection).Areas
        Call MergeArea(objRange)
    Next objRange
    
    Call SetOnUndo
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　MergeArea
'[ 概  要 ]　セルを結合し、値を空白と改行でつないで入力する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MergeArea(ByRef objRange As Range)
    Dim i As Long
    Dim j As Long
    Dim strValues   As String
    
    '**************************************
    '値が入力されたセルが１つ以下の時の処理
    '**************************************
    If WorksheetFunction.CountA(objRange) <= 1 Then
        Call objRange.Merge
        Exit Sub
    End If
    
    '**************************************
    '各行の文字列を改行で、各列の文字列を空白で区切って連結する
    '**************************************
    strValues = Replace$(GetRangeText(objRange), vbTab, " ")
    
    '**************************************
    'セルを結合して値を設定
    '**************************************
    '結合セルをすべて解除する
    Call objRange.UnMerge
    '一旦値を消去
    Call objRange.ClearContents
    'セルを結合
    Call objRange.Merge
    '書式を標準に変更(書式が文字列で255文字を超えたら表示出来ないため)
    objRange.NumberFormat = "General"
    objRange.Value = strValues
End Sub

'*****************************************************************************
'[ 関数名 ]　ParseCell
'[ 概  要 ]　セルの結合を解除し、値の１行をセルの１行にする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ParseCell()
On Error GoTo ErrHandle
    Dim lngCalculation As Long
    Dim objRange     As Range

    lngCalculation = Application.Calculation
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    '行方向の結合セルがなければ対象外
    Set objRange = GetMergeRange(Selection, E_MTROW)
    If objRange Is Nothing Then
        Exit Sub
    End If

    '行方向の結合セルがただ１つの結合セルの時
    If objRange.Count = 1 Then
        Call objRange.Select
    End If
    
    '****************************************
    '上揃え、中央揃え、下揃えを選択させる
    '****************************************
    Dim lngAlignment  As Long
    'データの入力されたセルが存在しない時
    If WorksheetFunction.CountA(objRange) = 0 Then
        lngAlignment = E_Top
    Else
        '対象セルが１つだけの時
        If objRange.Count = 1 Then
            'それが数式の時
            If objRange.HasFormula = True Then
                lngAlignment = E_Top
            Else
                If GetStrArray(objRange.Value) < objRange.MergeArea.Rows.Count Then
                    lngAlignment = InputAlignment()
                Else
                    lngAlignment = E_Top
                End If
            End If
        Else
            lngAlignment = InputAlignment()
        End If
    End If
    'キャンセルされた時
    If lngAlignment = E_Cancel Then
        Exit Sub
    End If
    
    'アンドゥ用に元の状態を保存する
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    Call SaveUndoInfo(E_MergeCell, Selection)
    
    '****************************************
    '値のあるセルを１行ごとに分解
    '****************************************
    Dim objCell As Range
    For Each objCell In objRange
        Call ParseOneValueRange(objCell.MergeArea, lngAlignment)
    Next
    
    Call SetOnUndo
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　InputAlignment
'[ 概  要 ]　文字列の縦方向の揃え方を選択させる
'[ 引  数 ]　なし
'[ 戻り値 ]　xlCancel:キャンセル、E_Top:上揃え、E_Center:中央揃え、E_Bottom:下揃え
'*****************************************************************************
Private Function InputAlignment() As Long
    With frmAlignment
        'フォームを表示
        Call .Show

        'キャンセル時
        If blnFormLoad = False Then
            InputAlignment = E_Cancel
            Exit Function
        End If

        InputAlignment = .SelectType
        Call Unload(frmAlignment)
    End With
End Function

'*****************************************************************************
'[ 関数名 ]　ParseOneValueRange
'[ 概  要 ]  値のあるセルを１行ごとに分解
'[ 引  数 ]　値のある結合を解除するセル、行方向の揃え
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ParseOneValueRange(ByRef objRange As Range, ByVal lngAlignment As Long)
    Dim strArray() As String
    Dim lngLine    As Long '行数
    
    '文字列を１行づつ配列に格納
    lngLine = GetStrArray(objRange(1, 1).Value, strArray())
    
    '数式セル または
    '値が１行だけで上揃えの時
    If (objRange(1, 1).HasFormula = True) Or _
       (lngLine = 1 And lngAlignment = E_Top) Then
    Else
        '一旦値を消去
        Call objRange.ClearContents
    End If
    
    '********************
    '結合を解除
    '********************
    Call objRange.Merge(True)
    'セル境界の罫線を消去
    objRange.Borders(xlInsideHorizontal).LineStyle = xlNone
    '「折り返して全体を表示する」のチェックをはずす
    objRange.WrapText = False
    
    '数式セル または　空白セル　または
    '値が１行だけで上揃えの時
    If (objRange(1, 1).HasFormula = True) Or (lngLine = 0) Or _
       (lngLine = 1 And lngAlignment = E_Top) Then
        Exit Sub
    End If
    
    '**************************************
    '１行単位に値設定
    '**************************************
    Dim lngStart   As Long
    Dim strLastCell As String
    Dim i  As Long
    Dim j  As Long
    If lngLine < objRange.Rows.Count Then
        Select Case lngAlignment
        Case E_Top
            lngStart = 1
        Case E_Center
            lngStart = Int((objRange.Rows.Count - lngLine) / 2) + 1
        Case E_Bottom
            lngStart = objRange.Rows.Count - lngLine + 1
        End Select
    Else
        lngStart = 1
    End If
    
    '行の数だけループ
    For i = lngStart To objRange.Rows.Count
        With objRange(i, 1)
            If .NumberFormat <> "@" And StrConv(Left$(strArray(j), 1), vbNarrow) = "=" Then
                .Value = "'" & strArray(j)
            Else
                .Value = strArray(j)
            End If
        End With
        j = j + 1
        If j > UBound(strArray) Then
            Exit For
        End If
    Next i
    
    If lngLine > objRange.Rows.Count Then
        strLastCell = strArray(j - 1)
        For i = j To UBound(strArray)
            strLastCell = strLastCell & vbLf & strArray(i)
        Next
        
        With objRange(objRange.Rows.Count, 1)
            If .NumberFormat <> "@" And StrConv(Left$(strLastCell, 1), vbNarrow) = "=" Then
                .Value = "'" & strLastCell
            Else
                .Value = strLastCell
            End If
        End With
    End If
End Sub
    
'*****************************************************************************
'[概要] 結合を解除して選択範囲で中央寄せ
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub UnmergeCells()
On Error GoTo ErrHandle
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If

    Dim objSelection As Range
    Set objSelection = Selection
    
    'アンドゥ用に元の状態を保存する
    Dim lngCalculation As Long
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    lngCalculation = Application.Calculation
    Application.Calculation = xlManual
    Call SaveUndoInfo(E_MergeCell, Selection)
    
    Dim objRange As Range
    '値の入力されたセルを取得
    With objSelection
        Dim objWk(1 To 2) As Range
        On Error Resume Next
        Set objWk(1) = .SpecialCells(xlCellTypeConstants)
        Set objWk(2) = .SpecialCells(xlCellTypeFormulas)
        Set objRange = UnionRange(objWk(1), objWk(2))
        On Error GoTo ErrHandle
    End With
    If objRange Is Nothing Then
        Call objSelection.UnMerge
        Call SetOnUndo
        Application.DisplayAlerts = True
        Application.Calculation = lngCalculation
        Exit Sub
    End If
    
    '値の入力されたセルのうち結合されたセルを取得
    Set objRange = IntersectRange(objSelection, objRange)
    Set objRange = ArrangeRange(GetMergeRange(objRange))
    If objRange Is Nothing Then
        Call objSelection.UnMerge
        Call SetOnUndo
        Application.DisplayAlerts = True
        Application.Calculation = lngCalculation
        Exit Sub
    End If
    
    With objRange
        .UnMerge
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
    
    Set objRange = ArrangeRange(MinusRange(objSelection, objRange))
    If Not (objRange Is Nothing) Then
        Call objRange.UnMerge
    End If
    
    Call SetOnUndo
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　PasteText
'[ 概  要 ]　値をセルに貼り付ける
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub PasteText()
On Error GoTo ErrHandle
    Dim objCopyRange  As Range
    Dim objSelection  As Range
    Dim strCopyText   As String
    Dim blnOnlyCell   As Boolean
    Dim blnAllCell    As Boolean

    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If

    Set objSelection = Selection

    Select Case Application.CutCopyMode
    Case xlCut
        Call MsgBox("切り取り時は、実行出来ません。", vbExclamation)
        Exit Sub
    Case xlCopy
        On Error Resume Next
        Set objCopyRange = GetCopyRange()
        On Error GoTo 0
        If Not (objCopyRange Is Nothing) Then
            blnOnlyCell = IsOnlyCell(objCopyRange)
            If blnOnlyCell Then
                strCopyText = GetCellText(objCopyRange.Cells(1, 1))
            Else
                strCopyText = MakeCopyText(objCopyRange)
            End If
        Else
            strCopyText = MakeCopyText()
        End If
    Case Else
        strCopyText = GetClipbordText()
        blnAllCell = CheckPasteMode(strCopyText, objSelection)
    End Select
    
    If strCopyText = "" Then
        Exit Sub
    Else
        '改行のCRLF→LF
        strCopyText = Replace$(strCopyText, vbCr, "")
    End If
    
    '選択範囲を貼り付け先に変更
    If blnOnlyCell = False And blnAllCell = False Then
        Set objSelection = GetPasteRange(strCopyText, objSelection)
    End If
    
    Application.ScreenUpdating = False
    
    'アンドゥ用に元の状態を保存する
    Call SaveUndoInfo(E_PasteValue, objSelection)
    If blnOnlyCell Or blnAllCell Then
        objSelection.Value = strCopyText
    Else
        'タブが含まれるかどうか？
        If InStr(1, strCopyText, vbTab, vbBinaryCompare) = 0 Then
            Call PasteRows(strCopyText, objSelection)
        Else
            Call PasteTabText(strCopyText, objSelection)
        End If
    End If
    Call SetOnUndo
    
    '貼り付けた文字列をクリップボードにコピー
    Call SetClipbordText(Replace$(strCopyText, vbLf, vbCrLf))
    
    Call objSelection.Parent.Activate
    Call objSelection.Select
ErrHandle:
End Sub
   
'*****************************************************************************
'[概要] コピー対象の文字列を作成する
'[引数] Copy中の領域
'[戻値] なし
'*****************************************************************************
Private Function MakeCopyText(Optional ByVal objCopyRange As Range = Nothing) As String
On Error GoTo ErrHandle
    If Not (objCopyRange Is Nothing) Then
        MakeCopyText = GetRangeText(objCopyRange)
        Exit Function
    End If
    
    'ワークシートを利用してコピーする
    Dim objSheet As Worksheet
    Set objSheet = ThisWorkbook.Worksheets("Workarea1")
    
    With objSheet
        Call .Range("A1").PasteSpecial(xlPasteValues)
        Call ThisWorkbook.Activate
        Call .Activate
    End With

    MakeCopyText = GetRangeText(Selection)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
Exit Function
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
End Function

'*****************************************************************************
'[ 関数名 ]　PasteRows
'[ 概  要 ]　クリップボードのテキストを1行ごとに分解して貼り付ける
'[ 引  数 ]　コピー文字列、コピー先のセル
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub PasteRows(ByVal strCopyText As String, ByRef objSelection As Range)
    Dim i           As Long
    Dim strArray()  As String
    
    Call GetStrArray(strCopyText, strArray())
        
    '1行毎に分解して貼り付け
    For i = 0 To UBound(strArray)
        objSelection.Rows(i + 1).Value = strArray(i)
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　PasteTabText
'[ 概  要 ]　クリップボードのテキストをタブごとに複数列に渡って貼り付ける
'[ 引  数 ]　クリップボードの文字列、選択範囲
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub PasteTabText(ByVal strText As String, ByRef objSelection As Range)
    Dim i           As Long
    Dim j           As Long
    Dim strArray()  As String
    Dim strCols     As Variant
    
    Call GetStrArray(strText, strArray())
    
    Application.ScreenUpdating = False
    
    '行の数だけループ
    For i = 0 To UBound(strArray)
        strCols = Split(strArray(i), vbTab)
        '列の数だけループ
        For j = 0 To UBound(strCols)
            objSelection(i + 1, j + 1).Value = strCols(j)
        Next j
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　GetPasteRange
'[ 概  要 ]　クリップボードのテキストを貼り付ける領域を再取得
'[ 引  数 ]　貼り付けるText、選択範囲
'[ 戻り値 ]　貼り付け領域
'*****************************************************************************
Public Function GetPasteRange(ByVal strCopyText As String, ByRef objSelection As Range) As Range
    Dim i           As Long
    Dim lngRowCount As Long
    Dim lngMaxCol   As Long
    Dim strArray()  As String
    Dim strCols     As Variant
    
    lngRowCount = GetStrArray(strCopyText, strArray())
    
    If InStr(1, strCopyText, vbTab, vbBinaryCompare) = 0 Then
        lngMaxCol = objSelection.Columns.Count
    Else
        '最大列の取得
        For i = 0 To UBound(strArray)
            strCols = Split(strArray(i), vbTab)
            lngMaxCol = WorksheetFunction.Max(lngMaxCol, UBound(strCols) + 1)
        Next i
    End If
    
    '対象領域を選択しなおす
    Set GetPasteRange = objSelection.Resize(lngRowCount, lngMaxCol)
End Function

'*****************************************************************************
'[ 関数名 ]　CheckPasteMode
'[ 概  要 ]　クリップボードのテキストを貼り付けるモードを判定
'[ 引  数 ]　貼り付けるText、選択範囲
'[ 戻り値 ]　True:すべてのセルに貼り付け、False:セル毎に分解して貼り付け
'*****************************************************************************
Private Function CheckPasteMode(ByVal strCopyText As String, ByRef objSelection As Range) As Boolean
    '選択範囲が複数の領域の時
    If objSelection.Areas.Count > 1 Then
        CheckPasteMode = True
        Exit Function
    End If
    
    'テキストが1行 かつ 1列の時
    If InStr(1, strCopyText, vbLf, vbBinaryCompare) = 0 And _
       InStr(1, strCopyText, vbTab, vbBinaryCompare) = 0 Then
        CheckPasteMode = True
        Exit Function
    End If
        
    '行方向に結合のない単一セル
    If objSelection.Rows.Count = 1 Then
        '行の高さがだいたい2行以上の時
        If objSelection.RowHeight > (objSelection.Font.Size + 2) * 2 Then
            CheckPasteMode = True
            Exit Function
        End If
    End If
    
    '先頭行に行方向の結合がある時
    If objSelection.Rows.Count > 1 Then
        If Not (IntersectRange(ArrangeRange(objSelection.Rows(1)), ArrangeRange(objSelection.Rows(2))) Is Nothing) Then
            CheckPasteMode = True
            Exit Function
        End If
    End If

    CheckPasteMode = False
End Function

'*****************************************************************************
'[ 関数名 ]　CopyText
'[ 概  要 ]　テキストをクリップボードにコピーする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub CopyText()
On Error GoTo ErrHandle
    Dim strText      As String
    Dim objSelection As Range
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = Selection
    
    '選択領域が複数の時
    If objSelection.Areas.Count > 1 Then
        Call objSelection.Copy
        strText = MakeCopyText()
    Else
        strText = GetRangeText(objSelection)
    End If
    
    If strText <> "" Then
        Call SetClipbordText(Replace$(strText, vbLf, vbCrLf))
    End If
    Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　MoveCell
'[ 概  要 ]　結合セルを含む領域を移動する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveCell()
    '選択されているオブジェクトを判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    'IMEをオフにする
    Call SetIMEOff

On Error GoTo ErrHandle
    Dim enmModeType  As EModeType
    Dim objToRange   As Range
    Dim objFromRange As Range
    Dim lngCutCopyMode As Long
    
    'モードを設定する
    lngCutCopyMode = Application.CutCopyMode
    Select Case lngCutCopyMode
    Case xlCopy
        enmModeType = E_Copy
    Case xlCut
        enmModeType = E_CutInsert
    Case Else
        enmModeType = E_Move
    End Select
    
    'コピー先のRangeを設定する
    Set objToRange = Selection
    
    'コピー元のRangeを設定する
    Select Case lngCutCopyMode
    Case xlCopy, xlCut
        Set objFromRange = GetCopyRange()
        If objFromRange Is Nothing Then
            Exit Sub
        End If

        'コピー元とコピー先が同じシートかどうか
        Dim blnSameSheet As Boolean
        blnSameSheet = CheckSameSheet(objFromRange.Worksheet, objToRange.Worksheet)
        If Not blnSameSheet Then
            Call MsgBox("他のシートからコピーする時は、[見たまま貼付け]コマンドを使用してください", vbExclamation)
            Exit Sub
        End If
    Case Else
        Set objFromRange = Selection
    End Select
    
    'チェックを行う
    Dim strErrMsg    As String
    If lngCutCopyMode = xlCut Then
        'コピー元とコピー先が同じシートかどうか
        If CheckSameSheet(objFromRange.Parent, objToRange.Parent) = False Then
            Call MsgBox("切り取り時は、同じシートでないと出来ません。", vbExclamation)
            Exit Sub
        End If
    End If
    strErrMsg = CheckMoveCell(objFromRange)
    If strErrMsg <> "" Then
        Call MsgBox(strErrMsg, vbExclamation)
        Exit Sub
    End If
    strErrMsg = CheckMoveCell(objToRange)
    If strErrMsg <> "" Then
        Call MsgBox(strErrMsg, vbExclamation)
        Exit Sub
    End If
    
    'EXCEL2013以降で起動直後にMoveCellを実行するとボタンが固まる謎の現象を回避するためにSetPixelInfoを呼ぶ
    Call SetPixelInfo
    Call ShowMoveCellForm(enmModeType, objFromRange, objToRange)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　ShowMoveCellForm
'[ 概  要 ]　結合セルを含む領域を移動する
'[ 引  数 ]　enmType:作業タイプ
'            objFromRange:移動(コピー元)の領域
'            objToRange:選択中の領域
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ShowMoveCellForm(ByVal enmType As EModeType, ByRef objFromRange As Range, ByRef objToRange As Range)
    Dim blnCopyObjectsWithCells  As Boolean
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells

On Error GoTo ErrHandle
    'フォームを表示
    With frmMoveCell
        Call .Initialize(enmType, objFromRange, objToRange)
        Call .Show
    End With
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
Exit Sub
ErrHandle:
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
    If blnFormLoad = True Then
        Call Unload(frmMoveCell)
    End If
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　CheckSameSheet
'[ 概  要 ]  objSheet1とobjSheet2が同じシートかどうか判定
'[ 引  数 ]  判定するWorkSheet
'[ 戻り値 ]  True:同じシート
'*****************************************************************************
Public Function CheckSameSheet(ByRef objSheet1 As Worksheet, ByRef objSheet2 As Worksheet) As Boolean
    If objSheet1.Name = objSheet2.Name And _
       objSheet1.Parent.Name = objSheet2.Parent.Name Then
        CheckSameSheet = True
    Else
        CheckSameSheet = False
    End If
End Function

'*****************************************************************************
'[ 関数名 ]　UnSelect
'[ 概  要 ]　選択されたセル領域から、一部の領域を非選択にする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub UnSelect()
    Static strLastSheet   As String '前回の領域の復元用
    Static strLastAddress As String '前回の領域の復元用
On Error GoTo ErrHandle
    Dim objUnSelect  As Range
    Dim objRange As Range
    Dim enmUnselectMode As EUnselectMode
    
    '図形が選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    '取消領域を選択させる
    With frmUnSelect
        '前回の復元用
        Call .SetLastSelect(strLastSheet, strLastAddress)
        
        'フォームを表示
        Call .Show
        'キャンセル時
        If blnFormLoad = False Then
            If Selection.Areas.Count > 1 And strLastAddress = "" Then
                strLastSheet = ActiveSheet.Name
                strLastAddress = GetAddress(Selection)
            End If
            Exit Sub
        End If
        
        enmUnselectMode = .Mode
        Select Case (enmUnselectMode)
        Case E_Unselect, E_Reverse, E_Intersect, E_Union
            Set objUnSelect = .SelectRange
        End Select
        Call Unload(frmUnSelect)
    End With

    Dim objSelection As Range
    Set objSelection = Selection
    
    Select Case (enmUnselectMode)
    Case E_Unselect  '取消し
        Set objRange = MinusRange(objSelection, objUnSelect)
        Set objRange = ReSelectRange(objSelection, objRange)
    Case E_Reverse   '反転
        Set objRange = UnionRange(MinusRange(objSelection, objUnSelect), MinusRange(objUnSelect, objSelection))
    Case E_Intersect '絞り込み
        Set objRange = IntersectRange(objSelection, objUnSelect)
        Set objRange = ReSelectRange(objSelection, objRange)
    Case E_Merge    '結合セルのみ選択
        Set objRange = ArrangeRange(GetMergeRange(objSelection))
        Set objRange = ReSelectRange(objSelection, objRange)
    Case E_Union     '追加
        Dim strAddress As String
        strAddress = objSelection.Address(False, False) & "," & objUnSelect.Address(False, False)
        If Len(strAddress) < 256 Then
            Set objRange = Range(strAddress)
        Else
            Set objRange = UnionRange(objSelection, objUnSelect)
        End If
    End Select
    
    Call objRange.Select
    If Not (MinusRange(Selection, objRange) Is Nothing) Then
        '重複の削除 (結合セルがあると選択領域がおかしくなることがある)
        Call ArrangeRange(objRange).Select
    End If
    
ErrHandle:
    strLastSheet = ActiveSheet.Name
    strLastAddress = GetAddress(Selection)
    If blnFormLoad = True Then
        Call Unload(frmUnSelect)
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　CheckMoveCell
'[ 概  要 ]　MoveCellが可能かどうかのチェック
'[ 引  数 ]　コピー元のRange
'[ 戻り値 ]　エラーの時、エラーメッセージ
'*****************************************************************************
Private Function CheckMoveCell(objSelection As Range) As String
    Dim objWorksheet As Worksheet
    Dim strSelection As String
    
    Application.ScreenUpdating = False
    Set objWorksheet = ActiveSheet
On Error GoTo ErrHandle
    
    '選択エリアが複数なら対象外
    If objSelection.Areas.Count <> 1 Then
        CheckMoveCell = "このコマンドは複数の選択範囲に対して実行できません。"
        Exit Function
    End If
    
    'すべての行が選択なら対象外(動作が非常に遅くなるため)
    If objSelection.Rows.Count = Rows.Count Then
        CheckMoveCell = "このコマンドはすべての行の選択時は実行できません。"
        Exit Function
    End If
    
    'すべての列が選択されて、行方向の結合セルを含む時
    If IsBorderMerged(objSelection) Then
        CheckMoveCell = "結合されたセルの一部を変更することはできません。"
        Exit Function
    End If
ErrHandle:
    Call objWorksheet.Activate
    Application.ScreenUpdating = True
End Function

'*****************************************************************************
'[ 関数名 ]　MoveShape
'[ 概  要 ]　図形を移動またはサイズ変更する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveShape()
    If FPressKey = E_ShiftAndCtrl Then
        Call UnSelect
        Exit Sub
    End If

    '選択されているオブジェクトを判定
    If CheckSelection() <> E_Shape Then
        Call MsgBox("図形が選択されていません")
        Exit Sub
    End If

    'IMEをオフにする
    Call SetIMEOff

On Error GoTo ErrHandle
    'フォームを表示
    Call frmMoveShape.Show
Exit Sub
ErrHandle:
    If blnFormLoad = True Then
        Call Unload(frmMoveShape)
    End If
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　FitShapes
'[ 概  要 ]  選択された図形を枠線にあわせる
'[ 引  数 ]　True:四方を枠線に合わせる、False:左上の位置を枠線に移動
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub FitShapes(ByVal blnSizeChg As Boolean)
On Error GoTo ErrHandle
    Dim i          As Long

    '図形が選択されているか判定
    Select Case (CheckSelection())
    Case E_Range
        Call MsgBox("図形が選択されていません", vbExclamation)
        Exit Sub
    Case E_Other
        Exit Sub
    End Select

    Application.ScreenUpdating = False
    'アンドゥ用に元のサイズを保存する
    Call SaveUndoInfo(E_ShapeSize, Selection.ShapeRange)

    '回転している図形をグループ化する
    Dim objGroups As ShapeRange
    Set objGroups = GroupSelection(Selection.ShapeRange)

    Call FitShapesGrid(objGroups, blnSizeChg)

    '回転している図形のグループ化を解除し元の図形を選択する
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　FitShapesGrid
'[ 概  要 ]  選択された図形を枠線にあわせる
'[ 引  数 ]　objShapeRange：対象図形、blnSizeChg：サイズを変更するかどうか
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub FitShapesGrid(ByRef objShapeRange As ShapeRange, Optional blnSizeChg As Boolean = True)
    Dim objShape   As Shape     '図形
    Dim objRange   As Range

    '図形の数だけループ
    For Each objShape In objShapeRange
        Set objRange = GetNearlyRange(objShape)
        With objShape
            .Top = objRange.Top
            .Left = objRange.Left
            If blnSizeChg = True Then
                If .Height > 0.5 Then
                    .Height = objRange.Height
                Else
                    .Height = 0
                End If
                If .Width > 0.5 Then
                    .Width = objRange.Width
                Else
                    .Width = 0
                End If
            End If
        End With
    Next objShape
End Sub

'*****************************************************************************
'[ 関数名 ]　GroupSelection
'[ 概  要 ]  変更対象の図形の中で回転しているものをグループ化する
'[ 引  数 ]　グループ化前の図形
'[ 戻り値 ]　グループ化後の図形
'*****************************************************************************
Public Function GroupSelection(ByRef objShapes As ShapeRange) As ShapeRange
    Dim i            As Long
    Dim objShape     As Shape
    Dim btePlacement As Byte
    ReDim blnRotation(1 To objShapes.Count) As Boolean
    ReDim lngIDArray(1 To objShapes.Count) As Variant
    
    '図形の数だけループ
    For i = 1 To objShapes.Count
        Set objShape = objShapes(i)
        lngIDArray(i) = objShape.ID
        
        Select Case objShape.Rotation
        Case 90, 270, 180
            blnRotation(i) = True
        End Select
    Next

    '図形の数だけループ
    For i = 1 To objShapes.Count
        If blnRotation(i) = True Then
            Set objShape = GetShapeFromID(lngIDArray(i))
            btePlacement = objShape.Placement
            'サイズと位置が同一のクローンを作成しグループ化する
            With objShape.Duplicate
                .Top = objShape.Top
                .Left = objShape.Left
                If objShape.Top < 0 Then
                    '図形が回転して座標がマイナスになった時ゼロになるため補正する
                    Call .IncrementTop(objShape.Top)
                End If
                If objShape.Left < 0 Then
                    '図形が回転して座標がマイナスになった時ゼロになるため補正する
                    Call .IncrementLeft(objShape.Left)
                End If
                
                '透明にする
                .Fill.Visible = msoFalse
                .Line.Visible = msoFalse
                With GetShapeRangeFromID(Array(.ID, objShape.ID)).Group
                    .AlternativeText = "EL_TemporaryGroup" & i
                    .Placement = btePlacement
                    lngIDArray(i) = .ID
                End With
            End With
        End If
    Next
    
    Set GroupSelection = GetShapeRangeFromID(lngIDArray)
End Function

'*****************************************************************************
'[ 関数名 ]　UnGroupSelection
'[ 概  要 ]  変更対象の図形の中でグループ化したものを元に戻す
'[ 引  数 ]　グループ解除前の図形
'[ 戻り値 ]　グループ解除後の図形
'*****************************************************************************
Public Function UnGroupSelection(ByRef objGroups As ShapeRange) As ShapeRange
    Dim i            As Long
    Dim btePlacement As Byte
    Dim objShape     As Shape
    ReDim blnRotation(1 To objGroups.Count) As Boolean
    ReDim lngIDArray(1 To objGroups.Count) As Variant
    
    '図形の数だけループ
    For i = 1 To objGroups.Count
        Set objShape = objGroups(i)
        lngIDArray(i) = objShape.ID
        
        If Left$(objShape.AlternativeText, 17) = "EL_TemporaryGroup" Then
            blnRotation(i) = True
        End If
    Next

    '図形の数だけループ
    For i = 1 To objGroups.Count
        If blnRotation(i) = True Then
            Set objShape = GetShapeFromID(lngIDArray(i))
            btePlacement = objShape.Placement
            With objShape.Ungroup
                .Item(1).Placement = btePlacement
                Call .Item(2).Delete
                lngIDArray(i) = .Item(1).ID
            End With
        End If
    Next i
    
    Set UnGroupSelection = GetShapeRangeFromID(lngIDArray)
End Function

'*****************************************************************************
'[概要] 図形を連結する
'[引数] True:図形を水平に連結する、False:図形を垂直に連結する
'[戻値] なし
'*****************************************************************************
Public Sub ConnectShapes(ByVal blnHorizontal As Boolean)
On Error GoTo ErrHandle
    Dim i          As Long

    '図形が選択されているか判定
    Select Case (CheckSelection())
    Case E_Range
        Call MsgBox("図形が選択されていません", vbExclamation)
        Exit Sub
    Case E_Other
        Exit Sub
    End Select

    Application.ScreenUpdating = False
    'アンドゥ用に元のサイズを保存する
    Call SaveUndoInfo(E_ShapeSize, Selection.ShapeRange)

    '回転している図形をグループ化する
    Dim objGroups As ShapeRange
    Set objGroups = GroupSelection(Selection.ShapeRange)

    If blnHorizontal Then
        Call ConnectShapesH(objGroups)
    Else
        Call ConnectShapesV(objGroups)
    End If

    '回転している図形のグループ化を解除し元の図形を選択する
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 図形を左右に連結する
'[引数] objShapes:図形
'[戻値] なし
'*****************************************************************************
Public Sub ConnectShapesH(ByRef objShapes As ShapeRange)
    Dim i     As Long
    
    ReDim udtSortArray(1 To objShapes.Count) As TSortArray
    For i = 1 To objShapes.Count
        With udtSortArray(i)
            .Key1 = objShapes(i).Left / DPIRatio
            .Key2 = objShapes(i).Width / DPIRatio
            .Key3 = i
        End With
    Next

    'Let,Widthの順でソートする
    Call SortArray(udtSortArray())

    Dim lngTopLeft    As Long
    lngTopLeft = udtSortArray(1).Key1
    
    For i = 2 To objShapes.Count
        If udtSortArray(i).Key1 > udtSortArray(i - 1).Key1 Then
            lngTopLeft = lngTopLeft + udtSortArray(i - 1).Key2
        End If
        
        objShapes(udtSortArray(i).Key3).Left = lngTopLeft * DPIRatio
    Next
End Sub

'*****************************************************************************
'[概要] 図形を上下に連結する
'[引数] objShapes:図形
'[戻値] なし
'*****************************************************************************
Public Sub ConnectShapesV(ByRef objShapes As ShapeRange)
    Dim i     As Long
    
    ReDim udtSortArray(1 To objShapes.Count) As TSortArray
    For i = 1 To objShapes.Count
        With udtSortArray(i)
            .Key1 = objShapes(i).Top / DPIRatio
            .Key2 = objShapes(i).Height / DPIRatio
            .Key3 = i
        End With
    Next

    'Top,Heightの順でソートする
    Call SortArray(udtSortArray())

    Dim lngTopLeft    As Long
    lngTopLeft = udtSortArray(1).Key1
    
    For i = 2 To objShapes.Count
        If udtSortArray(i).Key1 > udtSortArray(i - 1).Key1 Then
            lngTopLeft = lngTopLeft + udtSortArray(i - 1).Key2
        End If
        
        objShapes(udtSortArray(i).Key3).Top = lngTopLeft * DPIRatio
    Next
End Sub

'*****************************************************************************
'[ 関数名 ]　ChangeTextboxesToCells
'[ 概  要 ]　テキストボックスをセルに変換する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ChangeTextboxesToCells()
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objTextbox   As Shape
    Dim objSelection As ShapeRange
    
    '図形が選択されているか判定
    Select Case (CheckSelection())
    Case E_Range
        Call MsgBox("テキストボックスが選択されていません", vbExclamation)
        Exit Sub
    Case E_Other
        Exit Sub
    End Select
    
    Set objSelection = Selection.ShapeRange
    ReDim lngIDArray(1 To objSelection.Count) As Variant
    
    '**************************************
    '選択された図形のチェック
    '**************************************
    '図形の数だけループ
    For i = 1 To objSelection.Count 'For each構文だとExcel2007で型違いとなる(たぶんバグ)
        Set objTextbox = objSelection(i)
        lngIDArray(i) = objTextbox.ID
        'テキストボックスだけが選択されているか判定
        If CheckTextbox(objTextbox) = False Then
            Call MsgBox("テキストボックス以外は選択しないで下さい", vbExclamation)
            Exit Sub
        End If
    Next

    '**************************************
    '変換されるセルをobjRangeに設定
    '**************************************
    Dim objRange     As Range
    Dim blnNotMatch  As Boolean
    
    'テキストボックスの数だけループ
    For i = 1 To objSelection.Count 'For each構文だとExcel2007で型違いとなる(たぶんバグ)
        Set objTextbox = objSelection(i)
        With GetNearlyRange(objTextbox)
            If IsBorderMerged(.Cells) Then
                Call MsgBox("結合されたセルの一部を変更することはできません", vbExclamation)
                Call objTextbox.Select
                Exit Sub
            End If
        
            If IntersectRange(objRange, .Cells) Is Nothing Then
                Set objRange = UnionRange(objRange, .Cells)
            Else
                Call objTextbox.Select
                Call MsgBox("変換されるセルに重複があります", vbExclamation)
                Exit Sub
            End If
            
            'テキストボックスが枠線と一致しているか判定
            If .Top = objTextbox.Top And .Left = objTextbox.Left And _
               .Width = objTextbox.Width And .Height = objTextbox.Height Then
            Else
                blnNotMatch = True
            End If
        End With
    Next
    
    If blnNotMatch = True Then
        Call objRange.Select
        If MsgBox("テキストボックスがグリッド(枠線)にあっていません。" & vbLf & _
                  "位置・サイズの最も近いセルに変換します。" & vbLf & _
                  "よろしいですか？", vbOKCancel + vbQuestion) = vbCancel Then
            Exit Sub
        End If
    End If
    
    '**************************************
    'アンドゥ用に元の状態を保存する
    '**************************************
    Dim strSelectAddress  As String
    strSelectAddress = objRange.Address
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False 'コメントがあると警告が出る時がある
    
    Call SaveUndoInfo(E_TextToCell, objSelection, objRange)
    
    '**************************************
    'テキストボックスをセルに変換する
    '**************************************
    Dim objShapeRange As ShapeRange
    Set objShapeRange = GetShapeRangeFromID(lngIDArray)
    'テキストボックスの数だけループ
    For i = 1 To objShapeRange.Count 'For each構文だとExcel2007で型違いとなる(たぶんバグ)
        Set objTextbox = objShapeRange(i)
        Set objRange = GetNearlyRange(objTextbox)
        Call ChangeTextboxToCell(objTextbox, objRange)
    Next
    
    'テキストボックスを削除
    Call objShapeRange.Delete
    
    '変換されたセルを選択する
    Call Range(strSelectAddress).Select
    Call SetOnUndo
    Application.DisplayAlerts = True
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　CheckTextbox
'[ 概  要 ]　Shapeがテキストボックスかどうか判定する
'[ 引  数 ]　判定するShape
'[ 戻り値 ]　True:テキストボックス
'*****************************************************************************
Private Function CheckTextbox(ByRef objShape As Shape) As Boolean
On Error GoTo ErrHandle
    '回転しているか
    If objShape.Rotation <> 0 Then
        Exit Function
    End If
    
    If (objShape.Type <> msoTextBox) And (objShape.Type <> msoAutoShape) Then
        Exit Function
    End If
    
'    If (TypeOf objShape.DrawingObject Is TextBox) Or _
'       (TypeOf objShape.DrawingObject Is Rectangle) Then '四角形
'    Else
'       Exit Function
'    End If
    
    If IsNull(objShape.TextFrame.Characters.Count) Then
        Exit Function
    End If
    
    CheckTextbox = True
Exit Function
ErrHandle:
    CheckTextbox = False
End Function

'*****************************************************************************
'[ 関数名 ]　ChangeTextboxToCell
'[ 概  要 ]　テキストボックスをセルに変換する
'[ 引  数 ]　テキストボックス
'[ 戻り値 ]　セル
'*****************************************************************************
Private Sub ChangeTextboxToCell(ByRef objTextbox As Shape, ByRef objRange As Range)
    'セルを結合する
    Call objRange.UnMerge
    Call objRange.Merge
    
    '文字列の設定
    objRange(1, 1).Value = objTextbox.TextFrame2.TextRange.Text
    
    '縦書き or 横書き
    Select Case objTextbox.TextFrame2.Orientation
    Case msoTextOrientationDownward
        objRange.Orientation = xlDownward
    Case msoTextOrientationUpward
        objRange.Orientation = xlUpward
    Case msoTextOrientationHorizontalRotatedFarEast, _
         msoTextOrientationVerticalFarEast, _
         msoTextOrientationVertical
        objRange.Orientation = xlVertical '縦書き
    Case Else
        objRange.Orientation = xlHorizontal
    End Select

'    On Error Resume Next
'    With objTextbox.TextFrame.Font
'        objRange.Font.Name = .Name
'        objRange.Font.size = .size
'        objRange.Font.FontStyle = .FontStyle
'    End With
'    On Error GoTo 0
    
'    '配置の設定
'    With objTextbox.TextFrame
'        On Error Resume Next
'        If .Orientation = msoTextOrientationVertical Then
'            objRange.Orientation = xlVertical       '縦書き
'        End If
'        On Error GoTo 0
'
'        '横位置設定
'        On Error Resume Next
'        Select Case .HorizontalAlignment
'        Case xlHAlignLeft
'            objRange.HorizontalAlignment = xlLeft
'        Case xlHAlignCenter
'            objRange.HorizontalAlignment = xlCenter
'        Case xlHAlignRight
'            objRange.HorizontalAlignment = xlRight
'        Case xlHAlignDistributed
'            objRange.HorizontalAlignment = xlDistributed
'        Case xlHAlignJustify
'            objRange.HorizontalAlignment = xlJustify
'        End Select
'        On Error GoTo 0
'
'        '縦位置設定
'        On Error Resume Next
'        Select Case .VerticalAlignment
'        Case xlVAlignTop
'            objRange.VerticalAlignment = xlTop
'        Case xlVAlignCenter
'            objRange.VerticalAlignment = xlCenter
'        Case xlVAlignBottom
'            objRange.VerticalAlignment = xlBottom
'        Case xlVAlignDistributed
'            objRange.VerticalAlignment = xlDistributed
'        Case xlVAlignJustify
'            objRange.VerticalAlignment = xlJustify
'        End Select
'        On Error GoTo 0
'    End With

    '罫線の設定
    If objTextbox.Line.Visible <> msoFalse Then
        On Error Resume Next
        With objRange '斜線
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
        End With
        With objRange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objRange.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        On Error GoTo 0
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　ChangeCellsToTextboxes
'[ 概  要 ]　セルをテキストボックスに変換する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ChangeCellsToTextboxes()
On Error GoTo ErrHandle
    Dim objRange        As Range
    Dim objSelection    As Range
    Dim strTextboxes()  As Variant
    Dim blnClear        As Boolean
    Dim i As Long
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Dim strMsg As String
    strMsg = "元のセルの内容をクリアしますか？" & vbLf
    strMsg = strMsg & "　「 はい 」････ セルの値を消去し結合は解除する" & vbCrLf
    strMsg = strMsg & "　「いいえ」････ セルをそのままにしておく"
    Select Case MsgBox(strMsg, vbYesNoCancel + vbQuestion + vbDefaultButton2)
    Case vbYes
        blnClear = True
    Case vbNo
        blnClear = False
    Case vbCancel
        Exit Sub
    End Select
    
    '図形が非表示かどうか判定
    If ActiveWorkbook.DisplayDrawingObjects = xlHide Then
        ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
    End If
    
    Set objSelection = Selection
    
    Application.ScreenUpdating = False
    'アンドゥ用に元の状態を保存する
    Call SaveUndoInfo(E_CellToText, objSelection)
    
    'セル単位でループ
    For Each objRange In objSelection
        '結合セルの時、左上のセル以外は対象外
        If objRange(1, 1).Address = objRange.MergeArea(1, 1).Address Then
            i = i + 1
            ReDim Preserve strTextboxes(1 To i)
            strTextboxes(i) = ChangeCellToTextbox(objRange.MergeArea).Name
        End If
    Next objRange
        
    If blnClear Then
        With objSelection
            '元の領域をクリア
            Call .Clear
            If Cells(Rows.Count - 2, Columns.Count - 2).MergeCells = False Then
                'シート上の標準的な書式に設定
                Call Cells(Rows.Count - 2, Columns.Count - 2).Copy(objSelection)
                Call .ClearContents
            End If
        End With
    End If
    
    '作成したテキストボックスを選択
    Call ActiveSheet.Shapes.Range(strTextboxes).Select
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] セルをテキストボックスに変換する
'[引数] セル
'[戻値] テキストボックス
'*****************************************************************************
Private Function ChangeCellToTextbox(ByRef objRange As Range) As Shape
    Dim objTextbox As Shape
    Dim objCell    As Range
    
    With objRange
        'テキストボックス作成
        Set objTextbox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
    End With
    
    Set objCell = objRange(1, 1)
    
    'フォントと文字列の設定
    With objTextbox.TextFrame2.TextRange.Font
        .NameComplexScript = objCell.Font.Name
        .NameFarEast = objCell.Font.Name
        .Name = objCell.Font.Name
        .Size = objCell.Font.Size
    End With
    With objTextbox.TextFrame2.TextRange
        .Text = objCell.Value
    End With
    
    '配置の設定
    With objTextbox.TextFrame2
        Select Case objCell.Orientation
        Case xlDownward, xlUpward, xlVertical, 90, -90
            .Orientation = msoTextOrientationVerticalFarEast '縦書き
        End Select
    End With
    
    If objCell.Orientation = xlVertical Then '縦書き
        With objTextbox.TextFrame
            '横位置設定
            Select Case objCell.HorizontalAlignment
            Case xlLeft, xlJustify
                .VerticalAlignment = xlVAlignBottom
            Case xlRight
                .VerticalAlignment = xlVAlignTop
            Case Else
                .VerticalAlignment = xlHAlignCenter
            End Select
            '縦位置設定
            Select Case objCell.VerticalAlignment
            Case xlJustify, xlTop
                .HorizontalAlignment = xlHAlignLeft
            Case xlBottom
                .HorizontalAlignment = xlHAlignRight
            Case Else
                .HorizontalAlignment = xlVAlignCenter
            End Select
        End With
    Else '横書き
        With objTextbox.TextFrame2.TextRange.ParagraphFormat
            '横位置設定
            Select Case objCell.HorizontalAlignment
            Case xlGeneral, xlLeft, xlJustify
                .Alignment = msoAlignLeft
            Case xlRight
                .Alignment = msoAlignRight
            Case Else
                .Alignment = msoAlignCenter
            End Select
        End With
        With objTextbox.TextFrame2
            '縦位置設定
            Select Case objCell.VerticalAlignment
            Case xlJustify, xlTop
                .VerticalAnchor = msoAnchorTop
            Case xlBottom
                .VerticalAnchor = msoAnchorBottom
            Case Else
                .VerticalAnchor = msoAnchorMiddle
            End Select
        End With
    End If
    
    '線の設定
    With objTextbox.Line
        .Weight = DPIRatio
        .DashStyle = msoLineSolid
        .Style = msoLineSingle
        .ForeColor.Rgb = 0
        .BackColor.Rgb = Rgb(255, 255, 255)
    End With
    With objRange
        If .Borders(xlEdgeTop).LineStyle = xlNone Or _
           .Borders(xlEdgeLeft).LineStyle = xlNone Or _
           .Borders(xlEdgeBottom).LineStyle = xlNone Or _
           .Borders(xlEdgeRight).LineStyle = xlNone Then
            objTextbox.Line.Visible = msoFalse
        Else
            objTextbox.Line.Visible = msoTrue
        End If
    End With
    
    Set ChangeCellToTextbox = objTextbox
End Function

'*****************************************************************************
'[概要] コメントをテキストボックスに変換する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ChangeCommentsToTextboxes()
    Dim objRange As Range
    Set objRange = GetComments()
    If objRange Is Nothing Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    
    'アンドゥ用に元の状態を保存する
    If CheckSelection() = E_Range Then
        Call SaveUndoInfo(E_CommentToText, Selection)
    Else
        Call SaveUndoInfo(E_CommentToText, objRange)
    End If
    
    Dim strTextboxes()  As Variant
    Dim i As Long

    '図形が非表示かどうか判定
    If ActiveWorkbook.DisplayDrawingObjects = xlHide Then
        ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
    End If
    
    Dim objCell  As Range
    For Each objCell In objRange
        If objCell.MergeArea(1).Address = objCell.Address Then
            i = i + 1
            ReDim Preserve strTextboxes(1 To i)
            strTextboxes(i) = ChangeCommentToTextbox(objCell).Name
        End If
    Next
    Call objRange.ClearComments
    
    '作成したテキストボックスを選択
    Call ActiveSheet.Shapes.Range(strTextboxes).Select
    Call SetOnUndo
End Sub

'*****************************************************************************
'[概要] コメントをテキストボックスに変換する
'[引数] コメントオブジェクト
'[戻値] テキストボックス
'*****************************************************************************
Private Function ChangeCommentToTextbox(ByRef objCell As Range) As Shape
    Dim objTextbox As Shape
    Dim objComment As Comment
    Set objComment = objCell.Comment
    
    With objComment.Shape
        'テキストボックス作成
        Set objTextbox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
    End With
    
    'フォントと文字列の設定
    Dim objFont As Font
    Set objFont = objComment.Shape.DrawingObject.Font
    With objTextbox.TextFrame2.TextRange.Font
        .NameComplexScript = objFont.Name
        .NameFarEast = objFont.Name
        .Name = objFont.Name
        .Size = objFont.Size
    End With
    With objTextbox.TextFrame2.TextRange
        .Text = objComment.Text
    End With
    
    '線の設定
    With objTextbox.Line
        .Weight = DPIRatio
        .DashStyle = msoLineSolid
        .Style = msoLineSingle
        .ForeColor.Rgb = 0
        .BackColor.Rgb = Rgb(255, 255, 255)
        .Visible = msoTrue
    End With
    
    With objTextbox
        .Title = "コメント"
        .AlternativeText = objCell.Address(0, 0)
    End With

    Set ChangeCommentToTextbox = objTextbox
End Function

'*****************************************************************************
'[概要] テキストボックスをコメントに変換する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ChangeTextboxesToComments()
On Error GoTo ErrHandle
    Dim objSelection As ShapeRange
    Dim objRange As Range
    Set objSelection = Selection.ShapeRange
    ReDim lngIDArray(1 To objSelection.Count) As Variant
    
    '図形の数だけループ
    Dim i As Long
    Dim objTextbox As Shape
    For i = 1 To objSelection.Count 'For each構文だとExcel2007で型違いとなる(たぶんバグ)
        Set objTextbox = objSelection(i)
        lngIDArray(i) = objTextbox.ID
        Set objRange = UnionRange(objRange, Range(objTextbox.AlternativeText))
    Next
    
    '**************************************
    'アンドゥ用に元の状態を保存する
    '**************************************
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False 'コメントがあると警告が出る時がある
    
    Call SaveUndoInfo(E_TextToComment, objSelection, objRange)
    
    '**************************************
    'テキストボックスをコメントに変換する
    '**************************************
    Dim objShapeRange As ShapeRange
    Set objShapeRange = GetShapeRangeFromID(lngIDArray)
    'テキストボックスの数だけループ
    For i = 1 To objShapeRange.Count 'For each構文だとExcel2007で型違いとなる(たぶんバグ)
        Set objTextbox = objShapeRange(i)
        Call ChangeTextboxToComment(objTextbox)
    Next
    
    'テキストボックスを削除
    Call objShapeRange.Delete
    
    '変換されたセルを選択する
    Call objRange.Select
    Call SetOnUndo
    Application.DisplayAlerts = True
Exit Sub
ErrHandle:
End Sub

'*****************************************************************************
'[概要] コメントから変換されたテキストボックスだけが選択されているかどうか
'[引数] なし
'[戻値] True or False
'*****************************************************************************
Private Function ChangeTextboxToComment(ByRef objTextbox As Shape) As Boolean
    Dim objCell As Range
    Set objCell = Range(objTextbox.AlternativeText)
    Call objCell.Select
    Call objCell.ClearComments
    Call objCell.AddComment(objTextbox.TextFrame2.TextRange.Text)
End Function

'*****************************************************************************
'[概要] コメントから変換されたテキストボックスだけが選択されているかどうか
'[引数] なし
'[戻値] True or False
'*****************************************************************************
Public Function IsSelectCommentTextbox() As Boolean
On Error GoTo ErrHandle
    '図形が選択されているか判定
    If CheckSelection() <> E_Shape Then
        Exit Function
    End If
    
    Dim objSelection As ShapeRange
    Set objSelection = Selection.ShapeRange
    
    '図形の数だけループ
    Dim i As Long
    For i = 1 To objSelection.Count 'For each構文だとExcel2007で型違いとなる(たぶんバグ)
        If Not IsCommentTextbox(objSelection(i)) Then
            Exit Function
        End If
    Next
    IsSelectCommentTextbox = True
Exit Function
ErrHandle:
End Function

'*****************************************************************************
'[概要] オートシェイプがコメントから変換されたテキストボックスかどうか判定
'[引数] オートシェイプ
'[戻値] True:コメントから変換されたテキストボックス
'*****************************************************************************
Private Function IsCommentTextbox(ByRef objShape As Shape) As Boolean
On Error GoTo ErrHandle
    If objShape.Title = "コメント" Then
        Dim objRegExp As Object
        Set objRegExp = CreateObject("VBScript.RegExp")
        objRegExp.Global = False
        objRegExp.Pattern = "[A-Z]{1,3}[1-9][0-9]{0,6}"
        If objRegExp.Test(objShape.AlternativeText) Then
            IsCommentTextbox = True
        End If
    End If
ErrHandle:
End Function

'*****************************************************************************
'[概要] コメントを入力規則に変換する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ChangeCommentsToInputRules()
    Dim objRange As Range
    Set objRange = GetComments()
    If objRange Is Nothing Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    
    'アンドゥ用に元の状態を保存する
    If CheckSelection() = E_Range Then
        Call SaveUndoInfo(E_CommentToRule, Selection)
    Else
        Call SaveUndoInfo(E_CommentToRule, objRange)
    End If
    
    Call objRange.Validation.Delete
    Dim objCell  As Range
    For Each objCell In objRange
        If objCell.MergeArea(1).Address = objCell.Address Then
            Call ChangeCommentToInputRule(objCell)
        End If
    Next
    Call objRange.ClearComments
    Call objRange.Select
    Call SetOnUndo
End Sub

'*****************************************************************************
'[概要] コメントを入力規則に変換する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ChangeCommentToInputRule(ByRef objCell As Range)
    Call objCell.Select
    With objCell.Validation
        Call .Delete
        Call .Add(xlValidateInputOnly)
        .ShowInput = True
        .InputTitle = ""
        .InputMessage = objCell.MergeArea(1).Comment.Text
        .ShowError = False
        .ErrorTitle = ""
        .ErrorMessage = ""
        .IgnoreBlank = True
        .InCellDropdown = True
        .IMEMode = xlIMEModeNoControl
    End With
End Sub

'*****************************************************************************
'[概要] 選択されているセル範囲のコメントのあるセル範囲を取得
'[引数] なし
'[戻値] コメントのあるセル範囲
'*****************************************************************************
Public Function GetComments() As Range
    Select Case CheckSelection()
    Case E_Shape
        If Selection.ShapeRange.Type = msoComment Then
            Set GetComments = GetCommentsCell(Selection.ShapeRange.ID)
            Exit Function
        End If
    End Select
    
    Dim objRange As Range
    On Error Resume Next
    Set objRange = Selection.SpecialCells(xlCellTypeComments)
    Set GetComments = IntersectRange(Selection, objRange)
End Function

'*****************************************************************************
'[概要] 選択されているセル範囲のコメントのあるセル範囲を取得
'[引数] なし
'[戻値] コメントのあるセル範囲
'*****************************************************************************
Public Function GetCommentsCell(ByVal ID As Long) As Range
    Dim objRange As Range
    On Error Resume Next
    Set objRange = Cells.SpecialCells(xlCellTypeComments)
    On Error GoTo 0
    
    Dim objCell As Range
    For Each objCell In objRange
        If objCell.Comment.Shape.ID = ID Then
            Set GetCommentsCell = objCell(1)
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[概要] 入力規則をコメントに変換する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ChangeInputRulesToComments()
    Dim objRange As Range
    Set objRange = GetInputRules
    If objRange Is Nothing Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    
    'アンドゥ用に元の状態を保存する
    Call SaveUndoInfo(E_RuleToComment, Selection)
    
    Call objRange.ClearComments
    Dim objCell  As Range
    For Each objCell In objRange
        If objCell.MergeArea(1).Address = objCell.Address Then
            Call ChangeInputRuleToComment(objCell)
        End If
    Next
    Call objRange.Validation.Delete
    Call Cells(Rows.Count, Columns.Count).Select
    Call objRange.Select
    Call SetOnUndo
End Sub

'*****************************************************************************
'[概要] 入力規則をコメントに変換する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ChangeInputRuleToComment(ByRef objCell As Range)
    Dim strComment As String
    Dim strInputTitle As String
    Dim strInputMessage As String
    
    With objCell.Validation
        On Error Resume Next
        strInputTitle = .InputTitle
        strInputMessage = .InputMessage
        On Error GoTo 0
    End With
                
    If strInputTitle <> "" Then
        strComment = strInputTitle
        If strInputMessage <> "" Then
            strComment = strComment & vbLf & strInputMessage
        End If
    Else
        strComment = strInputMessage
    End If
    
    Call objCell.AddComment
    objCell.Comment.Visible = True
    Call objCell.Comment.Text(strComment)
End Sub

'*****************************************************************************
'[概要] 選択されているセル範囲の入力規則のあるセル範囲を取得
'[引数] blnCheckOnly:チェックのみする時（Enabled設定時の高速化）
'[戻値] 入力規則のあるセル範囲
'*****************************************************************************
Public Function GetInputRules(Optional ByVal blnCheckOnly As Boolean = False) As Range
    If CheckSelection <> E_Range Then
        Exit Function
    End If
    
    Dim objRange As Range
    On Error Resume Next
    Set objRange = Cells.SpecialCells(xlCellTypeAllValidation)
'    On Error GoTo 0
    
    Set objRange = IntersectRange(Selection, objRange)
    If objRange Is Nothing Then
        Exit Function
    End If
    Dim objCell As Range
    For Each objCell In objRange
        With objCell.Validation
            If .ShowInput = True Then
                If .InputMessage <> "" Or .InputTitle <> "" Then
                    If blnCheckOnly Then
                        Set GetInputRules = objCell
                    Else
                        Set GetInputRules = UnionRange(GetInputRules, objCell)
                    End If
                End If
            End If
        End With
    Next
End Function

'*****************************************************************************
'[ 関数名 ]　HideShapes
'[ 概  要 ]  ブック内のすべての図形を非表示にする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub HideShapes(ByVal blnHide As Boolean)
    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
    With ActiveWorkbook
        If blnHide Then
            .DisplayDrawingObjects = xlHide
        Else
            .DisplayDrawingObjects = xlAll
        End If
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　GetBorder
'[ 概  要 ]　BorderオブジェクトをTBorder構造体に代入する
'[ 引  数 ]　Borderオブジェクト
'[ 戻り値 ]　TBorder構造体
'*****************************************************************************
Public Function GetBorder(ByRef objBorder As Border) As TBorder
    With objBorder
        GetBorder.LineStyle = .LineStyle
        GetBorder.ColorIndex = .ColorIndex
        GetBorder.Weight = .Weight
        GetBorder.Color = .Color
    End With
End Function

'*****************************************************************************
'[ 関数名 ]　SetBorder
'[ 概  要 ]　TBorder構造体をBorderオブジェクトに設定する
'[ 引  数 ]　TBorder構造体
'            Borderオブジェクト
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetBorder(ByRef udtBorder As TBorder, ByRef objBorder As Border)
    With objBorder
        If .LineStyle <> udtBorder.LineStyle Then
            .LineStyle = udtBorder.LineStyle
        End If
        If .ColorIndex <> udtBorder.ColorIndex Then
            .ColorIndex = udtBorder.ColorIndex
        End If
        If .Weight <> udtBorder.Weight Then
            .Weight = udtBorder.Weight
        End If
        If .Color <> udtBorder.Color Then
            .Color = udtBorder.Color
        End If
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　SaveUndoInfo
'[ 概  要 ]　Undo情報を保存する
'[ 引  数 ]　enmType:Undoタイプ、objObject:復元対象の選択されたオブジェクト
'            varInfo:付加情報
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SaveUndoInfo(ByVal enmType As EUndoType, ByRef vSelection As Variant, Optional ByVal varInfo As Variant = Nothing)
    'すでにインスタンスが存在する時は、Rollbackに対応するためNewしない
    If clsUndoObject Is Nothing Then
        Set clsUndoObject = New CUndoObject
    End If
    
    'オートフィルタが設定されている時は、Undo不可にする
    If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
        Call clsUndoObject.SaveUndoInfo(enmType, vSelection, varInfo)
    Else
        Call clsUndoObject.SaveUndoInfo(E_FilterERR, vSelection, varInfo)
    End If
End Sub

'*****************************************************************************
'[概要] ApplicationオブジェクトのOnUndoイベントを設定
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub SetOnUndo()
    'キーの再定義 Excelのバグ?でキーが無効になることがあるため
    Call SetKeys

    Call clsUndoObject.SetOnUndo
    Call Application.OnTime(Now(), "SetOnRepeat")
End Sub

'*****************************************************************************
'[概要] ApplicationオブジェクトのOnRepeatイベントを設定
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SetOnRepeat()
    Call Application.OnRepeat("繰り返し", "OnRepeat")
End Sub

'*****************************************************************************
'[概要] 繰返しクリック時
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub OnRepeat()
    If FParam = "" Then
        Call Application.Run(FCommand)
    Else
        Call Application.Run(FCommand, FParam)
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　ExecUndo
'[ 概  要 ]　Undoを実行する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ExecUndo()
On Error GoTo ErrHandle
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Call clsUndoObject.ExecUndo
    Set clsUndoObject = Nothing
    Application.DisplayAlerts = True
    Call Application.OnRepeat("", "")
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Set clsUndoObject = Nothing
    Call MsgBox(Err.Description, vbExclamation)
    Call Application.OnRepeat("", "")
End Sub

'*****************************************************************************
'[ 関数名 ]　SetPlacement
'[ 概  要 ]　ShapeのPlacementプロパティを変更する
'　　　　　　セルにあわせてShapeの位置とサイズを変更させないようにする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetPlacement()
On Error GoTo ErrHandle
    Dim i          As Long
    Dim lngDisplay As Long
    
    If ActiveSheet.Shapes.Count = 0 Then
        Exit Sub
    End If
    
    lngDisplay = ActiveWorkbook.DisplayDrawingObjects
    ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
    
    ReDim udtPlacement(1 To ActiveSheet.Shapes.Count)
    For i = 1 To ActiveSheet.Shapes.Count
        With ActiveSheet.Shapes(i)
            udtPlacement(i).Placement = .Placement
            .Placement = xlFreeFloating
            If .Type = msoComment Then
                udtPlacement(i).Top = .Top
                udtPlacement(i).Height = .Height
                udtPlacement(i).Left = .Left
                udtPlacement(i).Width = .Width
            End If
        End With
    Next i
ErrHandle:
    ActiveWorkbook.DisplayDrawingObjects = lngDisplay
End Sub

'*****************************************************************************
'[ 関数名 ]　ResetPlacement
'[ 概  要 ]　ShapeのPlacementプロパティを元に戻す
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ResetPlacement()
On Error GoTo ErrHandle
    Dim i          As Long
    Dim lngDisplay As Long
    
    If ActiveSheet.Shapes.Count = 0 Then
        Exit Sub
    End If
    
    lngDisplay = ActiveWorkbook.DisplayDrawingObjects
    ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes

    For i = 1 To ActiveSheet.Shapes.Count
        With ActiveSheet.Shapes(i)
            .Placement = udtPlacement(i).Placement
            If .Type = msoComment Then
                .Top = udtPlacement(i).Top
                .Height = udtPlacement(i).Height
                .Left = udtPlacement(i).Left
                .Width = udtPlacement(i).Width
            End If
        End With
    Next i
ErrHandle:
    Erase udtPlacement()
    ActiveWorkbook.DisplayDrawingObjects = lngDisplay
End Sub

'*****************************************************************************
'[ 関数名 ]　GetShapeRangeFromID
'[ 概  要 ]　ShpesオブジェクトのIDからShapeRangeオブジェクトを取得
'[ 引  数 ]　IDの配列
'[ 戻り値 ]　ShapeRangeオブジェクト
'*****************************************************************************
Public Function GetShapeRangeFromID(ByRef lngID As Variant) As ShapeRange
    Dim i As Long
    Dim j As Long
    Dim lngShapeID As Long
    ReDim lngArray(LBound(lngID) To UBound(lngID)) As Variant
    
    For j = 1 To ActiveSheet.Shapes.Count
        lngShapeID = ActiveSheet.Shapes(j).ID
        For i = LBound(lngID) To UBound(lngID)
            If lngShapeID = lngID(i) Then
                lngArray(i) = j
                Exit For
            End If
        Next
    Next
    
    Set GetShapeRangeFromID = ActiveSheet.Shapes.Range(lngArray)
End Function

'*****************************************************************************
'[ 関数名 ]　GetShapeFromID
'[ 概  要 ]　ShapeオブジェクトのIDからShapeオブジェクトを取得
'[ 引  数 ]　ID
'[ 戻り値 ]　Shapeオブジェクト
'*****************************************************************************
Public Function GetShapeFromID(ByVal lngID As Long) As Shape
    Dim j As Long
    Dim lngIndex As Long
        
    For j = 1 To ActiveSheet.Shapes.Count
        If ActiveSheet.Shapes(j).ID = lngID Then
            lngIndex = j
            Exit For
        End If
    Next j
    
    Set GetShapeFromID = ActiveSheet.Shapes.Range(j).Item(1)
End Function

'*****************************************************************************
'[ 関数名 ]　OnPopupClick
'[ 概  要 ]　MoveShape画面のポップアップメニューをクリックした時実行される
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub OnPopupClick()
    Call frmMoveShape.OnPopupClick
End Sub

'*****************************************************************************
'[ 関数名 ]　OnPopupClick2
'[ 概  要 ]　入力補助画面のポップアップメニューをクリックした時実行される
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub OnPopupClick2()
    Call frmEdit.OnPopupClick
End Sub

'*****************************************************************************
'[ 関数名 ]　ConvertStr
'[ 概  要 ]　文字種の変換
'[ 引  数 ]　変換種類
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ConvertStr(ByVal strCommand As String)
On Error GoTo ErrHandle
    Dim objSelection As Range
    Dim objWkRange   As Range
    Dim objCell      As Range
    Dim strText      As String
    Dim strConvText  As String
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    Set objSelection = Selection
    
    '文字の入力されたセルのみ対象にする
    On Error Resume Next
    Set objWkRange = IntersectRange(objSelection, Cells.SpecialCells(xlCellTypeConstants))
    On Error GoTo 0
    If objWkRange Is Nothing Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim i As Long
    Dim j As Long
'    Dim t
'    t = Timer
    
    Call SaveUndoInfo(E_CellValue, objSelection)
    On Error Resume Next
    For Each objCell In objWkRange
        i = i + 1
        strText = objCell
        If strText <> "" Then '結合セルの左上以外を無視するため
            strConvText = StrConvert(strText, strCommand)
            If strConvText <> strText Then
'                objCell = strConvText
                Call SetTextToCell(objCell, strConvText)
                
                'ステータスバーに進捗状況を表示
                If i / objWkRange.Count * 12 <> j Then
                    j = i / objWkRange.Count * 12
                    Application.StatusBar = String(j, "■") & String(12 - j, "□")
                End If
            End If
        End If
    Next
    
'    MsgBox Timer - t
    
    Call SetOnUndo
    Application.StatusBar = False
    Call objSelection.Select
Exit Sub
ErrHandle:
    Application.StatusBar = False
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　SetTextToCell
'[ 概  要 ]　設定するCell,設定する文字列
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SetTextToCell(ByRef objCell As Range, ByVal strConvText As String)
On Error GoTo ErrHandle
    objCell = objCell.PrefixCharacter & strConvText
    If objCell.HasFormula Then
        objCell = "'" & strConvText
    End If
Exit Sub
ErrHandle:
    objCell = "'" & strConvText
End Sub

'*****************************************************************************
'[ 関数名 ]　OpenEdit
'[ 概  要 ]　入力支援エディタを開く
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub OpenEdit()
On Error GoTo ErrHandle
    Static udtEditInfo As TEditInfo
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    'デフォルト値の設定
    If udtEditInfo.Width = 0 Then
        frmEdit.StartUpPosition = 2 '画面の中央
        udtEditInfo.Height = 405
        udtEditInfo.Width = 660
        udtEditInfo.FontSize = 10
        udtEditInfo.WordWarp = False
    Else
        frmEdit.StartUpPosition = 0 '手動
    End If
    
    With frmEdit
        '画面位置の設定
        If .StartUpPosition = 0 Then
            .Top = udtEditInfo.Top
            .Left = udtEditInfo.Left
        End If
        
        'デフォルト値の設定
        .Height = udtEditInfo.Height
        .Width = udtEditInfo.Width
        .SpbSize = udtEditInfo.FontSize
        .Zoomed = udtEditInfo.Zoomed
        .chkWordWrap = udtEditInfo.WordWarp
        
        'フォームを表示
        Call .Show
        
        'フォームのサイズ等を保存
        If ActiveCell.HasFormula = True Then
            udtEditInfo.Height = WorksheetFunction.Max(.Height, udtEditInfo.Height)
        Else
            udtEditInfo.Height = .Height
        End If
        udtEditInfo.Top = .Top
        udtEditInfo.Left = .Left
        udtEditInfo.Width = .Width
        udtEditInfo.FontSize = .SpbSize
        udtEditInfo.Zoomed = .Zoomed
        udtEditInfo.WordWarp = .chkWordWrap
        
        Call Unload(frmEdit)
    End With
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 選択されたセル上の図形を選択する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SelectShapes()
On Error GoTo ErrHandle
    If ActiveWorkbook.DisplayDrawingObjects = xlHide Then
        Exit Sub
    End If
    
    Call ShowSelectionPane
    
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    If ActiveSheet.Shapes.Count = 0 Then
        Exit Sub
    End If
    
    ReDim lngArray(1 To ActiveSheet.Shapes.Count)
    Dim i As Long
    Dim j As Long
    For i = 1 To ActiveSheet.Shapes.Count
        'コメントの図形は対象外とする
        If ActiveSheet.Shapes(i).Type <> msoComment Then
            If IsInclude(Selection, ActiveSheet.Shapes(i)) Then
                j = j + 1
                lngArray(j) = i
            End If
        End If
    Next
    
    '対象の図形がない時は、図形選択モードにする
    If j = 0 Then
        On Error Resume Next
        Call CommandBars.ExecuteMso("ObjectsSelect")
        Exit Sub
    End If
    
    ReDim Preserve lngArray(1 To j)
    Call ActiveSheet.Shapes.Range(lngArray).Select
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 図形がRangeエリアに含まれるかどうか判定
'[引数] Rangeエリア、判定する図形
'[戻値] なし
'*****************************************************************************
Private Function IsInclude(ByRef objRange As Range, ByRef objShape As Shape) As Boolean
    On Error Resume Next
    With objShape
        IsInclude = (MinusRange(Range(.TopLeftCell, .BottomRightCell), objRange) Is Nothing)
    End With
End Function

'*****************************************************************************
'[ 関数名 ]　PressBackSpace
'[ 概  要 ]　バックスペースキーを押したときの動作を変更する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub PressBackSpace()
On Error GoTo ErrHandle
    Call Application.OnKey("{BS}")
    If CheckSelection() = E_Range Then
        Call SendKeys("{F2}")
    Else
        Call SendKeys("{BS}")
    End If
    Call Application.OnTime(Now(), "SetBackSpace")
ErrHandle:
End Sub

'*****************************************************************************
'[ 関数名 ]　SetBackSpace
'[ 概  要 ]　バックスペースキーを押したときの動作を変更する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SetBackSpace()
On Error GoTo ErrHandle
    Call Application.OnKey("{BS}", "PressBackSpace")
ErrHandle:
End Sub

'*****************************************************************************
'[ 関数名 ]　SetOption
'[ 概  要 ]　オプションの設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SetOption()
    'フォームを表示
    Call frmOption.Show
    
    'ショートカットキーの設定
    Call SetKeys
End Sub

'*****************************************************************************
'[ 関数名 ]　SetKeys
'[ 概  要 ]　ショートカットキーの設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetKeys()
On Error GoTo ErrHandle
    Dim strOption As String
    Dim blnKeys(1 To 4) As Boolean
    
    blnKeys(1) = GetSetting(REGKEY, "KEY", "OpenEdit", True)
    blnKeys(2) = GetSetting(REGKEY, "KEY", "CopyText", True)
    blnKeys(3) = GetSetting(REGKEY, "KEY", "PasteText", True)
    blnKeys(4) = GetSetting(REGKEY, "KEY", "BackSpace", True)
    
    If blnKeys(1) = True Then
        Call Application.OnKey("+{F2}", "OpenEdit")
    Else
        Call Application.OnKey("+{F2}")
    End If
    
    If blnKeys(2) = True Then
        Call Application.OnKey("+^{c}", "CopyText")
    Else
        Call Application.OnKey("+^{c}")
    End If
    
    If blnKeys(3) = True Then
        Call Application.OnKey("+^{v}", "PasteText")
    Else
        Call Application.OnKey("+^{v}")
    End If
    
    If blnKeys(4) = True Then
        Call Application.OnKey("{BS}", "PressBackSpace")
    Else
        Call Application.OnKey("{BS}")
    End If

    Call Application.OnKey("+ ", "SelectRow")
    Call Application.OnKey("^ ", "SelectCol")
    Call Application.OnKey("^6", "ToggleHideShapes")
ErrHandle:
End Sub

'*****************************************************************************
'[概要] Shift+{SPACE}で行全体を選択する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SelectRow()
    Dim objSelection As Range
    Dim objExceptRange As Range
    If CheckSelection = E_Range Then
        Set objSelection = Selection.EntireRow
        If objSelection.Areas.Count > 50 Then
            '高速化のため
            Call objSelection.Select
        Else
            '縦方向に結合がはみ出ているセルを除く
            Set objExceptRange = ArrangeRange(MinusRange(ArrangeRange(objSelection), objSelection))
            Call MinusRange(objSelection, objExceptRange).Select
        End If
    End If
End Sub

'*****************************************************************************
'[概要] Ctrl+{SPACE}で列全体を選択する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SelectCol()
    Dim objSelection As Range
    Dim objExceptRange As Range
    If CheckSelection = E_Range Then
        Set objSelection = Selection.EntireColumn
        '横方向に結合がはみ出ているセルを除く
        Set objExceptRange = ArrangeRange(MinusRange(ArrangeRange(objSelection), objSelection))
        Call MinusRange(objSelection, objExceptRange).Select
    End If
End Sub

'*****************************************************************************
'[概要] 見たまま貼付け
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub PasteAppearance()
    Dim CopyMode  As XlCutCopyMode
    CopyMode = Application.CutCopyMode
    
    '選択されているオブジェクトを判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    If CopyMode = xlCopy Or CopyMode = xlCut Then
    Else
        Call MsgBox("セルを [コピー] または [切取り] してから実行して下さい", vbExclamation)
        Exit Sub
    End If
    
    'IMEをオフにする
    Call SetIMEOff

On Error GoTo ErrHandle
    Dim objFromRange   As Range
    Dim objToRange As Range
    Dim lngCutCopyMode As Long
    
    'コピー元のRangeを設定する
    Set objFromRange = GetCopyRange()
    If objFromRange Is Nothing Then
        Call MsgBox("セルをコピーしてから実行して下さい", vbExclamation)
        Exit Sub
    End If
    
    'コピー先のRangeを設定する
    Set objToRange = Selection(1)
    
    If CopyMode = xlCut Then
        If Not CheckSameSheet(objFromRange.Worksheet, objToRange.Worksheet) Then
            Call MsgBox("切り取り時は同じシートでなければなりません", vbExclamation)
            Exit Sub
        End If
    End If
    
    'EXCEL2013以降で起動直後にMoveCellを実行するとボタンが固まる謎の現象を回避するためにSetPixelInfoを呼ぶ
    Call SetPixelInfo
    Call ShowCopyCellForm(objFromRange, objToRange, CopyMode = xlCut)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 結合セルを含む領域を移動する
'[引数] objFromRange:移動(コピー元)の領域
'       objToRange:選択中の領域
'       blnMove:True=移動する
'[戻値] なし
'*****************************************************************************
Private Sub ShowCopyCellForm(ByRef objFromRange As Range, ByRef objToRange As Range, ByVal blnMove As Boolean)
    Dim blnCopyObjectsWithCells  As Boolean
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells
On Error GoTo ErrHandle
    'フォームを表示
    With frmCopyCell
        Call .Initialize(objFromRange, objToRange, blnMove)
        Call .Show
    End With
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
Exit Sub
ErrHandle:
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
    If blnFormLoad = True Then
        Call Unload(frmCopyCell)
    End If
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 数式をプレーンテキストでコピーします
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub CopyFormula()
On Error GoTo ErrHandle
    Dim strText      As String
    Dim objSelection As Range
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = Selection
    
    '選択領域が複数の時
    If objSelection.Areas.Count > 1 Then
        Call MsgBox("このコマンドは複数の選択範囲に対して実行できません。", vbExclamation)
        Exit Sub
    End If
    
    'すべての行の選択時
    Dim lngLast As Long
    If objSelection.Rows.Count = Rows.Count Then
        '使用されている最後の行
        lngLast = Cells.SpecialCells(xlCellTypeLastCell).Row
    Else
        lngLast = objSelection.Rows.Count
    End If

    '行の数だけループ
    Dim i As Long
    strText = GetRowFormula(objSelection.Rows(1))
    For i = 2 To lngLast
        strText = strText & vbLf & GetRowFormula(objSelection.Rows(i))
    Next
    
    If strText <> "" Then
        Call SetClipbordText(Replace$(strText, vbLf, vbCrLf))
    End If
    Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 対象行の数式をTabで連結して取得
'[引数] 対象行
'[戻値] Tabで連結した数式
'*****************************************************************************
Private Function GetRowFormula(ByRef objRange As Range) As String
    Dim i       As Long
    Dim strText As String
    
    '列の数だけループ
    GetRowFormula = objRange.Columns(1).Formula
    For i = 2 To objRange.Columns.Count
        GetRowFormula = GetRowFormula & vbTab & objRange.Columns(i).Formula
    Next i
End Function

'*****************************************************************************
'[概要] 数式や値が文字列になった時、書式を反映します
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ApplyFormat()
On Error GoTo ErrHandle
    Dim objSelection  As Range
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If

    Set objSelection = Selection
    Dim objUsedRange  As Range
    Set objUsedRange = IntersectRange(objSelection, objSelection.Worksheet.UsedRange)
    
    Application.ScreenUpdating = False
    
    'アンドゥ用に元の状態を保存する
    Call SaveUndoInfo(E_ApplyFormat, objSelection)
    Dim objArea As Range
    For Each objArea In objUsedRange.Areas
        If FPressKey = E_Shift Then
            objArea.Value = objArea.Value
        Else
            objArea.Formula = objArea.Formula
        End If
    Next
    Call objSelection.Select
    Call SetOnUndo
ErrHandle:
End Sub


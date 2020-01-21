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

Private objRepeatCmd As CommandBarControl

'*****************************************************************************
'[ 関数名 ]　ColRowChange
'[ 概  要 ]　行列切替え
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ColRowChange()
    Dim i As Long
    Dim strClipText As String
    
    Call SetKeys
    
    'Ctlr + Shift が押下されていれば
    If GetKeyState(vbKeyControl) < 0 Or GetKeyState(vbKeyShift) < 0 Then
        'オプション設定画面を表示
        Call SetOption
        Exit Sub
    End If
    
    strClipText = GetClipbordText()
    
    With CommandBars("かんたんレイアウト").Controls(1)
        If .Caption = "行" Then
            .Caption = "列"
        Else
            .Caption = "行"
        End If
    End With
    
    Application.ScreenUpdating = False
    On Error Resume Next
    With CommandBars("かんたんレイアウト").Controls
        For i = 2 To .Count
            Call SetCommand(.Item(1).Caption, .Item(i))
        Next i
    End With
    
    'クリップボードの復元
    If strClipText = "" Then
        Call ClearClipbord
    Else
        Call SetClipbordText(strClipText)
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　SetCommand
'[ 概  要 ]　コマンドバーボタンにアイコン･コマンド･TooltipTextを設定
'[ 引  数 ]　strGroup : "行" or "列" or "その他"
'            objCmdBarBtn : コマンドを切替えるボタン
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetCommand(ByVal strGroup As String, ByRef objCmdBarBtn As CommandBarButton)
    Dim i As Long
    Dim strCommand As String
    Dim objRange   As Range
'    Dim objMask    As IPictureDisp
    
    strCommand = objCmdBarBtn.Caption
    
    'ワークシ－トのコマンドを設定する
    Set objRange = ThisWorkbook.Worksheets("Commands").Cells(1, "A").CurrentRegion

    For i = 2 To objRange.Rows.Count
        If objRange(i, "A") = strGroup And objRange(i, "B") = strCommand Then
            With objCmdBarBtn
                'アイコンの設定
                If objRange(i, "C") <> "" Then
                    .FaceId = objRange(i, "C")
                Else
                    If CopyIconFromHidden(strGroup & "_" & strCommand) = True Then
                        Call .PasteFace
                    End If
'                    Set objMask = Nothing
'                    If CopyIconFromCell(objRange.Cells(i, "E")) = True Then
'                        Call .PasteFace
'                        Set objMask = .Picture
'                    End If
'                    If CopyIconFromCell(objRange.Cells(i, "D")) = True Then
'                        Call .PasteFace
'                        If Not (objMask Is Nothing) Then
'                            .Mask = objMask
'                        End If
'                    End If
'                    .Picture = objRange.Parent.OLEObjects(strPrefix & .Caption).Object.Picture
                End If
                
                'コマンドの設定
                If objRange(i, "F") = .ID Then
                    .OnAction = ""
                Else
                    .OnAction = objRange(i, "G")
                    .Parameter = objRange(i, "H")
                End If
                
                'ヘルプの設定
                If Val(Application.Version) >= 12 Then
                    .TooltipText = Replace$(objRange(i, "I"), vbLf, "  ")
                Else
                    .TooltipText = objRange(i, "I")
                End If

'                .Tag = .Caption 'サイズ一覧の情報の保存にTagを使用するので注意すること
            End With
                    
            Exit Sub
        End If
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　CopyIconFromCell
'[ 概  要 ]　引数のセルに含まれるアイコンをクリップボードにコピーする
'[ 引  数 ]　アイコンを含むセル
'[ 戻り値 ]　True:成功、False:失敗
'*****************************************************************************
'Private Function CopyIconFromCell(ByRef objCell As Range) As Boolean
'    Dim objShape As Shape
'    For Each objShape In objCell.Worksheet.Shapes
'        If objCell.Top = objShape.Top And _
'           objCell.Left = objShape.Left Then
'            Call objShape.CopyPicture(xlScreen, xlBitmap)
'            CopyIconFromCell = True
'            Exit Function
'        End If
'    Next objShape
'End Function

'*****************************************************************************
'[ 関数名 ]　CopyIconFromHidden
'[ 概  要 ]　かんたんレイアウトアイコンのアイコンをクリップボードにコピーする
'[ 引  数 ]　コマンドの名前　例：列_縮小
'[ 戻り値 ]　True:成功、False:失敗
'*****************************************************************************
Private Function CopyIconFromHidden(ByVal strCommand As String) As Boolean
On Error GoTo ErrHandle
    Dim objBtn As CommandBarButton

    For Each objBtn In CommandBars("かんたんレイアウトアイコン").Controls
        If objBtn.Caption = strCommand Then
            Call objBtn.CopyFace
            CopyIconFromHidden = True
            Exit Function
        End If
    Next objBtn
ErrHandle:
End Function

'*****************************************************************************
'[ 関数名 ]　OpenHelp
'[ 概  要 ]　ヘルプファイルを開く
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub OpenHelp()
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
    
    strSelection = Selection.Address
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
    Call SaveUndoInfo(E_MergeCell, Range(strSelection))
    
    'エリアの数だけループ
    For Each objRange In Range(strSelection).Areas
        Call MergeArea(objRange)
    Next objRange
    
    Call SetOnUndo
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
    Call SetOnRepeat
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
    Call SetOnRepeat
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
'[ 関数名 ]　PasteValue
'[ 概  要 ]　値を貼り付ける
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub PasteValue()
    Call SetKeys
    
    If GetKeyState(vbKeyControl) < 0 Then
        Call CopyText  'セルの値をクリップボードにコピーする
    Else
        Call PasteText '値をセルに貼り付ける
    End If
End Sub
    
'*****************************************************************************
'[ 関数名 ]　PasteText
'[ 概  要 ]　値をセルに貼り付ける
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub PasteText()
On Error GoTo ErrHandle
    Dim strCopyRange  As String
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
        strCopyRange = GetCopyRangeAddress()
        On Error GoTo 0
        If strCopyRange <> "" Then
            blnOnlyCell = IsOnlyCell(Range(strCopyRange))
            If blnOnlyCell Then
                strCopyText = GetCellText(Range(strCopyRange).Cells(1, 1))
            Else
                strCopyText = MakeCopyText(strCopyRange)
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
    Call objSelection.Parent.Activate
    Call objSelection.Select
    
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
    Call SetOnRepeat
    
    '貼り付けた文字列をクリップボードにコピー
    Call SetClipbordText(Replace$(strCopyText, vbLf, vbCrLf))
ErrHandle:
End Sub
   
'*****************************************************************************
'[ 関数名 ]　MakeCopyText
'[ 概  要 ]　コピー対象の文字列を作成する
'[ 引  数 ]　Copy中の領域
'[ 戻り値 ]　なし
'*****************************************************************************
Private Function MakeCopyText(Optional ByVal strCopyRange As String = "") As String
On Error GoTo ErrHandle
    If strCopyRange <> "" Then
        MakeCopyText = GetRangeText(Range(strCopyRange))
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
'[ 関数名 ]　MoveObject
'[ 概  要 ]　Range選択時：結合セルを含む領域を移動する
'　　　　　　Shape選択時：図形を移動またはサイズ変更する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveObject()
    If GetKeyState(vbKeyControl) < 0 Then
        Call UnSelect
        Exit Sub
    End If

    'IMEをオフにする
    Call SetIMEOff

    '選択されているオブジェクトを判定
    Select Case CheckSelection()
    Case E_Range
        Call MoveCell
    Case E_Shape
        Call MoveShape
    End Select
    
    Call Application.OnRepeat("", "")
End Sub
    
'*****************************************************************************
'[ 関数名 ]　MoveCell
'[ 概  要 ]　結合セルを含む領域を移動する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveCell()
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
        Dim strFromRange As String
        strFromRange = GetCopyRangeAddress()
        If strFromRange <> "" Then
            Set objFromRange = Range(strFromRange)
        Else
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
    Dim objSelection As Range
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
            If objSelection.Areas.Count > 1 And strLastAddress = "" Then
                strLastSheet = ActiveSheet.Name
                strLastAddress = Selection.Address(False, False)
            End If
            Exit Sub
        End If
        
        enmUnselectMode = .Mode
        Set objSelection = Selection
        Select Case (enmUnselectMode)
        Case E_Unselect, E_Reverse, E_Intersect, E_Union
            Set objUnSelect = .SelectRange
        End Select
        Call Unload(frmUnSelect)
    End With

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
    strLastAddress = Selection.Address(False, False)
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
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub FitShapes()
On Error GoTo ErrHandle
    Dim blnOK      As Boolean
    Dim blnSizeChg As Boolean 'サイズを変更するかどうか
    Dim i          As Long
    
    '図形が選択されているか判定
    Select Case (CheckSelection())
    Case E_Range
        Call MsgBox("図形が選択されていません", vbExclamation)
        Exit Sub
    Case E_Other
        Exit Sub
    End Select
    
    '****************************************
    'タイプを選択させる
    '****************************************
    Dim enmSizeType As EFitType  '選択されたタイプ
    With frmFitShapes
        'フォームを表示
        Call .Show

        'キャンセル時
        If blnFormLoad = False Then
            Exit Sub
        End If

        enmSizeType = .SelectType
        Call Unload(frmFitShapes)
    End With
    
    Select Case enmSizeType
    Case E_Default
        blnSizeChg = True
    Case E_TopLeft
        blnSizeChg = False
    Case E_Another
        Call MoveShape
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
    Call SetOnRepeat
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
    Call SetOnRepeat
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
    
    'フォントと文字列の設定
    objRange(1, 1).Value = GetCharactersText(objTextbox.TextFrame)
    On Error Resume Next
    With objTextbox.TextFrame.Font
        objRange.Font.Name = .Name
        objRange.Font.Size = .Size
        objRange.Font.FontStyle = .FontStyle
    End With
    On Error GoTo 0
    
    '配置の設定
    With objTextbox.TextFrame
        On Error Resume Next
        If .Orientation = msoTextOrientationVertical Then
            objRange.Orientation = xlVertical       '縦書き
        End If
        On Error GoTo 0
        
        '横位置設定
        On Error Resume Next
        Select Case .HorizontalAlignment
        Case xlHAlignLeft
            objRange.HorizontalAlignment = xlLeft
        Case xlHAlignCenter
            objRange.HorizontalAlignment = xlCenter
        Case xlHAlignRight
            objRange.HorizontalAlignment = xlRight
        Case xlHAlignDistributed
            objRange.HorizontalAlignment = xlDistributed
        Case xlHAlignJustify
            objRange.HorizontalAlignment = xlJustify
        End Select
        On Error GoTo 0
        
        '縦位置設定
        On Error Resume Next
        Select Case .VerticalAlignment
        Case xlVAlignTop
            objRange.VerticalAlignment = xlTop
        Case xlVAlignCenter
            objRange.VerticalAlignment = xlCenter
        Case xlVAlignBottom
            objRange.VerticalAlignment = xlBottom
        Case xlVAlignDistributed
            objRange.VerticalAlignment = xlDistributed
        Case xlVAlignJustify
            objRange.VerticalAlignment = xlJustify
        End Select
        On Error GoTo 0
    End With

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
    Call SetOnRepeat
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　ChangeCellToTextbox
'[ 概  要 ]　セルをテキストボックスに変換する
'[ 引  数 ]　セル
'[ 戻り値 ]　テキストボックス
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
    objTextbox.DrawingObject.Formula = objCell.Address
    objTextbox.DrawingObject.Formula = ""
    If Val(Application.Version) >= 12 Then
        With objTextbox.TextFrame.Characters
            If .Text <> "" Then
                .Font.Name = objCell.Font.Name
                .Font.Size = objCell.Font.Size
                .Font.FontStyle = objCell.Font.FontStyle
            End If
        End With
    End If
    
    '配置の設定
    With objTextbox.TextFrame
        If objCell.Orientation = xlVertical Then
            .Orientation = msoTextOrientationVertical     '縦書き
        End If
        
        '横位置設定
        Select Case objCell.HorizontalAlignment
        Case xlLeft
            .HorizontalAlignment = xlHAlignLeft
        Case xlCenter
            .HorizontalAlignment = xlHAlignCenter
        Case xlRight
            .HorizontalAlignment = xlHAlignRight
        Case xlDistributed
            .HorizontalAlignment = xlHAlignDistributed
        Case xlJustify
            .HorizontalAlignment = xlHAlignJustify
        End Select
        
        '縦位置設定
        Select Case objCell.VerticalAlignment
        Case xlTop
            .VerticalAlignment = xlVAlignTop
        Case xlCenter
            .VerticalAlignment = xlVAlignCenter
        Case xlBottom
            .VerticalAlignment = xlVAlignBottom
        Case xlDistributed
            .VerticalAlignment = xlVAlignDistributed
        Case xlJustify
            .VerticalAlignment = xlVAlignJustify
        End Select
    End With
    
    '線の設定
    With objTextbox.Line
        .Weight = DPIRatio
        .DashStyle = msoLineSolid
        .Style = msoLineSingle
        .ForeColor.RGB = 0
        .BackColor.RGB = RGB(255, 255, 255)
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
'[ 関数名 ]　HideShapes
'[ 概  要 ]  ブック内のすべての図形を非表示にする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub HideShapes()
    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
    With ActiveWorkbook
        If .DisplayDrawingObjects = xlDisplayShapes Then
            If MsgBox("ワークブック内のすべての図形を非表示にします" & vbLf & _
                      "実行すれば、図形の編集が出来なくなります" & vbLf & _
                      "よろしいですか？" & vbLf & _
                      "※再度表示するには、もう一度クリックして下さい" & vbLf & _
                      "※「Ctrl+6」を押下してもExcelの標準機能で同様の操作が行えます" _
                      , vbOKCancel + vbQuestion) = vbOK Then
                .DisplayDrawingObjects = xlHide
            End If
        Else
            .DisplayDrawingObjects = xlDisplayShapes
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
Public Sub SaveUndoInfo(ByVal enmType As EUndoType, ByRef objObject As Object, Optional ByVal varInfo As Variant = Nothing)
    'すでにインスタンスが存在する時は、Rollbackに対応するためNewしない
    If clsUndoObject Is Nothing Then
        Set clsUndoObject = New CUndoObject
    End If
    
    'オートフィルタが設定されている時は、Undo不可にする
    If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
        Call clsUndoObject.SaveUndoInfo(enmType, objObject, varInfo)
    Else
        Call clsUndoObject.SaveUndoInfo(E_FilterERR, objObject, varInfo)
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　SetOnUndo
'[ 概  要 ]　ApplicationオブジェクトのOnUndoイベントを設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetOnUndo()
    'キーの再定義 Excelのバグ?でキーが無効になることがあるため
    Call SetKeys

    Call clsUndoObject.SetOnUndo
    Call Application.OnRepeat("", "")
End Sub

'*****************************************************************************
'[ 関数名 ]　SetOnRepeat
'[ 概  要 ]　ApplicationオブジェクトのOnRepeatイベントを設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetOnRepeat()
    Set objRepeatCmd = CommandBars.ActionControl
    If Not (objRepeatCmd Is Nothing) Then
        Call Application.OnRepeat("繰り返し　" & CommandBars.ActionControl.Caption, "OnRepeat")
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　OnRepeat
'[ 概  要 ]　繰り返しクリック時
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub OnRepeat()
    Call objRepeatCmd.Execute
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
Public Sub OnPopupClick()
    Call frmMoveShape.OnPopupClick
End Sub

'*****************************************************************************
'[ 関数名 ]　OnPopupClick2
'[ 概  要 ]　入力補助画面のポップアップメニューをクリックした時実行される
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub OnPopupClick2()
    Call frmEdit.OnPopupClick
End Sub

'*****************************************************************************
'[ 関数名 ]　ConvertStr
'[ 概  要 ]　文字種の変換
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ConvertStr()
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
            strConvText = StrConvert(strText, CommandBars.ActionControl.Parameter)
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
    
    Set objRepeatCmd = CommandBars.ActionControl
    Call Application.OnRepeat("繰り返し " & objRepeatCmd.Caption, "OnRepeat")
    Application.StatusBar = False
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
'[ 関数名 ]　OnElseMenuClick
'[ 概  要 ]　｢その他｣メニューを開く時に、各コマンドのEnabledの初期設定を行う
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub OnElseMenuClick()
    Dim objMenu          As CommandBarPopup
    Dim enmSelectionType As ESelectionType
    Dim objCommand       As Object
        
    Call SetKeys
    
    enmSelectionType = CheckSelection()
    Set objMenu = CommandBars.ActionControl
    
    'Menuがカスタマイズされていた場合の考慮
    For Each objCommand In objMenu.Controls
        objCommand.Enabled = True
    Next
    
    On Error Resume Next
    objMenu.Controls("テキストボックスをセルに変換").Enabled = (enmSelectionType = E_Shape)
    objMenu.Controls("セルをテキストボックスに変換").Enabled = (enmSelectionType = E_Range)
    objMenu.Controls("図形をグリッドに合せる").Enabled = (enmSelectionType = E_Shape)
    objMenu.Controls("文字種の変換").Enabled = (enmSelectionType = E_Range)
    objMenu.Controls("選択領域の一部取消し").Enabled = (enmSelectionType = E_Range)
    objMenu.Controls("入力補助画面").Enabled = (enmSelectionType = E_Range)

    '図形非表示が図形再表示になっていたら元に戻す
    objMenu.Controls("図形再表示").Caption = "図形非表示"
    If ActiveWorkbook Is Nothing Then
        objMenu.Controls("図形非表示").Enabled = False
    Else
        If ActiveWorkbook.DisplayDrawingObjects <> xlDisplayShapes Then
            objMenu.Controls("図形非表示").Caption = "図形再表示"
        End If
    End If
    On Error GoTo 0
End Sub

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
Public Sub SetBackSpace()
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
    
    strOption = CommandBars("かんたんレイアウト").Controls(1).Tag
    blnKeys(1) = Not (InStr(1, strOption, "{S+F2}") = 0)
    blnKeys(2) = Not (InStr(1, strOption, "{C+S+C}") = 0)
    blnKeys(3) = Not (InStr(1, strOption, "{C+S+V}") = 0)
    blnKeys(4) = Not (InStr(1, strOption, "{BS}") = 0)
    
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
ErrHandle:
End Sub

''*****************************************************************************
''[ 関数名 ]　SubClassProc
''[ 概  要 ]　入力補助画面をマウスホイールでスクロールさせる
''[ 引  数 ]　CallBack関数のそれ
''[ 戻り値 ]　CallBack関数のそれ
''*****************************************************************************
'Public Function SubClassProc(ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'On Error Resume Next
'    If MSG = WM_MOUSEWHEEL Then
'        If 0 < wParam Then
'            Call SendKeys("{UP}")
'            Call SendKeys("{UP}")
''            frmEdit.txtEdit.CurLine = frmEdit.txtEdit.CurLine - 2
'        Else
'            Call SendKeys("{DOWN}")
'            Call SendKeys("{DOWN}")
''            frmEdit.txtEdit.CurLine = frmEdit.txtEdit.CurLine + 2
'        End If
'    End If
'
'    'デフォルトウィンドウプロシージャを呼び出す
'    SubClassProc = CallWindowProc(frmEdit.WndProc, hWnd, MSG, wParam, lParam)
'End Function

''*****************************************************************************
''[ 関数名 ]　LoadThisWorkbook
''[ 概  要 ]　Open時に高速化のため、ThisWorkbookをロードさせる
''[ 引  数 ]　なし
''[ 戻り値 ]　なし
''*****************************************************************************
'Public Sub LoadThisWorkbook()
'On Error Resume Next
'    Application.StatusBar = "開いています  かんたんレイアウト"
'    With ThisWorkbook.Styles("Normal").Font
''        .Size = 11
'    End With
'    Application.StatusBar = False
'End Sub

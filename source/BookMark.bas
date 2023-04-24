Attribute VB_Name = "BookMark"
Option Explicit
Option Private Module

Public Const C_PatternColor = &H800080  '何でも良い(決めの問題だけ)
Public FBMarkColor As Long
Public FFillColor As Long '塗りつぶし色
    
'*****************************************************************************
'[概要] 選択セル(または図形)の塗りつぶし
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub FillColor()
On Error GoTo ErrHandle
    Select Case CheckSelection()
    Case E_Range
        Dim objSelection As Range
        Set objSelection = Selection
        'アンドゥ用に元の情報を保存する
        Call SaveUndoInfo(E_FillRange, GetAddress(Selection))
        Call objSelection.Select '選択範囲を戻す(なぜか全選択される事象が起きることがある)
    Case E_Shape
        Dim objShapeRange As ShapeRange
        Set objShapeRange = HasInteriorShapes(Selection.ShapeRange)
        If objShapeRange Is Nothing Then
            Exit Sub
        End If
        'アンドゥ用に元の情報を保存する
        Call SaveUndoInfo(E_FillShape, objShapeRange)
        Call objShapeRange.Select
        '塗りつぶしありとなしが混在する時は、そのままでは塗りつぶしなしの図形に色がつかないので
        'いったんすべての図形をクリアする
        Selection.Interior.ColorIndex = xlNone
    End Select
    
    Selection.Interior.Color = FFillColor
    Call SetOnUndo
ErrHandle:
    Call GetRibbonUI.InvalidateControl("B631")
End Sub

'*****************************************************************************
'[概要] ShapeRangeのうちInteriorが有効なShapeRangeのみ返す
'[引数] ShapeRange
'[戻値] Interiorを持つShapeRange
'*****************************************************************************
Private Function HasInteriorShapes(ByRef objShapeRange As ShapeRange) As ShapeRange
    Dim i As Long, j As Long
    Dim Dummy
    ReDim lngIDArray(1 To objShapeRange.Count) As Variant
    
    '図形の数だけループ
    For j = 1 To objShapeRange.Count 'For each構文だとExcel2007で型違いとなる(たぶんバグ)
        On Error Resume Next
        Dummy = objShapeRange(j).DrawingObject.Interior.Color
        Dummy = objShapeRange(j).ID
        If Err.Number = 0 Then
            i = i + 1
            lngIDArray(i) = objShapeRange(j).ID
        End If
        On Error GoTo 0
    Next
    If i > 0 Then
        ReDim Preserve lngIDArray(1 To i)
        Set HasInteriorShapes = GetShapeRangeFromID(lngIDArray)
    End If
End Function

'*****************************************************************************
'[概要] 選択されセルにBookmarkを設定/解除する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SetBookmark()
On Error GoTo ErrHandle
    Dim objRange As Range
    
    'Rangeオブジェクトが選択されているか判定
    If TypeOf Selection Is Range Then
        Set objRange = Selection
    Else
        Exit Sub
    End If
        
    With objRange.Cells(1).Interior
        If .Pattern = xlSolid And _
           .PatternColor = C_PatternColor Then
            '書式のクリア
            objRange.Interior.ColorIndex = xlNone
            Exit Sub
        End If
    End With
    
    With objRange.Interior
        .Color = FBMarkColor
        .Pattern = xlSolid
        .PatternColor = C_PatternColor
    End With
ErrHandle:
    Call GetRibbonUI.InvalidateControl("B621")
    Call GetRibbonUI.InvalidateControl("C2")
End Sub

'*****************************************************************************
'[概要] 次のBookmarkに移動
'[引数] 検索方向
'[戻値] なし
'*****************************************************************************
Private Sub NextOrPrevBookmark(ByVal xlDirection As XlSearchDirection)
    '[Shift]または[Ctrl]Keyが押下されていれば、逆方向に検索
    If xlDirection = xlNext Then
        If FPressKey <> 0 Then
            Call JumpBookmark(xlPrevious)
        Else
            Call JumpBookmark(xlNext)
        End If
    Else
        If FPressKey <> 0 Then
            Call JumpBookmark(xlNext)
        Else
            Call JumpBookmark(xlPrevious)
        End If
    End If
'    Call GetRibbonUI.InvalidateControl("C2")
End Sub

'*****************************************************************************
'[概要] 次のBookmarkに移動
'[引数] 検索方向
'[戻値] なし
'*****************************************************************************
Private Sub JumpBookmark(ByVal xlDirection As XlSearchDirection)
On Error GoTo ErrHandle
    Dim objCell      As Range
    Dim objNextCell  As Range
    Dim objSheetCell As Range
    
    Call SetFindFormat
    
    '****************************************
    'アクティブシート内の検索
    '****************************************
    Dim blnFind  As Boolean
    Set objCell = ActiveCell
    Set objNextCell = FindNextFormat(objCell, xlDirection)
    If (objNextCell Is Nothing) Then
        '他のシートを対象とするかどうか
        If GetTmpControl("C2").State = False Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    Else
        Set objSheetCell = objNextCell
        If TypeOf Selection Is Range Then
            '他のシートを対象とするかどうか
            If GetTmpControl("C2").State = False Then
                blnFind = True
            Else
                If xlDirection = xlNext Then
                    If objNextCell.Row > objCell.Row Or _
                      (objNextCell.Row = objCell.Row And objNextCell.Column > objCell.Column) Then
                        blnFind = True
                    End If
                Else
                    If objNextCell.Row < objCell.Row Or _
                      (objNextCell.Row = objCell.Row And objNextCell.Column < objCell.Column) Then
                        blnFind = True
                    End If
                End If
            End If
        Else
            blnFind = True
        End If
    End If
    
    If blnFind = True Then
        Call objNextCell.Select
        Application.FindFormat.Clear
        Exit Sub
    End If
    
    '****************************************
    '隣のシートの検索
    '****************************************
    Dim i As Long
    Dim j As Long
    Dim lngSheetCnt As Long
    Dim lngStartIdx As Long
    
    lngSheetCnt = ActiveWorkbook.Worksheets.Count
    j = ActiveWorkbook.ActiveSheet.Index
    
    For i = 2 To lngSheetCnt
        If xlDirection = xlNext Then
            j = j + 1
            If j > lngSheetCnt Then
                j = 1
            End If
            Set objCell = ActiveWorkbook.Worksheets(j).Cells(Rows.Count, Columns.Count)
        Else
            j = j - 1
            If j < 1 Then
                j = lngSheetCnt
            End If
            Set objCell = ActiveWorkbook.Worksheets(j).Cells(1, 1)
        End If
        
        Set objNextCell = FindNextFormat(objCell, xlDirection)
        If Not (objNextCell Is Nothing) Then
            Call objNextCell.Worksheet.Select
            Call objNextCell.Select
            Application.FindFormat.Clear
            Exit Sub
        End If
    Next

    If Not (objSheetCell Is Nothing) Then
        Call objSheetCell.Select
    End If
ErrHandle:
    Application.FindFormat.Clear
End Sub

'*****************************************************************************
'[概要] Bookmark検索用のセル書式を設定する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SetFindFormat()
    Application.FindFormat.Clear
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColor = C_PatternColor
    End With
    
    If TypeOf Selection Is Range Then
        '選択されているセルが1つだけか判定
        If Not IsOnlyCell(Selection) Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    With ActiveCell.Interior
        If .Pattern = xlSolid And _
           .PatternColor = C_PatternColor Then
            Application.FindFormat.Interior.Color = .Color
        End If
    End With
End Sub

'*****************************************************************************
'[概要] すべてのBookmarkを選択
'[引数] なし
'[戻値] なし
'*****************************************************************************
'Private Sub SelectAllBookmarks()
'On Error GoTo ErrHandle
'    Dim objRange As Range
'
'    Call SetFindFormat
'    Set objRange = GetBookmarks(ActiveWorkbook.ActiveSheet)
'    If Not (objRange Is Nothing) Then
'        Call objRange.Select
'    End If
'ErrHandle:
'    Application.FindFormat.Clear
'End Sub

'*****************************************************************************
'[概要] すべてのBookmarkをクリア
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ClearBookmarks()
On Error GoTo ErrHandle
    Dim i As Long
    Dim j1 As Long
    Dim j2 As Long
    Dim objRange As Range
    Dim objActiveSheetRange As Range
    Dim colRange As Collection
    
    Application.FindFormat.Clear
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColor = C_PatternColor
    End With
    
    'アクティブシート内のブックマークを取得
    Set objActiveSheetRange = GetBookmarks(ActiveWorkbook.ActiveSheet)
    If (objActiveSheetRange Is Nothing) Then
        '他のシートを対象としない時
        If GetTmpControl("C2").State = False Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    End If
    
    'アクティブシートのみ対象の時
    If GetTmpControl("C2").State = False Then
        '複数セルを選択している時
        If Not IsOnlyCell(Selection) Then
            Set objRange = IntersectRange(Selection, objActiveSheetRange)
            If Not (objRange Is Nothing) Then
                If MsgBox("選択範囲中の " & objRange.Count & " 個のブックマークを削除します" & vbLf & "よろしいですか？", vbOKCancel + vbQuestion) <> vbCancel Then
                    ArrangeRange(objRange).Interior.ColorIndex = xlNone
                End If
                Application.FindFormat.Clear
                Exit Sub
            End If
        End If
        j1 = objActiveSheetRange.Count
    Else
        Set colRange = New Collection
        'すべてのブックマークの数を計算
        For i = 1 To ActiveWorkbook.Worksheets.Count
            If i = ActiveWorkbook.ActiveSheet.Index Then
                Set objRange = objActiveSheetRange
            Else
                Set objRange = GetBookmarks(ActiveWorkbook.Worksheets(i))
            End If
            If Not (objRange Is Nothing) Then
                j1 = j1 + objRange.Count
                Call colRange.Add(objRange)
            End If
        Next
    End If
    
    If j1 = 0 Then
        Application.FindFormat.Clear
        Exit Sub
    End If
    
    '選択セルが単一のブックマークのセルの時
    If IsOnlyCell(Selection) And _
       Not (IntersectRange(Selection, objActiveSheetRange) Is Nothing) Then
        '選択セルと同色のブックマークの数を計算
        Application.FindFormat.Interior.Color = ActiveCell.Interior.Color
        If GetTmpControl("C2").State = False Then
            '他のシートを対象としない時
            Set objActiveSheetRange = GetBookmarks(ActiveWorkbook.ActiveSheet)
            If Not (objActiveSheetRange Is Nothing) Then
                j2 = objActiveSheetRange.Count
            End If
        Else
            Set colRange = New Collection
            For i = 1 To ActiveWorkbook.Worksheets.Count
                Set objRange = GetBookmarks(ActiveWorkbook.Worksheets(i))
                If Not (objRange Is Nothing) Then
                    j2 = j2 + objRange.Count
                    Call colRange.Add(objRange)
                End If
            Next
        End If
    Else
        j2 = j1
    End If
    
    '****************************************
    '実行確認
    '****************************************
    If j1 = j2 Then
        If MsgBox(j1 & " 個のブックマークを削除します" & vbLf & "よろしいですか？", vbOKCancel + vbQuestion) = vbCancel Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    Else
        If MsgBox(j1 & " 個のブックマークのうち" & vbLf & "選択されたセルと同じ色の " & j2 & " 個のブックマークを削除します" & vbLf & "よろしいですか？", vbOKCancel + vbQuestion) = vbCancel Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    End If
    
    '****************************************
    'すべてのブックマークを削除
    '****************************************
    Application.ScreenUpdating = False
    
    'アクティブシートのみ対象の時
    If GetTmpControl("C2").State = False Then
        ArrangeRange(objActiveSheetRange).Interior.ColorIndex = xlNone
    Else
        For Each objRange In colRange
            ArrangeRange(objRange).Interior.ColorIndex = xlNone
        Next
    End If
ErrHandle:
    Application.FindFormat.Clear
End Sub

'*****************************************************************************
'[概要] 対象シートのすべてのBookmarkを取得
'[引数] 対象シート
'[戻値] Bookmarkが設定されたセルすべて
'*****************************************************************************
Private Function GetBookmarks(ByRef objSheet As Worksheet) As Range
    Dim objCell  As Range
    
    Set objCell = GetLastCell(objSheet)
    Do While (True)
        Set objCell = FindNextFormat(objCell, xlNext)
        If objCell Is Nothing Then
            Exit Function
        End If
        
        If IntersectRange(GetBookmarks, objCell) Is Nothing Then
            Set GetBookmarks = UnionRange(GetBookmarks, objCell)
        Else
            Exit Function
        End If
    Loop
End Function

'*****************************************************************************
'[概要] 次の書式のセルに移動
'[引数] 検索開始セル、検索方向
'[戻値] 次の書式のセル
'*****************************************************************************
Private Function FindNextFormat(ByRef objCell As Range, _
                                ByVal xlDirection As XlSearchDirection) As Range
    Dim objUsedRange As Range
    With objCell.Worksheet
        Set objUsedRange = .Range(.Range("A1"), .Cells.SpecialCells(xlLastCell))
        Set objUsedRange = UnionRange(objUsedRange, objCell)
    End With
    Set FindNextFormat = objUsedRange.Find("", objCell, _
                  xlFormulas, xlPart, xlByRows, xlDirection, False, False, True)
End Function

'*****************************************************************************
'[概要] 次を検索
'[引数] 検索方向
'[戻値] なし
'*****************************************************************************
Private Sub FindNext()
    FPressKey = 0
    Call FindNextOrPrev(xlNext)
End Sub

'*****************************************************************************
'[概要] 次を検索
'[引数] 検索方向
'[戻値] なし
'*****************************************************************************
Private Sub FindPrev()
    FPressKey = 0
    Call FindNextOrPrev(xlPrevious)
End Sub

'*****************************************************************************
'[概要] 次または前を検索
'[引数] 検索方向
'[戻値] なし
'*****************************************************************************
Private Sub FindNextOrPrev(ByVal xlDirection As XlSearchDirection)
On Error GoTo ErrHandle
    Dim objCell As Range
    
    '[Shift]または[Ctrl]Keyが押下されていれば、逆方向に検索
    If xlDirection = xlNext Then
        If FPressKey <> 0 Then
            Set objCell = Cells.FindPrevious(ActiveCell)
        Else
            Set objCell = Cells.FindNext(ActiveCell)
        End If
    Else
        If FPressKey <> 0 Then
            Set objCell = Cells.FindNext(ActiveCell)
        Else
            Set objCell = Cells.FindPrevious(ActiveCell)
        End If
    End If
    
'    Set objCell = FindJump(ActiveCell, xlDirection)
    If Not (objCell Is Nothing) Then
        Call objCell.Select
    End If
ErrHandle:
    Call ActiveCell.Worksheet.Select
End Sub

'*****************************************************************************
'[概要] 次を検索
'[引数] 検索開始セル、検索方向
'[戻値] 次のセル
'*****************************************************************************
'Private Function FindJump(ByRef objNowCell As Range, ByVal xlDirection As XlSearchDirection) As Range
'On Error GoTo ErrHandle
'    Dim objCell As Range
'
'    If xlDirection = xlNext Then
'        Set objCell = Cells.FindNext(objNowCell)
'    Else
'        Set objCell = Cells.FindPrevious(objNowCell)
'    End If
'    If Not (objCell Is Nothing) Then
'        '自分自身のセルを選択する意味不明なバグ対応
'        If objCell.Address = objNowCell.Address Then
'            If xlDirection = xlNext Then
'                Set objCell = Cells.FindNext(objNowCell)
'            Else
'                Set objCell = Cells.FindPrevious(objNowCell)
'            End If
'        End If
'        If objCell.Value <> "" Then
'            Set FindJump = objCell
'        End If
'    End If
'ErrHandle:
'End Function

'*****************************************************************************
'[概要] 使用されている最後のセルを取得する
'[引数] 対象のシート
'[戻値] 最後のセル
'*****************************************************************************
Private Function GetLastCell(ByRef objSheet As Worksheet) As Range
    Set GetLastCell = objSheet.Cells.SpecialCells(xlLastCell)
End Function

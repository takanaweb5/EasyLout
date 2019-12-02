Attribute VB_Name = "RowChangeTools"
Option Explicit

Private Const MaxRowHeight = 409.5  '高さの最大サイズ

Private Sub ReduceRowsHeight()
    Call ChangeRowsHeight(-1)
End Sub
Private Sub ExpandRowsHeight()
    Call ChangeRowsHeight(1)
End Sub
Private Sub MoveRowBorderU()
    Call MoveRowBorder(-1)
End Sub
Private Sub MoveRowBorderD()
    Call MoveRowBorder(1)
End Sub
Private Sub MoveCellBorderU()
    Call MoveCellBorder(-1)
End Sub
Private Sub MoveCellBorderD()
    Call MoveCellBorder(1)
End Sub

'*****************************************************************************
'[ 関数名 ]　ChangeRowsHeight
'[ 概  要 ]　複数行サイズ変更
'[ 引  数 ]　lngSize:変更サイズ(単位：ピクセル)
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ChangeRowsHeight(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objSelection As Range   '選択されたすべての列
    Dim strSelection As String
    
    '[Ctrl]Keyが押下されていれば、移動高さを5倍にする
    If GetKeyState(vbKeyControl) < 0 Then
        lngSize = lngSize * 5
    End If
    
    '選択されているオブジェクトを判定
    Select Case CheckSelection()
    Case E_Other
        Exit Sub
    Case E_Shape
        Call ChangeShapeHeight(lngSize)
        Exit Sub
    End Select
    
    '選択範囲のRowsの和集合を取り重複行を排除する
    strSelection = Selection.Address
    Set objSelection = Union(Selection.EntireRow, Selection.EntireRow)
    
    '***********************************************
    '非表示の行があるかどうかの判定
    '***********************************************
    Dim objVisible     As Range    '可視Range
    '選択範囲の可視部分を取出す
    Set objVisible = GetVisibleCells(objSelection)
    If objVisible Is Nothing Then
        If lngSize < 0 Then
            Call MsgBox("これ以上縮小出来ません", vbExclamation)
            Exit Sub
        End If
    Else
        '非表示の行がある時
        If objSelection.Address <> objVisible.Address Then
            If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
                Select Case MsgBox("非表示の行を対象としますか？", vbYesNoCancel + vbQuestion + vbDefaultButton2)
                Case vbYes
                    If lngSize < 0 Then
                        Call MsgBox("これ以上縮小出来ません", vbExclamation)
                        Exit Sub
                    End If
                Case vbNo
                    '可視セルのみ選択する
                    Call IntersectRange(Selection, objVisible).Select
                    Set objSelection = objVisible
                Case vbCancel
                    Exit Sub
                End Select
            Else
                '可視セルのみ選択する
                Call IntersectRange(Selection, objVisible).Select
                Set objSelection = objVisible
            End If
        End If
    End If
    
    '***********************************************
    '同じ高さの塊ごとにAddressを取得する
    '***********************************************
    Dim colAddress  As New Collection
    If objVisible Is Nothing Then
        Call colAddress.Add(objSelection.Address)
    Else
        Set colAddress = GetSameHeightAddresses(objSelection)
    End If
    
    '***********************************************
    '変更後のサイズのチェック
    '***********************************************
    Dim lngPixel    As Long    '幅(単位:Pixel)
    For i = 1 To colAddress.Count
        lngPixel = HeightToPixel(Range(colAddress(i)).Rows(1).RowHeight) + lngSize
        If lngPixel < 0 Then
            Call MsgBox("これ以上縮小出来ません", vbExclamation)
            Exit Sub
        ElseIf lngPixel > HeightToPixel(MaxRowHeight) Then
            Call MsgBox("これ以上拡大出来ません", vbExclamation)
            Exit Sub
        End If
    Next i
    
    '***********************************************
    'サイズの変更
    '***********************************************
    Dim blnDisplayPageBreaks As Boolean  '改ページ表示
    Application.ScreenUpdating = False
    
    '高速化のため改ページを非表示にする
    If ActiveSheet.DisplayAutomaticPageBreaks = True Then
        blnDisplayPageBreaks = True
        ActiveSheet.DisplayAutomaticPageBreaks = False
    End If
    
    'アンドゥ用に元のサイズを保存する
    Call SaveUndoInfo(E_RowSize, Range(strSelection), colAddress)
    
    '同じ高さの塊ごとに高さを設定する
    For i = 1 To colAddress.Count
        lngPixel = HeightToPixel(Range(colAddress(i)).Rows(1).RowHeight) + lngSize
        Range(colAddress(i)).RowHeight = PixelToHeight(lngPixel)
    Next i
    
    '改ページを元に戻す
    If blnDisplayPageBreaks = True Then
        ActiveSheet.DisplayAutomaticPageBreaks = True
    End If
    Call SetOnUndo
Exit Sub
ErrHandle:
    If blnDisplayPageBreaks = True Then
        ActiveSheet.DisplayAutomaticPageBreaks = True
    End If
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　GetSameHeightAddresses
'[ 概  要 ]　同じ高さの塊ごとのアドレスを配列で取得する
'[ 引  数 ]　選択された領域
'[ 戻り値 ]　アドレスの配列
'*****************************************************************************
Public Function GetSameHeightAddresses(ByRef objSelection As Range) As Collection
    Dim i           As Long
    Dim objRange    As Range
    Dim lngLastCell As Long
    Dim lngLastRow  As Long    '使用されている最後の行
    Dim objWkRange  As Range
    
    Set GetSameHeightAddresses = New Collection
    
    '使用されている最後の行
    lngLastRow = Cells.SpecialCells(xlCellTypeLastCell).Row
   
    '***********************************************
    '使用された最後の行以前の領域の設定
    '***********************************************
    Set objWkRange = IntersectRange(Range(Rows(1), Rows(lngLastRow)), objSelection)
    If Not (objWkRange Is Nothing) Then
        'エリアの数だけループ
        For Each objRange In objWkRange.Areas
            i = objRange.Row
            lngLastCell = Application.WorksheetFunction.Min(i + objRange.Rows.Count - 1, lngLastRow)
                    
            '同じ高さの塊ごとに行高を判定する
            While i <= lngLastCell
                '同じ高さの行のアドレスを保存
                 Call GetSameHeightAddresses.Add(GetSameHeightAddress(i, lngLastCell))
            Wend
        Next objRange
    End If
    
    '***********************************************
    '使用された最後の行以降の領域の設定
    '***********************************************
    If lngLastRow = Rows.Count Then
        Exit Function
    End If
    '使用された最後の行以降の領域
    Set objWkRange = IntersectRange(Range(Rows(lngLastRow + 1), Rows(Rows.Count)), objSelection)
    If objWkRange Is Nothing Then
        Exit Function
    End If
    
    '使用された最後の行以降の可視領域の設定
    Dim objWkRange1 As Range
    Set objWkRange1 = GetVisibleCells(objWkRange)
    If objWkRange1 Is Nothing Then
        '使用された最後の行以降の不可視領域の設定
        Call GetSameHeightAddresses.Add(objWkRange.Address)
        Exit Function
    Else
        '使用された最後の行以降の可視領域の設定
        Call GetSameHeightAddresses.Add(objWkRange1.Address)
        If objWkRange1.Address = objWkRange.Address Then
            Exit Function
        End If
    End If
        
    '***********************************************
    '使用された最後の行以降の不可視領域の設定
    '***********************************************
    Dim lngLastRow2 As Long   '使用されている最後の行以降で選択された可視の行
    lngLastRow2 = Rows.Count
    'エリアの数だけループ
    For Each objRange In objWkRange1.Areas
        If objRange.Row < lngLastRow2 Then
            lngLastRow2 = objRange.Row
        End If
    Next objRange
    
    '使用されている最後の行以降の不可視領域の設定
    Set objWkRange = IntersectRange(Range(Rows(lngLastRow + 1), Rows(lngLastRow2 - 1)), objSelection)
    Call GetSameHeightAddresses.Add(objWkRange.Address)
End Function

'*****************************************************************************
'[ 関数名 ]　GetSameHeightAddress
'[ 概  要 ]　連続する行でlngRowと同じ高さの行を表わすアドレスを取得する
'[ 引  数 ]　lngRow:最初の行(実行後は次の行)、lngLastCell:検索の最後の行
'[ 戻り値 ]　なし
'*****************************************************************************
Private Function GetSameHeightAddress(ByRef lngRow As Long, ByVal lngLastCell As Long) As String
    Dim lngPixel As Long
    Dim i As Long
    lngPixel = Rows(lngRow).Height / 0.75
    
    For i = lngRow + 1 To lngLastCell
        If (Rows(i).Height / 0.75) <> lngPixel Then
            Exit For
        End If
    Next i
    GetSameHeightAddress = Range(Rows(lngRow), Rows(i - 1)).Address
    lngRow = i
End Function

'*****************************************************************************
'[ 関数名 ]　ChangeShapeHeight
'[ 概  要 ]　図形のサイズ変更
'[ 引  数 ]　lngSize:変更サイズ
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ChangeShapeHeight(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim objGroups      As ShapeRange
    Dim blnFitGrid     As Boolean
    
    'アンドゥ用に元のサイズを保存する
    Application.ScreenUpdating = False
    Call SaveUndoInfo(E_ShapeSize, Selection.ShapeRange)
    
    '回転しているものをグループ化する
    Set objGroups = GroupSelection(Selection.ShapeRange)
    
    '[Shift]Keyが押下されていれば、枠線に合わせて変更する
    If GetKeyState(vbKeyShift) < 0 Then
        blnFitGrid = True
    End If
    
    '図形のサイズを変更
    Call ChangeShapesHeight(objGroups, lngSize, blnFitGrid)
    
    '回転している図形のグループ化を解除し元の図形を選択する
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　ChangeShapesHeight
'[ 概  要 ]　図形のサイズ変更
'[ 引  数 ]　objShapes:図形
'            lngSize:変更サイズ(Pixel)
'            blnFitGrid:枠線にあわせるか
'            blnTopLeft:左または上方向に変化させる
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ChangeShapesHeight(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean, Optional ByVal blnTopLeft As Boolean = False)
    Dim objShape     As Shape
    Dim lngTop       As Long
    Dim lngBottom    As Long
    Dim lngOldHeight As Long
    Dim lngNewHeight As Long
    Dim lngNewTop    As Long
    Dim lngNewBottom As Long
    
    '図形の数だけループ
    For Each objShape In objShapes
        lngOldHeight = Round(objShape.Height / 0.75)
        lngTop = Round(objShape.Top / 0.75)
        lngBottom = Round((objShape.Top + objShape.Height) / 0.75)
        
        '枠線にあわせるか
        If blnFitGrid = True Then
            If blnTopLeft = True Then
                If lngSize > 0 Then
                    lngNewTop = GetTopGrid(lngTop, objShape.TopLeftCell.EntireRow)
                Else
                    lngNewTop = GetBottomGrid(lngTop, objShape.TopLeftCell.EntireRow)
                End If
                lngNewHeight = lngBottom - lngNewTop
            Else
                If lngSize < 0 Then
                    lngNewBottom = GetTopGrid(lngBottom, objShape.BottomRightCell.EntireRow)
                Else
                    lngNewBottom = GetBottomGrid(lngBottom, objShape.BottomRightCell.EntireRow)
                End If
                lngNewHeight = lngNewBottom - lngTop
            End If
            If lngNewHeight < 0 Then
                lngNewHeight = 0
            End If
        Else
            'ピクセル単位の変更をする
            If lngOldHeight + lngSize >= 0 Then
                If blnTopLeft = True And lngTop = 0 And lngSize > 0 Then
                    lngNewHeight = lngOldHeight
                Else
                    lngNewHeight = lngOldHeight + lngSize
                End If
            Else
                lngNewHeight = lngOldHeight
            End If
        End If
    
        If lngSize > 0 And blnTopLeft = True Then
            objShape.Top = (lngBottom - lngNewHeight) * 0.75
        End If
        objShape.Height = lngNewHeight * 0.75
        
        'Excel2007のバグ対応
        If Round(objShape.Height / 0.75) <> lngNewHeight Then
            objShape.Height = (lngNewHeight + lngSize) * 0.75
        End If
        
        If Round(objShape.Height / 0.75) <> lngOldHeight Then
            If blnTopLeft = True Then
                objShape.Top = (lngBottom - lngNewHeight) * 0.75
            Else
                objShape.Top = lngTop * 0.75
            End If
        End If
    Next objShape
End Sub

'*****************************************************************************
'[ 関数名 ]　GetTopGrid
'[ 概  要 ]　入力の位置の上の枠線の位置を取得(単位ピクセル)
'[ 引  数 ]　lngPos:位置(単位ピクセル)
'            objRow: lngPosを含む行
'[ 戻り値 ]　図形の上側の枠線の位置(単位ピクセル)
'*****************************************************************************
Public Function GetTopGrid(ByVal lngPos As Long, ByRef objRow As Range) As Long
    Dim i      As Long
    Dim lngTop As Long
    
    If lngPos <= Round(Rows(2).Top / 0.75) Then
        GetTopGrid = 0
        Exit Function
    End If
    
    For i = objRow.Row To 1 Step -1
        lngTop = Round(Rows(i).Top / 0.75)
        If lngTop < lngPos Then
            GetTopGrid = lngTop
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[ 関数名 ]　GetBottomGrid
'[ 概  要 ]　入力の位置の下の枠線の位置を取得(単位ピクセル)
'[ 引  数 ]　lngPos:位置(単位ピクセル)
'            objRow: lngPosを含む行
'[ 戻り値 ]　図形の下側の枠線の位置(単位ピクセル)
'*****************************************************************************
Public Function GetBottomGrid(ByVal lngPos As Long, ByRef objRow As Range) As Long
    Dim i         As Long
    Dim lngBottom As Long
    Dim lngMax    As Long
    
    lngMax = Round((Rows(Rows.Count).Top + Rows(Rows.Count).Height) / 0.75)
    
    If lngPos >= Round(Rows(Rows.Count).Top / 0.75) Then
        GetBottomGrid = lngMax
        Exit Function
    End If
    
    For i = objRow.Row + 1 To Rows.Count
        lngBottom = Round(Rows(i).Top / 0.75)
        If lngBottom > lngPos Then
            GetBottomGrid = lngBottom
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[ 関数名 ]　GetVisibleCells
'[ 概  要 ]　可視セルを取得
'[ 引  数 ]　選択セル
'[ 戻り値 ]　可視セル
'*****************************************************************************
Private Function GetVisibleCells(ByRef objRange As Range) As Range
On Error GoTo ErrHandle
    Dim objCells As Range
    Set objCells = objRange.SpecialCells(xlCellTypeVisible)
    
    '列の非表示は選択する
    Set GetVisibleCells = Union(objCells.EntireRow, objCells.EntireRow)
    Set GetVisibleCells = IntersectRange(GetVisibleCells, objRange)
Exit Function
ErrHandle:
    Set GetVisibleCells = Nothing
End Function

'*****************************************************************************
'[ 関数名 ]　MoveRowBorder
'[ 概  要 ]　行の境界線を上下に移動する
'[ 引  数 ]　lngSize : 移動サイズ(単位:ピクセル)
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveRowBorder(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim strSelection      As String
    Dim objRange          As Range
    Dim lngPixel(1 To 2)  As Long   '先頭行と最終行のサイズ
    Dim k                 As Long   '最終行の行番号
    
    '[Ctrl]Keyが押下されていれば、移動高さを5倍にする
    If GetKeyState(vbKeyControl) < 0 Then
        lngSize = lngSize * 5
    End If
    
    '選択されているオブジェクトを判定
    Select Case CheckSelection()
    Case E_Other
        Exit Sub
    Case E_Shape
'        Call MoveShape(lngSize)
        Exit Sub
    End Select
    
    strSelection = Selection.Address
    Set objRange = Selection

    '選択エリアが複数なら対象外
    If objRange.Areas.Count <> 1 Then
        Call MsgBox("このコマンドは複数の選択範囲に対して実行できません。", vbExclamation)
        Exit Sub
    End If

    '選択行が１行なら対象外
    If objRange.Rows.Count = 1 Then
        Exit Sub
    End If
    
    '最終行の行番号
    k = objRange.Rows.Count
    
    '変更後のサイズ
    lngPixel(1) = HeightToPixel(objRange.Rows(1).RowHeight) + lngSize '先頭行
    lngPixel(2) = HeightToPixel(objRange.Rows(k).RowHeight) - lngSize '最終行
    
    'サイズのチェック
    If (0 <= lngPixel(1) And lngPixel(1) <= HeightToPixel(MaxRowHeight)) And _
       (0 <= lngPixel(2) And lngPixel(2) <= HeightToPixel(MaxRowHeight)) Then
    Else
        Exit Sub
    End If
    
    '***********************************************
    'サイズの変更
    '***********************************************
    Application.ScreenUpdating = False
    'アンドゥ用に元のサイズを保存する
    Dim colAddress  As New Collection
    Call colAddress.Add(objRange.Rows(1).Address)
    Call colAddress.Add(objRange.Rows(k).Address)
    Call SaveUndoInfo(E_RowSize, Selection, colAddress)
    
    'サイズの変更
    objRange.Rows(1).RowHeight = PixelToHeight(lngPixel(1))
    objRange.Rows(k).RowHeight = PixelToHeight(lngPixel(2))
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

''*****************************************************************************
''[ 関数名 ]　MoveShape
''[ 概  要 ]　図形を上下に移動する
''[ 引  数 ]　lngSize：移動サイズ
''[ 戻り値 ]　なし
''*****************************************************************************
'Private Sub MoveShape(ByVal lngSize As Long)
'On Error GoTo ErrHandle
'    Dim blnFitGrid  As Boolean
'    Dim objGroups   As ShapeRange
'
'    'アンドゥ用に元のサイズを保存する
'    Application.ScreenUpdating = False
'    Call SaveUndoInfo(E_ShapeSize, Selection.ShapeRange)
'
'    '[Shift]Keyが押下されていれば、枠線に合わせて変更する
'    If GetKeyState(vbKeyShift) < 0 Then
'        blnFitGrid = True
'    End If
'
'    '枠線にあわせるか
'    If blnFitGrid = True Then
'        '回転している図形をグループ化する
'        Set objGroups = GroupSelection(Selection.ShapeRange)
'
'        '図形を左右に移動する
'        Call MoveShapesUD(objGroups, lngSize, blnFitGrid)
'
'        '回転している図形のグループ化を解除し元の図形を選択する
'        Call UnGroupSelection(objGroups).Select
'    Else
'        '図形を左右に移動する
'        Call MoveShapesUD(Selection.ShapeRange, lngSize, blnFitGrid)
'    End If
'
'    Call SetOnUndo
'Exit Sub
'ErrHandle:
'    Call MsgBox(Err.Description, vbExclamation)
'End Sub

'*****************************************************************************
'[ 関数名 ]　DistributeRowsHeight
'[ 概  要 ]　選択された行の高さを揃える
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub DistributeRowsHeight()
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objRange     As Range
    Dim lngRowCount  As Long    '選択された行の数
    Dim dblHeight    As Double  '選択された行の高さの合計
    Dim objSelection As Range   '選択されたすべての行のコレクション
    Dim strSelection As String
    Dim objVisible   As Range   '選択された可視の行
    
    '選択されているオブジェクトを判定
    Select Case CheckSelection()
    Case E_Other
        Exit Sub
    Case E_Shape
        Call DistributeShapeHeight
        Exit Sub
    End Select
    
    '選択範囲のRowsの和集合を取り重複行を排除する
    strSelection = Selection.Address
    Set objSelection = Union(Selection.EntireRow, Selection.EntireRow)
    
    '選択範囲の可視部分を取出す
    Set objVisible = GetVisibleCells(objSelection)
    
    'すべて非表示の時
    If objVisible Is Nothing Then
        Exit Sub
    End If
    
    '非表示の行がある時
    If objSelection.Address <> objVisible.Address Then
        If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
            Select Case MsgBox("非表示の行を対象としますか？", vbYesNoCancel + vbQuestion + vbDefaultButton2)
            Case vbNo
                Set objSelection = objVisible
            Case vbCancel
                Exit Sub
            End Select
        Else
            Set objSelection = objVisible
        End If
    End If
    
    'エリアの数だけループ
    For Each objRange In objSelection.Areas
        dblHeight = dblHeight + GetHeight(objRange)
        lngRowCount = lngRowCount + objRange.Rows.Count
    Next objRange
    
    If lngRowCount = 1 Then
        Exit Sub
    End If
    
    '***********************************************
    'サイズの変更
    '***********************************************
    Application.ScreenUpdating = False
    'アンドゥ用に元のサイズを保存する
    Call SaveUndoInfo(E_RowSize, Range(strSelection), GetSameHeightAddresses(objSelection))
    objSelection.RowHeight = dblHeight / lngRowCount
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　DistributeShapeHeight
'[ 概  要 ]　選択された図形の高さを揃える
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub DistributeShapeHeight()
On Error GoTo ErrHandle
    If Selection.ShapeRange.Count = 1 Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    'アンドゥ用に元のサイズを保存する
    Call SaveUndoInfo(E_ShapeSize, Selection.ShapeRange)
    
    '回転している図形をグループ化する
    Dim objGroups As ShapeRange
    Set objGroups = GroupSelection(Selection.ShapeRange)
    
    Call DistributeShapesHeight(objGroups)
    
    '回転している図形のグループ化を解除し元の図形を選択する
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　DistributeShapesHeight
'[ 概  要 ]　図形の高さを揃える
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub DistributeShapesHeight(ByRef objShapeRange As ShapeRange)
    Dim objShape   As Shape
    Dim dblHeight  As Double
    
    '図形の数だけループ
    For Each objShape In objShapeRange
        dblHeight = dblHeight + objShape.Height
    Next objShape
    With objShapeRange
        .Height = Round(dblHeight / .Count / 0.75) * 0.75
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　GetHeight
'[ 概  要 ]　選択エリアの高さを取得
'            Heightプロパティは32767以上の高さを計算出来ないため
'[ 引  数 ]　高さを取得するエリア
'[ 戻り値 ]　高さ(Heightプロパティ)
'*****************************************************************************
Private Function GetHeight(ByRef objRange As Range) As Double
    With objRange
        GetHeight = .Rows(.Rows.Count).Top - .Top + .Rows(.Rows.Count).Height
    End With
    
'    Dim lngCount   As Long
'    Dim lngHalf    As Long
'    Dim MaxHeight  As Double '高さの最大値
'
'    MaxHeight = 32767 * 0.75
'    If objRange.Height < MaxHeight Then
'        GetHeight = objRange.Height
'    Else
'        With objRange
'            '前半＋後半の高さを合計
'            lngCount = .Rows.Count
'            lngHalf = lngCount / 2
'            GetHeight = GetHeight(Range(.Rows(1), .Rows(lngHalf))) + _
'                        GetHeight(Range(.Rows(lngHalf + 1), .Rows(lngCount)))
'        End With
'    End If
End Function

'*****************************************************************************
'[ 関数名 ]　MergeCellsAsRow
'[ 概  要 ]　縦方向に結合(複数行に値がある時は改行で連結)
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MergeCellsAsRow()
On Error GoTo ErrHandle
    Dim i            As Long
    Dim strSelection As String
    Dim objWkRange   As Range
    Dim objRange     As Range
    Dim objMergeCell As Range
    Dim strValues    As String
    Dim lngCalculation As Long
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    strSelection = Selection.Address
    lngCalculation = Application.Calculation
    
    '***********************************************
    '重複領域のチェック
    '***********************************************
    If CheckDupRange(Range(strSelection)) = True Then
        Call MsgBox("選択されている領域に重複があります", vbExclamation)
        Exit Sub
    End If
    
    '***********************************************
    '変更
    '***********************************************
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    'アンドゥ用に元の状態を保存する
    Call SaveUndoInfo(E_MergeCell, Range(strSelection))
    Call Range(strSelection).UnMerge
    
    'エリアの数だけループ
    For Each objRange In Range(strSelection).Areas
        '列の数だけループ
        For i = 1 To objRange.Columns.Count
            
            '複数のセルに値がある時連結する
            Set objMergeCell = objRange.Columns(i)
            If WorksheetFunction.CountA(objMergeCell) > 1 Then
                strValues = Replace$(GetRangeText(objMergeCell), vbTab, " ")
            Else
                strValues = ""
            End If
            
            'セルを結合する
            Call objMergeCell.Merge
            
            '連結した値を再設定する
            If strValues <> "" Then
                objMergeCell.Value = strValues
            End If
        Next i
    Next objRange

    Call Range(strSelection).Select
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
'[ 関数名 ]　SplitRow
'[ 概  要 ]　行を分割する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SplitRow()
On Error GoTo ErrHandle
    Dim i               As Long
    Dim objRange        As Range
    Dim lngPixel        As Double  '１行の高さ
    Dim lngSplitCount   As Long    '分割数
    Dim blnCheckInsert  As Boolean
    Dim objNewRow       As Range    '新しい行
    Dim objNewSelection As Range   '分割後の選択範囲
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objRange = Selection
    
    '選択エリアが複数なら対象外
    If objRange.Areas.Count <> 1 Then
        Call MsgBox("このコマンドは複数の選択範囲に対して実行できません。", vbExclamation)
        Exit Sub
    End If
    
    '複数行選択なら対象外
    If objRange.Rows.Count <> 1 Then
        Call MsgBox("このコマンドは複数の選択行に対して実行できません。", vbExclamation)
        Exit Sub
    End If
    
    '元の高さ
    lngPixel = HeightToPixel(objRange.EntireRow.RowHeight)
    
    '****************************************
    '分割数を選択させる
    '****************************************
    Dim enmSizeType As ESizeType  '選択されたタイプ
    With frmSplitCount
        'フォームを表示
        Call .Show
    
        'キャンセル時
        If blnFormLoad = False Then
            Exit Sub
        End If
        
        lngSplitCount = .Count
        blnCheckInsert = .CheckInsert
        Call Unload(frmSplitCount)
    End With
    
    '****************************************
    '分割開始
    '****************************************
    Dim blnDisplayPageBreaks As Boolean
        
    '高速化のため改ページを非表示にする
    If ActiveSheet.DisplayAutomaticPageBreaks = True Then
        blnDisplayPageBreaks = True
        ActiveSheet.DisplayAutomaticPageBreaks = False
    End If
    Application.ScreenUpdating = False
    
    'アンドゥ用に元の状態を保存する
    Call SaveUndoInfo(E_SplitRow, objRange, lngSplitCount)
    If blnCheckInsert = False Then
        Call SetPlacement
    End If
    
    '選択行の下に１行挿入
    Call objRange(2, 1).EntireRow.Insert
    
    '新しい行
    Set objNewRow = objRange(2, 1).EntireRow
    
    '*************************************************
    '罫線を整える
    '*************************************************
    '挿入列の１セル毎に罫線をコピーする
    Call CopyBorders(False, objNewRow)
    
    '境界線を消す
    With Range(objRange, objNewRow).EntireRow
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    '*************************************************
    '分割を繰返す
    '*************************************************
    '分割数だけ、行を挿入する
    For i = 3 To lngSplitCount
        Call objNewRow.EntireRow.Insert
    Next i
    
    Set objNewSelection = objRange.Resize(lngSplitCount)
    
    '*************************************************
    '高さの整備
    '*************************************************
    '新しい高さに設定
    If blnCheckInsert = False Then
        If lngSplitCount = 2 Then
            objRange.EntireRow.RowHeight = PixelToHeight(Int(lngPixel / 2 + 0.5))
            objNewRow.EntireRow.RowHeight = PixelToHeight(Int(lngPixel / 2))
        Else
            With Range(objRange, objNewRow).EntireRow
                If lngPixel < lngSplitCount Then
                    .RowHeight = PixelToHeight(1)
                Else
                    .RowHeight = PixelToHeight(lngPixel / lngSplitCount)
                End If
            End With
        End If
    End If
    
    '*************************************************
    '後処理
    '*************************************************
    Call objNewSelection.Select
    If blnCheckInsert = False Then
        Call ResetPlacement
    End If
    Call SetOnUndo
    
    If blnDisplayPageBreaks = True Then
        ActiveSheet.DisplayAutomaticPageBreaks = True
    End If
    Application.ScreenUpdating = True
Exit Sub
ErrHandle:
    If blnDisplayPageBreaks = True Then
        ActiveSheet.DisplayAutomaticPageBreaks = True
    End If
    Call MsgBox(Err.Description, vbExclamation)
    If blnCheckInsert = False Then
        Call ResetPlacement
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　EraseRow
'[ 概  要 ]　選択された行を消去する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub EraseRow()
On Error GoTo ErrHandle
    Dim i                 As Long
    Dim objSelection      As Range
    Dim strSelection      As String
    Dim objRange          As Range
    Dim lngTopRow         As Long  '上隣の行番号
    Dim lngBottomRow      As Long  '下隣の行番号
    Dim objColumn         As Range '終了時に選択させる列
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = Selection
    strSelection = objSelection.Address
    Set objRange = objSelection.EntireRow
    
    '終了時に選択させる列
    Set objColumn = objSelection.EntireColumn
    
    '選択エリアが複数なら対象外
    If objSelection.Areas.Count <> 1 Then
        Call MsgBox("このコマンドは複数の選択範囲に対して実行できません。", vbExclamation)
        Exit Sub
    End If
    
    If objSelection.Rows.Count = Rows.Count Then
        Call MsgBox("このコマンドはすべての行の選択に対して実行できません。", vbExclamation)
        Exit Sub
    End If
    
    '上隣の行番号
    lngTopRow = objRange.Row - 1
    '下隣の行番号
    lngBottomRow = objRange.Row + objRange.Rows.Count
    
    '****************************************
    '消去のパターンを選択させる
    '****************************************
    Dim enmSelectType As ESelectType  '消去パターン
    Dim blnHidden   As Boolean        '非表示とするかどうか
    With frmEraseSelect
        'シートの１行目が選択されている時
        If lngTopRow = 0 Then
            Call .SetEnabled(E_Back)
            Call .SetEnabled(E_Middle)
            .SelectType = E_Front
        'シートの最終行が選択されている時
        ElseIf lngBottomRow > Rows.Count Then
            Call .SetEnabled(E_Front)
            Call .SetEnabled(E_Middle)
            .SelectType = E_Back
        Else
            .SelectType = E_Back
        End If
        
        '選択フォームを表示
        Call .Show
    
        '取消時
        If blnFormLoad = False Then
            Exit Sub
        End If
        
        enmSelectType = .SelectType
        blnHidden = .Hidden
        Call Unload(frmEraseSelect)
    End With
    
    '****************************************
    '値のある行を削除するかどうか判定
    '****************************************
    Dim objValueCell As Range '削除される列で値の含まれるセル
    If blnHidden = False Then
        Set objValueCell = SearchValueCell(objRange)
        '削除される行で値の含まれるセルがあった時
        If Not (objValueCell Is Nothing) Then
            Call objValueCell.Select
            If MsgBox("値の入力されているセルが削除されますが、よろしいですか？", vbOKCancel + vbQuestion + vbDefaultButton2) = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    
    '****************************************
    'Undo用に列幅を保存するための情報取得
    '****************************************
    Dim colAddress   As New Collection
    Set colAddress = GetSameHeightAddresses(Range(strSelection).EntireRow)
    
    '****************************************
    '選択行の上下の行を待避
    '****************************************
    Dim objRow(0 To 1)    As Range  '行の高さを変える行
    Select Case enmSelectType
    Case E_Front
        Set objRow(0) = Rows(lngBottomRow)
        Call colAddress.Add(objRow(0).Address)
    Case E_Back
        Set objRow(0) = Rows(lngTopRow)
        Call colAddress.Add(objRow(0).Address)
    Case E_Middle
        Set objRow(0) = Rows(lngBottomRow)
        Set objRow(1) = Rows(lngTopRow)
        Call colAddress.Add(objRow(0).Address)
        Call colAddress.Add(objRow(1).Address)
    End Select
    
    '****************************************
    '行を消去
    '****************************************
    Dim lngPixel  As Long   '消去される行の高さ(単位:ピクセル)
    Application.ScreenUpdating = False
    
    'アンドゥ用に元の状態を保存する
    If blnHidden = True Then
        Call SaveUndoInfo(E_RowSize, Range(strSelection), colAddress)
    Else
        Call SaveUndoInfo(E_EraseRow, Range(strSelection), colAddress)
    End If
    
    '図形は移動させない
    Call SetPlacement
    
    '消去される行の高さを保存
    lngPixel = objRange.Height / 0.75
    
    If blnHidden = True Then
        '非表示
        objRange.Hidden = True
    Else
        '罫線をコピーする
        Call CopyBorder(objRange)
        '削除
        Call objRange.Delete(xlShiftUp)
    End If
    
    '****************************************
    '行高を拡大
    '****************************************
    Dim lngWkPixel  As Long
    Select Case enmSelectType
    Case E_Front, E_Back
        objRow(0).RowHeight = WorksheetFunction.Min(MaxRowHeight, PixelToHeight(objRow(0).Height / 0.75 + lngPixel))
    Case E_Middle
        objRow(0).RowHeight = WorksheetFunction.Min(MaxRowHeight, PixelToHeight(objRow(0).Height / 0.75 + Int(lngPixel / 2 + 0.5)))
        objRow(1).RowHeight = WorksheetFunction.Min(MaxRowHeight, PixelToHeight(objRow(1).Height / 0.75 + Int(lngPixel / 2)))
    End Select
    
    '****************************************
    'セルを選択
    '****************************************
    Select Case enmSelectType
    Case E_Front, E_Back
        IntersectRange(objColumn, objRow(0)).Select
    Case E_Middle
        IntersectRange(objColumn, UnionRange(objRow(0), objRow(1))).Select
    End Select
    Call ResetPlacement
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
    Call ResetPlacement
End Sub

'*****************************************************************************
'[ 関数名 ]　CopyBorder
'[ 概  要 ]　新しい行に元の行の罫線をコピーする
'[ 引  数 ]　objRange:削除行
'[ 戻り値 ]　0:成功、1以上:整備を中断した行
'*****************************************************************************
Private Sub CopyBorder(ByRef objRange As Range)
    If objRange.Row > 1 Then
        '１セル毎に罫線をコピーする
        With objRange
            Call CopyBottomBorder(.Rows(.Rows.Count), .Rows(0))
        End With
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　CopyBorders
'[ 概  要 ]　新しい行に元の行の罫線をコピーする
'[ 引  数 ]　blnMerge:すべての行を前の行と結合するかどうか、objRange:挿入行
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub CopyBorders(ByVal blnMerge As Boolean, ByRef objRange As Range)
    Dim i                 As Long
    Dim objCell           As Range   'コピー先のセル
    Dim objOrgCell        As Range   'コピー元のセル
    Dim udtBorder(0 To 2) As TBorder '罫線の種類(左･右･下)
    Dim lngLast    As Long
    
    'すべての列を整備するか、最後のセルまで整備すれば終了する
    lngLast = WorksheetFunction.Min(Cells.SpecialCells(xlCellTypeLastCell).Column, _
                              objRange.Column + objRange.Columns.Count - 1)
    '1列毎にループ
    For i = 1 To lngLast
        Set objCell = objRange.Columns(i)
        
        '新しい行のセルが結合セルか
        If objCell.MergeCells = False Then
            '元のセルを代入
            Set objOrgCell = objCell.Offset(-1, 0)
            
            '罫線の種類を保存(左･右･下)
            With objOrgCell.MergeArea.Borders
                udtBorder(0) = GetBorder(.Item(xlEdgeLeft))
                udtBorder(1) = GetBorder(.Item(xlEdgeRight))
                udtBorder(2) = GetBorder(.Item(xlEdgeBottom))
            End With
            
            '元のセルが結合セルか
            If blnMerge = True Or objOrgCell.MergeCells = True Then
                With objOrgCell.MergeArea
                    '挿入行のセルを結合
                    Call .Resize(.Rows.Count + 1).Merge
                End With
            End If
            
            With objCell.MergeArea.Borders
                '罫線を書く(左･右･下)
                Call SetBorder(udtBorder(0), .Item(xlEdgeLeft))
                Call SetBorder(udtBorder(1), .Item(xlEdgeRight))
                Call SetBorder(udtBorder(2), .Item(xlEdgeBottom))
            End With
        End If
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　CopyBottomBorder
'[ 概  要 ]　削除行の下端の罫線をコピーする
'[ 引  数 ]　objFromRange：コピ－元の行
'　　　　　　objToRange：コピー先の行
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub CopyBottomBorder(ByRef objFromRange As Range, ByRef objToRange As Range)
    Dim i          As Long
    Dim udtBorder  As TBorder    '下端の罫線の種類
    Dim lngLast    As Long
    
    'すべての列を整備するか、最後のセルまで整備すれば終了する
    lngLast = WorksheetFunction.Min(Cells.SpecialCells(xlCellTypeLastCell).Column, _
                              objFromRange.Column + objFromRange.Columns.Count - 1)
    '1列毎にループ
    For i = 1 To lngLast
        'コピ－元のセルの罫線の種類を保存
        udtBorder = GetBorder(objFromRange.Columns(i).Borders(xlEdgeBottom))
        'コピー先のセルにコピー
        Call SetBorder(udtBorder, objToRange.Columns(i).Borders(xlEdgeBottom))
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　ShowHeight
'[ 概  要 ]　エリア毎に選択された行の高さを一覧で表示する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ShowHeight()
On Error GoTo ErrHandle
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Call frmSizeList.Initialize(E_Row)
    Call frmSizeList.Show
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　MoveCellBorder
'[ 概  要 ]　行の境界のセルを上下に移動する
'[ 引  数 ]　-1:上に移動、1:下に移動
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveCellBorder(ByVal lngUpDown As Long)
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objSelection As Range
    Dim objWkRange   As Range
    Dim lngRowCount  As Long  '選択領域の行数
    Dim blnCopyObjectsWithCells As Boolean
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() = E_Range Then
        Set objSelection = Selection
    Else
        Exit Sub
    End If
    
    '選択エリアが複数なら対象外
    If objSelection.Areas.Count <> 1 Then
        Call MsgBox("このコマンドは複数の選択範囲に対して実行できません。", vbExclamation)
        Exit Sub
    End If
    
    '******************************************************************
    '縮小不可能なセルがないかチェック
    '******************************************************************
    lngRowCount = objSelection.Rows.Count
    Dim objChkRange(0 To 2) As Range

    If objSelection.Rows.Count = 1 Then
        Exit Sub
    End If

    '上に移動する時
    If lngUpDown < 0 Then
        Set objChkRange(1) = ArrangeRange(Range(objSelection.Rows(1), objSelection.Rows(2)))
        Set objChkRange(2) = ArrangeRange(objSelection.Rows(2))
    Else '下に移動する時
        Set objChkRange(1) = ArrangeRange(Range(objSelection.Rows(lngRowCount - 1), objSelection.Rows(lngRowCount)))
        Set objChkRange(2) = ArrangeRange(objSelection.Rows(lngRowCount - 1))
    End If
    
    '移動する境界がない時
    If MinusRange(objSelection, objChkRange(1)) Is Nothing Then
        Exit Sub
    End If
    
    Set objChkRange(0) = MinusRange(objChkRange(1), objChkRange(2))
    If Not (objChkRange(0) Is Nothing) Then
        Call objChkRange(0).Select
        Call MsgBox("これ以上縮小できないセルがあります")
        Call objSelection.Select
        Exit Sub
    End If
    
    '****************************************
    '移動開始
    '****************************************
    '図形をコピーの対象外にする
    Application.CopyObjectsWithCells = False
    Application.ScreenUpdating = False
    'アンドゥ用に元の状態を保存する
    Call SaveUndoInfo(E_MergeCell, objSelection)
    
    '****************************************
    '元の領域を、"Workarea1"シートにコピー
    '****************************************
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    With ThisWorkbook.Worksheets("Workarea1")
        'コメントをまともな位置に配置出来るように、
        '元の幅と高さをコピーするため、シート全体をコピーした後クリア
        Call ActiveSheet.Cells.Copy(.Cells)
        Call .Cells.Clear
        
        Set objWkRange = .Range(objSelection.Address)
        
        '領域をコピー
        Call objSelection.Copy(objWkRange)
    End With
    
    '****************************************
    '境界を移動する
    '****************************************
    If lngUpDown < 0 Then
        '上に移動する
        Call CopyBottomBorder(objWkRange.Rows(2), objWkRange.Rows(1))
        Call objWkRange.Rows(2).Delete(xlUp)
        Call CopyBorders(True, objWkRange.Rows(lngRowCount))
    Else
        '下に移動する
        Call CopyBottomBorder(objWkRange.Rows(lngRowCount), objWkRange.Rows(lngRowCount - 1))
        Call objWkRange.Rows(lngRowCount).Delete(xlUp)
        Call objWkRange.Rows(2).Insert(xlDown)
        Call CopyBorders(True, objWkRange.Rows(2))
    End If
    
    Call objWkRange.Worksheet.Range(objSelection.Address).Copy(objSelection)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    Call SetOnUndo
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
Exit Sub
ErrHandle:
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
    Call MsgBox(Err.Description, vbExclamation)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
End Sub

'*****************************************************************************
'[ 関数名 ]　AutoRowFit
'[ 概  要 ]　行の高さを文字の高さにあわせる
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub AutoRowFit()
On Error GoTo ErrHandle
    Dim objSelection  As Range
    Dim objSelection2 As Range
    Dim objUsedRange  As Range
    Dim blnVisible    As Boolean
    Dim dblNewHeight  As Double
    Dim i As Long
    
    '選択されているオブジェクトを判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = IntersectRange(Selection, Range(Cells(1, 1), Cells.SpecialCells(xlCellTypeLastCell)))
    If objSelection Is Nothing Then
        Exit Sub
    End If
    
    '選択エリアが複数なら対象外
    If objSelection.Areas.Count <> 1 Then
        Call MsgBox("このコマンドは複数の選択範囲に対して実行できません。", vbExclamation)
        Exit Sub
    End If
    
    '非表示の行を対象とするか確認？
    Dim objVisible    As Range
    Dim objNonVisible As Range
    Set objVisible = GetVisibleCells(objSelection)
    Set objNonVisible = MinusRange(objSelection, objVisible)
    If Not (objNonVisible Is Nothing) Then
        If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
            Select Case MsgBox("非表示のセルを対象としますか？", vbYesNoCancel + vbQuestion + vbDefaultButton2)
            Case vbYes
                blnVisible = True
                Set objSelection2 = objSelection
            Case vbNo
                blnVisible = False
                Set objSelection2 = objVisible
            Case vbCancel
                Exit Sub
            End Select
        Else
            blnVisible = False
            Set objSelection2 = objVisible
        End If
    Else
        blnVisible = True
        Set objSelection2 = objSelection
    End If
    
    'WorkSheetの標準スタイルを実行シートとあわせる
    Call SetPixelInfo
    
    '***********************************************
    '実行
    '***********************************************
    Application.ScreenUpdating = False
    
    'アンドゥ用に元のサイズを保存する
    Call SaveUndoInfo(E_RowSize, objSelection, GetSameHeightAddresses(objSelection.EntireRow))
    
    '標準コマンドで高さを適正化(編集でFontを変えても自動で調製されるようにする対応)
    Call objSelection2.Rows.AutoFit
    
    If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
    Else
        Call SetOnUndo
        Exit Sub
    End If
    
    '動作が非常に遅くなるための対応
    If objSelection.Rows.Count > 100 Then
        Call SetOnUndo
        Exit Sub
    End If
    
    '*********************************************************
    'WorkSheetを利用し、行の高さを適正化する
    '*********************************************************
    Dim objWorksheet As Worksheet
    Set objWorksheet = ThisWorkbook.Worksheets("Workarea1")
    Call DeleteSheet(objWorksheet)
    With objWorksheet
        .Columns.ColumnWidth = 255
        .Range(.Rows(1), .Rows(objSelection.Rows.Count)).Font.Size = 1
        Call objSelection.Copy(.Cells(1, 1))
    End With
    
    '行数分ループ
    Dim objRow     As Range
    Dim objWorkRow As Range
    For i = 1 To objSelection.Rows.Count
        Set objRow = objSelection.Rows(i)
        Set objWorkRow = objWorksheet.Rows(i)
        
        '非表示を対象外にするかどうか
        If blnVisible Or objRow.Hidden = False Then
            '行方向の結合がある行は、標準のAutoFitのみ行う（すでにAutoFitは完了している）
            If IsBorderMerged(objRow) = False Then
                'WorkSheetを利用し、行の高さを適正化
                dblNewHeight = GetFitRow(objRow, objWorkRow)
                '編集でFontを変えても自動で調製されるようために判定する
                If objRow.RowHeight <> dblNewHeight Then
                    objRow.RowHeight = dblNewHeight
                End If
            End If
        End If
    Next
    
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
End Sub

'*****************************************************************************
'[ 関数名 ]　GetFitRow
'[ 概  要 ]　WorkSheetを利用し、行の高さを適正化
'[ 引  数 ]　対象の行(元のシート)、対象の行(ワークシート)
'[ 戻り値 ]　適正化後のRowHeight
'*****************************************************************************
Private Function GetFitRow(ByRef objRow As Range, ByRef objWorkRow As Range) As Double
    Dim objCell        As Range
    Dim dblColumnWidth As Double
    Dim objValueCells As Range
    
    '値の入力されたセルの幅を整備する
    On Error Resume Next
    Set objValueCells = objWorkRow.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If Not (objValueCells Is Nothing) Then
        For Each objCell In objValueCells
            '列の幅をコピーする
            dblColumnWidth = WorksheetFunction.Min(PixelToWidth(objRow.Columns(objCell.Column).MergeArea.Width / 0.75), 255)
            With objCell
                If .ColumnWidth <> dblColumnWidth Then
                    .ColumnWidth = dblColumnWidth
                End If
                Call .UnMerge
            End With
        Next
    End If
    
    '高さを設定
    Call objWorkRow.AutoFit
    
    GetFitRow = objWorkRow.RowHeight
End Function

'*****************************************************************************
'[ 関数名 ]　PixelToHeight
'[ 概  要 ]　高さの単位を変換
'[ 引  数 ]　lngPixel : 高さ(単位:ピクセル)
'[ 戻り値 ]　Height
'*****************************************************************************
Public Function PixelToHeight(ByVal lngPixel As Long) As Double
    PixelToHeight = lngPixel * 0.75
End Function

'*****************************************************************************
'[ 関数名 ]　HeightToPixel
'[ 概  要 ]　高さの単位を変換
'[ 引  数 ]　dblHeight : 高さ
'[ 戻り値 ]　高さ(単位:ピクセル)
'*****************************************************************************
Private Function HeightToPixel(ByVal dblHeight As Double) As Long
    HeightToPixel = Int(dblHeight / 0.75)
End Function

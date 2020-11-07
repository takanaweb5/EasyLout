Attribute VB_Name = "ColChangeTools"
Option Explicit
Option Private Module

Private Type TFont  '標準スタイルのフォントの情報
    Name        As String
    size        As Long
    Bold        As Boolean
    Italic      As Boolean
End Type

Private x1 As Byte '1文字のピクセル
Private x2 As Byte '2文字のピクセル

Private Const MaxColumnWidth = 255  '幅の最大サイズ

'*****************************************************************************
'[ 関数名 ]　ChangeWidth
'[ 概  要 ]　幅の変更
'[ 引  数 ]　lngSize:変更サイズ(単位：ピクセル)
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ChangeWidth(ByVal lngSize As Long)
On Error GoTo ErrHandle
'    Dim lngSize As Long
    
    '[Ctrl]Keyが押下されていれば、移動幅を5倍にする
'    lngSize = CommandBars.ActionControl.Parameter
    If GetKeyState(vbKeyControl) < 0 Then
        lngSize = lngSize * 5
    End If
    
    '選択されているオブジェクトを判定
    Select Case CheckSelection()
    Case E_Range
        Call ChangeColsWidth(lngSize)
    Case E_Shape
        Call ChangeShapeWidth(lngSize)
    End Select
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　MoveVerticalBorder
'[ 概  要 ]　列の境界線を左右に移動する
'[ 引  数 ]　lngSize:変更サイズ(単位：ピクセル)
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveVerticalBorder(ByVal lngSize As Long)
On Error GoTo ErrHandle
'    Dim lngSize As Long
    
    '[Ctrl]Keyが押下されていれば、移動幅を5倍にする
'    lngSize = CommandBars.ActionControl.Parameter
    If GetKeyState(vbKeyControl) < 0 Then
        lngSize = lngSize * 5
    End If
    
    '選択されているオブジェクトを判定
    Select Case CheckSelection()
    Case E_Range
        Call MoveBorder(lngSize)
    Case E_Shape
'        Call MoveShape(lngSize)
        Exit Sub
    End Select
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　ChangeColsWidth
'[ 概  要 ]　複数列サイズ変更
'[ 引  数 ]　lngSize:変更サイズ(単位：ピクセル)
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ChangeColsWidth(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objSelection As Range   '選択されたすべての列
    Dim strSelection As String
    Dim lngWindowView As Long
        
    '選択範囲のColumnsの和集合を取り重複列を排除する
    strSelection = Selection.Address
    Set objSelection = Union(Selection.EntireColumn, Selection.EntireColumn)
    
    '標準プレビューに変更する
    Application.ScreenUpdating = False
    lngWindowView = ActiveWindow.View
    ActiveWindow.View = xlNormalView
    
    '***********************************************
    '非表示の列があるかどうかの判定
    '***********************************************
    Dim objVisible     As Range    '可視Range
    '選択範囲の可視部分を取出す
    Set objVisible = GetVisibleCells(objSelection)
    If objVisible Is Nothing Then
        If lngSize < 0 Then
            ActiveWindow.View = lngWindowView
            Application.ScreenUpdating = True
            Call MsgBox("これ以上縮小出来ません", vbExclamation)
            Exit Sub
        End If
    Else
        '非表示の列がある時
        If objSelection.Address <> objVisible.Address Then
            If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
                ActiveWindow.View = lngWindowView
                Application.ScreenUpdating = True
                Select Case MsgBox("非表示の列を対象としますか？", vbYesNoCancel + vbQuestion + vbDefaultButton2)
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
            
            '標準プレビューに再度変更する
            Application.ScreenUpdating = False
            lngWindowView = ActiveWindow.View
        End If
    End If
    
    '***********************************************
    '同じ幅の塊ごとにAddressを取得する
    '***********************************************
    Dim colAddress  As New Collection
    If objVisible Is Nothing Then
        Call colAddress.Add(objSelection.Address)
    Else
        Set colAddress = GetSameWidthAddresses(objSelection)
    End If
    
    '***********************************************
    '変更後のサイズのチェック
    '***********************************************
    Dim lngPixel    As Long    '幅(単位:Pixel)
    For i = 1 To colAddress.Count
        lngPixel = Range(colAddress(i)).Columns(1).Width / DPIRatio + lngSize
        If lngPixel < 0 Then
            ActiveWindow.View = lngWindowView
            Application.ScreenUpdating = True
            Call MsgBox("これ以上縮小出来ません", vbExclamation)
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
    Call SaveUndoInfo(E_ColSize2, Range(strSelection), colAddress)
    
    '同じ幅の塊ごとに幅を設定する
    For i = 1 To colAddress.Count
        lngPixel = Range(colAddress(i)).Columns(1).Width / DPIRatio + lngSize
        Range(colAddress(i)).ColumnWidth = PixelToWidth(lngPixel)
    Next i
    
    '改ページを元に戻す
    If blnDisplayPageBreaks = True Then
        ActiveSheet.DisplayAutomaticPageBreaks = True
    End If
    Call SetOnUndo
'    Call SetOnRepeat

    ActiveWindow.View = lngWindowView
Exit Sub
ErrHandle:
    If blnDisplayPageBreaks = True Then
        ActiveSheet.DisplayAutomaticPageBreaks = True
    End If
    ActiveWindow.View = lngWindowView
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　GetSameWidthAddresses
'[ 概  要 ]　同じ幅の塊ごとのアドレスを配列で取得する
'[ 引  数 ]　選択された領域
'[ 戻り値 ]　アドレスの配列
'*****************************************************************************
Public Function GetSameWidthAddresses(ByRef objSelection As Range) As Collection
    Dim i           As Long
    Dim objRange    As Range
    Dim lngLastCell As Long
    Dim objColumns  As Range
    
    Set GetSameWidthAddresses = New Collection
    
    '選択範囲のColumnsの和集合を取り重複列を排除する
    Set objColumns = Union(objSelection.EntireColumn, objSelection.EntireColumn)
    
    'エリアの数だけループ
    For Each objRange In objColumns.Areas
        i = objRange.Column
        lngLastCell = i + objRange.Columns.Count - 1
        
        '同じ幅の塊ごとに列幅を判定する
        While i <= lngLastCell
            '同じ幅の列のアドレスを保存
            Call GetSameWidthAddresses.Add(GetSameWidthAddress(i, lngLastCell))
        Wend
    Next objRange
End Function

'*****************************************************************************
'[ 関数名 ]　GetSameWidthAddress
'[ 概  要 ]　連続する列でlngColと同じ幅の列を表わすアドレスを取得する
'[ 引  数 ]　lngCol:最初の列(実行後は次の列)、lngLastCell:検索の最後の列
'[ 戻り値 ]　なし
'*****************************************************************************
Private Function GetSameWidthAddress(ByRef lngCol As Long, ByVal lngLastCell As Long) As String
    Dim lngPixel As Long
    Dim i As Long
    lngPixel = Columns(lngCol).Width / DPIRatio
    
    For i = lngCol + 1 To lngLastCell
        If (Columns(i).Width / DPIRatio) <> lngPixel Then
            Exit For
        End If
    Next i
    GetSameWidthAddress = Range(Columns(lngCol), Columns(i - 1)).Address
    lngCol = i
End Function

'*****************************************************************************
'[ 関数名 ]　ChangeShapeWidth
'[ 概  要 ]　図形のサイズ変更
'[ 引  数 ]　lngSize:変更サイズ
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ChangeShapeWidth(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim objGroups      As ShapeRange
    Dim blnFitGrid     As Boolean
    
    'アンドゥ用に元のサイズを保存する
    Application.ScreenUpdating = False
    Call SaveUndoInfo(E_ShapeSize2, Selection.ShapeRange)
    
    '回転している図形をグループ化する
    Set objGroups = GroupSelection(Selection.ShapeRange)
    
    '[Shift]Keyが押下されていれば、枠線に合わせて変更する
    If GetKeyState(vbKeyShift) < 0 Then
        blnFitGrid = True
    End If
    
    '図形のサイズを変更
    Call ChangeShapesWidth(objGroups, lngSize, blnFitGrid)
    
    '回転している図形のグループ化を解除し元の図形を選択する
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
'    Call SetOnRepeat
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　ChangeShapesWidth
'[ 概  要 ]　図形のサイズ変更
'[ 引  数 ]　objShapes:図形
'            lngSize:変更サイズ(Pixel)
'            blnFitGrid:枠線にあわせるか
'            blnTopLeft:左または上方向に変化させる
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ChangeShapesWidth(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean, Optional ByVal blnTopLeft As Boolean = False)
    Dim objShape     As Shape
    Dim lngLeft      As Long
    Dim lngRight     As Long
    Dim lngOldWidth  As Long
    Dim lngNewWidth  As Long
    Dim lngNewLeft   As Long
    Dim lngNewRight  As Long
    
    '図形の数だけループ
    For Each objShape In objShapes
        lngOldWidth = Round(objShape.Width / DPIRatio)
        lngLeft = Round(objShape.Left / DPIRatio)
        lngRight = Round((objShape.Left + objShape.Width) / DPIRatio)
        
        '枠線にあわせるか
        If blnFitGrid = True Then
            If blnTopLeft = True Then
                If lngSize > 0 Then
                    lngNewLeft = GetLeftGrid(lngLeft, objShape.TopLeftCell.EntireColumn)
                Else
                    lngNewLeft = GetRightGrid(lngLeft, objShape.TopLeftCell.EntireColumn)
                End If
                lngNewWidth = lngRight - lngNewLeft
            Else
                If lngSize < 0 Then
                    lngNewRight = GetLeftGrid(lngRight, objShape.BottomRightCell.EntireColumn)
                Else
                    lngNewRight = GetRightGrid(lngRight, objShape.BottomRightCell.EntireColumn)
                End If
                lngNewWidth = lngNewRight - lngLeft
            End If
            If lngNewWidth < 0 Then
                lngNewWidth = 0
            End If
        Else
            'ピクセル単位の変更をする
            If lngOldWidth + lngSize >= 0 Then
                If blnTopLeft = True And lngLeft = 0 And lngSize > 0 Then
                    lngNewWidth = lngOldWidth
                Else
                    lngNewWidth = lngOldWidth + lngSize
                End If
            Else
                lngNewWidth = lngOldWidth
            End If
        End If
    
        If lngSize > 0 And blnTopLeft = True Then
            objShape.Left = (lngRight - lngNewWidth) * DPIRatio
        End If
        objShape.Width = lngNewWidth * DPIRatio
        
        'Excel2007のバグ対応
        If Round(objShape.Width / DPIRatio) <> lngNewWidth Then
            objShape.Width = (lngNewWidth + lngSize) * DPIRatio
        End If
        
        If Round(objShape.Width / DPIRatio) <> lngOldWidth Then
            If blnTopLeft = True Then
                objShape.Left = (lngRight - lngNewWidth) * DPIRatio
            Else
                objShape.Left = lngLeft * DPIRatio
            End If
        End If
    Next objShape
End Sub

'*****************************************************************************
'[ 関数名 ]　GetLeftGrid
'[ 概  要 ]　入力の位置の左横の枠線の位置を取得(単位ピクセル)
'[ 引  数 ]　lngPos:位置(単位ピクセル)
'            objColumn: lngPosを含む列
'[ 戻り値 ]　図形の左側の枠線の位置(単位ピクセル)
'*****************************************************************************
Public Function GetLeftGrid(ByVal lngPos As Long, ByRef objColumn As Range) As Long
    Dim i       As Long
    Dim lngLeft As Long
    
    If lngPos <= Round(Columns(2).Left / DPIRatio) Then
        GetLeftGrid = 0
        Exit Function
    End If
    
    For i = objColumn.Column To 1 Step -1
        lngLeft = Round(GetWidth(Range(Columns(1), Columns(i - 1))) / DPIRatio)
        If lngLeft < lngPos Then
            GetLeftGrid = lngLeft
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[ 関数名 ]　GetRightGrid
'[ 概  要 ]　入力の位置の右横の枠線の位置を取得(単位ピクセル)
'[ 引  数 ]　lngPos:位置(単位ピクセル)
'            objColumn: lngPosを含む列
'[ 戻り値 ]　図形の右側の枠線の位置(単位ピクセル)
'*****************************************************************************
Public Function GetRightGrid(ByVal lngPos As Long, ByRef objColumn As Range) As Long
    Dim i        As Long
    Dim lngRight As Long
    
    If lngPos >= Round(GetWidth(Range(Columns(1), Columns(Columns.Count - 1))) / DPIRatio) Then
        GetRightGrid = Round(GetWidth(Columns) / DPIRatio)
        Exit Function
    End If
    
    For i = objColumn.Column + 1 To Columns.Count
        lngRight = Round(GetWidth(Range(Columns(1), Columns(i - 1))) / DPIRatio)
        If lngRight > lngPos Then
            GetRightGrid = lngRight
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
    
    '行の非表示は選択する
    Set GetVisibleCells = Union(objCells.EntireColumn, objCells.EntireColumn)
    Set GetVisibleCells = IntersectRange(GetVisibleCells, objRange)
Exit Function
ErrHandle:
    Set GetVisibleCells = Nothing
End Function

'*****************************************************************************
'[ 関数名 ]　MoveBorder
'[ 概  要 ]　列の境界線を左右に移動する
'[ 引  数 ]　lngSize : 移動サイズ(単位:ピクセル)
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveBorder(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim strSelection      As String
    Dim objRange          As Range
    Dim lngPixel(1 To 2)  As Long  '先頭列と最終列のサイズ
    Dim k                 As Long  '最終列の列番号
    
    strSelection = Selection.Address
    Set objRange = Selection

    '選択エリアが複数なら対象外
    If objRange.Areas.Count <> 1 Then
        Call MsgBox("このコマンドは複数の選択範囲に対して実行できません。", vbExclamation)
        Exit Sub
    End If

    '選択列が１列なら対象外
    If objRange.Columns.Count = 1 Then
        Exit Sub
    End If
    
    '最終列の列番号
    k = objRange.Columns.Count
    
    '変更後のサイズ
    lngPixel(1) = objRange.Columns(1).Width / DPIRatio + lngSize '先頭列
    lngPixel(2) = objRange.Columns(k).Width / DPIRatio - lngSize '最終列
    
    'サイズのチェック
    If lngPixel(1) < 0 Or lngPixel(2) < 0 Then
        Exit Sub
    End If
    
    '***********************************************
    'サイズの変更
    '***********************************************
    Application.ScreenUpdating = False
    'アンドゥ用に元のサイズを保存する
    Dim colAddress  As New Collection
    Call colAddress.Add(objRange.Columns(1).Address)
    Call colAddress.Add(objRange.Columns(k).Address)
    Call SaveUndoInfo(E_ColSize2, Range(strSelection), colAddress)
    
    'サイズの変更
    objRange.Columns(1).ColumnWidth = PixelToWidth(lngPixel(1))
    objRange.Columns(k).ColumnWidth = PixelToWidth(lngPixel(2))
    Call SetOnUndo
'    Call SetOnRepeat
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

''*****************************************************************************
''[ 関数名 ]　MoveShape
''[ 概  要 ]　図形を左右に移動する
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
'    Call SaveUndoInfo(E_ShapeSize2, Selection.ShapeRange)
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
'        Call MoveShapesLR(objGroups, lngSize, blnFitGrid)
'
'        '回転している図形のグループ化を解除し元の図形を選択する
'        Call UnGroupSelection(objGroups).Select
'    Else
'        '図形を左右に移動する
'        Call MoveShapesLR(Selection.ShapeRange, lngSize, blnFitGrid)
'    End If
'
'    Call SetOnUndo
'Exit Sub
'ErrHandle:
'    Call MsgBox(Err.Description, vbExclamation)
'End Sub
'
'*****************************************************************************
'[ 関数名 ]　DistributeColsWidth
'[ 概  要 ]　選択された列の幅を揃える
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub DistributeColsWidth()
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objRange     As Range
    Dim lngColCount  As Long    '選択された列の数
    Dim dblWidth     As Double  '選択された列の幅の合計
    Dim objSelection As Range   '選択されたすべての列のコレクション
    Dim strSelection As String
    Dim objVisible   As Range   '可視Range
    
    '選択されているオブジェクトを判定
    Select Case CheckSelection()
    Case E_Other
        Exit Sub
    Case E_Shape
        Call DistributeShapeWidth
        Exit Sub
    End Select
    
    '選択範囲のColumnsの和集合を取り重複列を排除する
    strSelection = Selection.Address
    Set objSelection = Union(Selection.EntireColumn, Selection.EntireColumn)
    
    '選択範囲の可視部分を取出す
    Set objVisible = GetVisibleCells(objSelection)
    
    'すべて非表示の時
    If objVisible Is Nothing Then
        Exit Sub
    End If
    
    '非表示の列がある時
    If objSelection.Address <> objVisible.Address Then
        If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
            Select Case MsgBox("非表示の列を対象としますか？", vbYesNoCancel + vbQuestion + vbDefaultButton2)
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
        dblWidth = dblWidth + GetWidth(objRange)
        lngColCount = lngColCount + objRange.Columns.Count
    Next objRange
    
    If lngColCount = 1 Then
        Exit Sub
    End If
    
    '***********************************************
    'サイズの変更
    '***********************************************
    Application.ScreenUpdating = False
    
    'アンドゥ用に元のサイズを保存する
    Call SaveUndoInfo(E_ColSize, Selection, GetSameWidthAddresses(objSelection))
    objSelection.ColumnWidth = PixelToWidth(dblWidth / DPIRatio / lngColCount)
    Call SetOnUndo
'    Call SetOnRepeat
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　DistributeShapeWidth
'[ 概  要 ]　選択された図形の幅を揃える
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub DistributeShapeWidth()
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
    
    Call DistributeShapesWidth(objGroups)
    
    '回転している図形のグループ化を解除し元の図形を選択する
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
'    Call SetOnRepeat
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　DistributeShapesWidth
'[ 概  要 ]　図形の幅を揃える
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub DistributeShapesWidth(ByRef objShapeRange As ShapeRange)
    Dim objShape   As Shape
    Dim dblWidth   As Double
    
    '図形の数だけループ
    For Each objShape In objShapeRange
        dblWidth = dblWidth + objShape.Width
    Next objShape
    With objShapeRange
        .Width = Round(dblWidth / .Count / DPIRatio) * DPIRatio
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　GetWidth
'[ 概  要 ]　選択エリアの幅を取得
'            Width/Leftプロパティは32767以上の幅を計算出来ないため
'[ 引  数 ]　幅を取得するエリア
'[ 戻り値 ]　幅(Widthプロパティ)
'*****************************************************************************
Private Function GetWidth(ByRef objRange As Range) As Double
'    With objRange
'        GetWidth = .Columns(.Columns.Count).Left - .Left + .Columns(.Columns.Count).Width
'    End With
    
    Dim lngCount   As Long
    Dim lngHalf    As Long
    Dim MaxWidth   As Double '幅の最大値

    MaxWidth = 32767 * DPIRatio
    If objRange.Width < MaxWidth Then
        GetWidth = objRange.Width
    Else
        With objRange
            '前半＋後半の幅を合計
            lngCount = .Columns.Count
            lngHalf = lngCount / 2
            GetWidth = GetWidth(Range(.Columns(1), .Columns(lngHalf))) + _
                       GetWidth(Range(.Columns(lngHalf + 1), .Columns(lngCount)))
        End With
    End If
End Function

'*****************************************************************************
'[ 関数名 ]　MergeCellsAsColumn
'[ 概  要 ]　横方向に結合(複数列に値がある時は空白で連結)
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MergeCellsAsColumn()
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
        '行の数だけループ
        For i = 1 To objRange.Rows.Count
            
            '複数のセルに値がある時連結する
            Set objMergeCell = objRange.Rows(i)
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
'    Call SetOnRepeat
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　SplitColumn
'[ 概  要 ]　列を分割する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SplitColumn()
On Error GoTo ErrHandle
    Dim i               As Long
    Dim objRange        As Range
    Dim lngPixel        As Double  '１列の幅
    Dim lngSplitCount   As Long    '分割数
    Dim blnCheckInsert  As Boolean
    Dim objNewCol       As Range   '新しい列
    
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
    
    '複数列選択なら対象外
    If objRange.Columns.Count <> 1 Then
        Call MsgBox("このコマンドは複数の選択列に対して実行できません。", vbExclamation)
        Exit Sub
    End If
    
    '元の幅
    lngPixel = objRange.EntireColumn.Width / DPIRatio
    
    '****************************************
    '分割数を選択させる
    '****************************************
    With frmSplitCount
        Call .SetChkLabel(True)
        
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
    Call SaveUndoInfo(E_SplitCol, objRange, lngSplitCount)
    If blnCheckInsert = False Then
        Call SetPlacement
    End If
    
    '選択列の右側に１列挿入
    Call objRange(1, 2).EntireColumn.Insert
    
    '新しい列
    Set objNewCol = objRange(1, 2).EntireColumn
    
    '*************************************************
    '罫線を整える
    '*************************************************
    '挿入列の１セル毎に罫線をコピーする
    If blnCheckInsert = True Then
        Call CopyBorder("左右上下", objRange.EntireColumn, objNewCol)
    Else
        Call CopyBorder("右上下", objRange.EntireColumn, objNewCol)
    End If
    
    '*************************************************
    '横方向に結合する
    '*************************************************
    If blnCheckInsert = False Then
        Call MergeCols(2, objRange.EntireColumn, objNewCol)
    Else
        Call MergeCols(3, objRange.EntireColumn, objNewCol)
    End If
    
    '*************************************************
    '分割を繰返す
    '*************************************************
    '分割数だけ、列を挿入する
    For i = 3 To lngSplitCount
        Call objNewCol.EntireColumn.Insert
    Next i
    
    '*************************************************
    '各列を縦方向に結合する
    '*************************************************
    If blnCheckInsert = True Then
        For i = 2 To lngSplitCount
            Call MergeCols(4, objRange.EntireColumn, objRange(1, i).EntireColumn)
        Next i
    End If

    '*************************************************
    '幅の整備
    '*************************************************
    If blnCheckInsert = False Then
        If lngSplitCount = 2 Then
            objRange.EntireColumn.ColumnWidth = PixelToWidth(Round(lngPixel / 2))
            objNewCol.EntireColumn.ColumnWidth = PixelToWidth(Int(lngPixel / 2))
        Else
            With Range(objRange, objNewCol).EntireColumn
                If lngPixel < lngSplitCount Then
                    .ColumnWidth = PixelToWidth(1)
                Else
                    .ColumnWidth = PixelToWidth(lngPixel / lngSplitCount)
                End If
            End With
        End If
    End If
    
    '*************************************************
    '境界線を消す
    '*************************************************
    If blnCheckInsert = False Then
        With Range(objRange, objNewCol).EntireColumn
            .Borders(xlInsideVertical).LineStyle = xlNone
        End With
    End If
    
    '*************************************************
    '後処理
    '*************************************************
    Call Range(objRange, objRange(1, lngSplitCount)).Select
    If blnCheckInsert = False Then
        Call ResetPlacement
    End If
    Call SetOnUndo
    
    If blnDisplayPageBreaks = True Then
        ActiveSheet.DisplayAutomaticPageBreaks = True
    End If
    Application.ScreenUpdating = True
    Call Application.OnRepeat("", "")
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
'[ 関数名 ]　EraseColumn
'[ 概  要 ]　選択された列を消去する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub EraseColumn()
On Error GoTo ErrHandle
    Dim objSelection      As Range
    Dim strSelection      As String
    Dim objRange          As Range
    Dim lngLeftCol        As Long  '左隣の列番号
    Dim lngRightCol       As Long  '右隣の列番号
    Dim objRow            As Range '終了時に選択させる行
    
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = Selection
    strSelection = objSelection.Address
    Set objRange = objSelection.EntireColumn
    
    '終了時に選択させる行
    Set objRow = objSelection.EntireRow
    
    '選択エリアが複数なら対象外
    If objSelection.Areas.Count <> 1 Then
        Call MsgBox("このコマンドは複数の選択範囲に対して実行できません。", vbExclamation)
        Exit Sub
    End If
    
    If objSelection.Columns.Count = Columns.Count Then
        Call MsgBox("このコマンドはすべての列の選択に対して実行できません。", vbExclamation)
        Exit Sub
    End If
    
    '左隣の列番号
    lngLeftCol = objRange.Column - 1
    '右隣の列番号
    lngRightCol = objRange.Column + objRange.Columns.Count
    
    '****************************************
    '消去のパターンを選択させる
    '****************************************
    Dim enmSelectType As ESelectType  '消去パターン
    Dim blnHidden   As Boolean        '非表示とするかどうか
    With frmEraseSelect
        'シートの１列目が選択されているか
        .TopSelect = (lngLeftCol = 0)
        
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
    '値のある列を削除するかどうか判定
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
    Set colAddress = GetSameWidthAddresses(Range(strSelection))
    
    '****************************************
    '選択列の左右の列を待避
    '****************************************
    Dim objCol(0 To 1)   As Range  '列の幅を変える列
    Select Case enmSelectType
    Case E_Front
        Set objCol(0) = Columns(lngRightCol)
        Call colAddress.Add(objCol(0).Address)
    Case E_Back
        Set objCol(0) = Columns(lngLeftCol)
        Call colAddress.Add(objCol(0).Address)
    Case E_Middle
        Set objCol(0) = Columns(lngRightCol)
        Set objCol(1) = Columns(lngLeftCol)
        Call colAddress.Add(objCol(0).Address)
        Call colAddress.Add(objCol(1).Address)
    End Select
    
    '****************************************
    '列を消去
    '****************************************
    Dim lngPixel  As Long   '消去される列の幅(単位:ピクセル)
    Application.ScreenUpdating = False
    
    'アンドゥ用に元の状態を保存する
    If blnHidden = True Then
        Call SaveUndoInfo(E_ColSize, Range(strSelection), colAddress)
    Else
        Call SaveUndoInfo(E_EraseCol, Range(strSelection), colAddress)
    End If
    
    '図形は移動させない
    Call SetPlacement
    
    '消去される列の幅を保存
    lngPixel = objRange.Width / DPIRatio
    
    If blnHidden = True Then
        '非表示
        objRange.Hidden = True
    Else
        '右端の罫線をコピーする
        With objRange
            If .Column > 1 Then
                Call CopyBorder("右", .Columns(.Columns.Count), .Columns(0))
            End If
        End With
        
        '削除
        Call objRange.Delete(xlShiftToLeft)
    End If
    
    '****************************************
    '列幅を拡大
    '****************************************
    Dim lngWkPixel  As Long
    Select Case enmSelectType
    Case E_Front, E_Back
        objCol(0).ColumnWidth = WorksheetFunction.Min(MaxColumnWidth, PixelToWidth(objCol(0).Width / DPIRatio + lngPixel))
    Case E_Middle
        objCol(0).ColumnWidth = WorksheetFunction.Min(MaxColumnWidth, PixelToWidth(objCol(0).Width / DPIRatio + Int(lngPixel / 2 + 0.5)))
        objCol(1).ColumnWidth = WorksheetFunction.Min(MaxColumnWidth, PixelToWidth(objCol(1).Width / DPIRatio + Int(lngPixel / 2)))
    End Select
    
    '****************************************
    'セルを選択
    '****************************************
    Select Case enmSelectType
    Case E_Front, E_Back
        IntersectRange(objRow, objCol(0)).Select
    Case E_Middle
        IntersectRange(objRow, UnionRange(objCol(0), objCol(1))).Select
    End Select
    Call ResetPlacement
    Call SetOnUndo
    Call Application.OnRepeat("", "")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
    Call ResetPlacement
End Sub

'*****************************************************************************
'[ 関数名 ]　CopyBorder
'[ 概  要 ]　罫線をコピーする
'[ 引  数 ]　罫線のタイプ(複数指定可):上下左右
'            objFromCol：コピ－元の列、objToCol：コピー先の列
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub CopyBorder(ByVal strBorderType As String, ByRef objFromCol As Range, ByRef objToCol As Range)
    Dim i As Long
    Dim j As Long
    Dim udtBorder(0 To 3) As TBorder '罫線の種類(上下左右)
    Dim lngLast    As Long
    
    Call ActiveSheet.UsedRange '最後のセルを修正する Undo出来なくなります
    If objFromCol.Rows.Count = Rows.Count Then
        '最後のセルまで整備すれば終了する
        lngLast = Cells.SpecialCells(xlCellTypeLastCell).Row
        If lngLast > MAXROWCOLCNT Then
            lngLast = MAXROWCOLCNT
        End If
    Else
        '選択されたすべての行を整備する
        lngLast = objFromCol.Rows.Count
    End If
    
    '1行毎にループ
    For i = 1 To lngLast
        '罫線の種類を保存
        With objFromCol.Rows(i)
            If InStr(1, strBorderType, "上") <> 0 Then
                udtBorder(0) = GetBorder(.Borders(xlEdgeTop))
            End If
            If InStr(1, strBorderType, "下") <> 0 Then
                udtBorder(1) = GetBorder(.Borders(xlEdgeBottom))
            End If
            If InStr(1, strBorderType, "左") <> 0 Then
                udtBorder(2) = GetBorder(.Borders(xlEdgeLeft))
            End If
            If InStr(1, strBorderType, "右") <> 0 Then
                udtBorder(3) = GetBorder(.Borders(xlEdgeRight))
            End If
        End With
        
        '罫線を書く
        With objToCol.Rows(i)
            If InStr(1, strBorderType, "上") <> 0 Then
                Call SetBorder(udtBorder(0), .Borders(xlEdgeTop))
            End If
            If InStr(1, strBorderType, "下") <> 0 Then
                Call SetBorder(udtBorder(1), .Borders(xlEdgeBottom))
            End If
            If InStr(1, strBorderType, "左") <> 0 Then
                Call SetBorder(udtBorder(2), .Borders(xlEdgeLeft))
            End If
            If InStr(1, strBorderType, "右") <> 0 Then
                Call SetBorder(udtBorder(3), .Borders(xlEdgeRight))
            End If
        End With

        'ステータスバーに進捗状況を表示
        If i / lngLast * 12 <> j Then
            j = i / lngLast * 12
            Application.StatusBar = String(j, "■") & String(12 - j, "□")
        End If
    Next i
    Application.StatusBar = False
End Sub

'*****************************************************************************
'[ 関数名 ]　MergeCols
'[ 概  要 ]　先頭列から右端の列まで横方向に結合する
'[ 引  数 ]　lngType:結合のタイプ、
'            objTopRow：結合の先頭列、objBottomRow：結合の右端列
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MergeCols(ByVal lngtype As Long, ByRef objTopCol As Range, ByRef objRightCol As Range)
    Dim i          As Long
    Dim lngLast    As Long
    Dim objRange As Range

    Call ActiveSheet.UsedRange '最後のセルを修正する Undo出来なくなります
    If objTopCol.Rows.Count = Rows.Count Then
        '最後のセルまで整備すれば終了する
        lngLast = Cells.SpecialCells(xlCellTypeLastCell).Row
        If lngLast > MAXROWCOLCNT Then
            lngLast = MAXROWCOLCNT
        End If
    Else
        '選択されたすべての行を整備する
        lngLast = objTopCol.Rows.Count
    End If
    
    '1行毎にループ
    For i = 1 To lngLast
        With objRightCol.Cells(i, 1)
            '右端の列のセルが結合セルか
            If .MergeArea.Count = 1 Then
                Set objRange = GetMergeColRange(lngtype, objTopCol.Cells(i, 1), .Cells)
                If Not (objRange Is Nothing) Then
                    Call objRange.Merge
                End If
            End If
        End With
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　GetMergeColRange
'[ 概  要 ]　横方向に結合する領域を取得する
'[ 引  数 ]　lngType:結合のタイプ、
'            objBaseCell:先頭列のセル、objTergetCell:対象の列のセル
'[ 戻り値 ]　結合する領域(Nothing:結合しない時)
'*****************************************************************************
Public Function GetMergeColRange(ByVal lngtype As Long, _
                                  ByRef objBaseCell As Range, _
                                  ByRef objTergetCell As Range) As Range
    Select Case lngtype
    Case 1 '先頭列から最終の列まで横方向に結合する
    Case 2 '先頭列が結合セルの時、先頭列から最終の列まで横方向に結合する
        If objBaseCell.MergeArea.Count = 1 Then
            Exit Function
        End If
    Case 3 '先頭列が横方向に結合セルの時、先頭列から最終の列まで横方向に結合する
        If objBaseCell.MergeArea.Columns.Count = 1 Then
            Exit Function
        End If
    Case 4 '先頭列が縦方向のみ結合セルの時、対象の列のセルを縦方向に結合する
        If objBaseCell.MergeArea.Rows.Count = 1 Or _
           objBaseCell.MergeArea.Columns.Count > 1 Then
            Exit Function
        End If
    End Select
    
    Select Case lngtype
    Case 1, 2, 3
        Set GetMergeColRange = Range(objBaseCell.MergeArea, objTergetCell)
    Case 4
        Set GetMergeColRange = objTergetCell.Resize(objBaseCell.MergeArea.Rows.Count, 1)
    End Select
End Function

'*****************************************************************************
'[ 関数名 ]　ShowWidth
'[ 概  要 ]　エリア毎に選択された列の幅を一覧で表示する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ShowWidth()
On Error GoTo ErrHandle
    'Rangeオブジェクトが選択されているか判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Call frmSizeList.Initialize(E_Col)
    Call frmSizeList.Show
    Call Application.OnRepeat("", "")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　MoveColumnsBorder
'[ 概  要 ]　列の境界のセルを左右に移動する
'[ 引  数 ]　-1:左に移動、1:右に移動
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveColumnsBorder(ByVal lngLeftRight As Long)
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objSelection As Range
    Dim objWkRange   As Range
    Dim lngColCount  As Long  '選択領域の列数
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
    lngColCount = objSelection.Columns.Count
    Dim objChkRange(0 To 2) As Range

    If objSelection.Columns.Count = 1 Then
        Exit Sub
    End If

    '左に移動する時
    If lngLeftRight < 0 Then
        Set objChkRange(1) = ArrangeRange(Range(objSelection.Columns(1), objSelection.Columns(2)))
        Set objChkRange(2) = ArrangeRange(objSelection.Columns(2))
    Else '右に移動する時
        Set objChkRange(1) = ArrangeRange(Range(objSelection.Columns(lngColCount - 1), objSelection.Columns(lngColCount)))
        Set objChkRange(2) = ArrangeRange(objSelection.Columns(lngColCount - 1))
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
    Call SaveUndoInfo(E_CellBorder, objSelection)
    
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
    If lngLeftRight < 0 Then
        '左に移動する
        Call CopyBorder("右", objWkRange.Columns(2), objWkRange.Columns(1))
        Call objWkRange.Columns(2).Delete(xlToLeft)
        Call CopyBorder("右上下", objWkRange.Columns(lngColCount - 1), objWkRange.Columns(lngColCount))
        Call MergeCols(1, objWkRange.Columns(lngColCount - 1), objWkRange.Columns(lngColCount))
    Else
        '右に移動する
        Call CopyBorder("右", objWkRange.Columns(lngColCount), objWkRange.Columns(lngColCount - 1))
        Call objWkRange.Columns(lngColCount).Delete(xlToLeft)
        Call objWkRange.Columns(2).Insert(xlToRight)
        Call CopyBorder("右上下", objWkRange.Columns(1), objWkRange.Columns(2))
        Call MergeCols(1, objWkRange.Columns(1), objWkRange.Columns(2))
    End If
    
    Call objWkRange.Worksheet.Range(objSelection.Address).Copy(objSelection)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    Call SetOnUndo
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
'    Call SetOnRepeat
Exit Sub
ErrHandle:
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
    Call MsgBox(Err.Description, vbExclamation)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
End Sub

'*****************************************************************************
'[ 関数名 ]　AutoColFit
'[ 概  要 ]　列の幅を文字列の長さにあわせる
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub AutoColFit()
On Error GoTo ErrHandle
    Dim objSelection As Range
    Dim objWorkRange As Range
    Dim colColumns   As Collection '中身は列のRange
    Dim i As Long
    Dim strErrMsg As String
    
    '選択されているオブジェクトを判定
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = IntersectRange(Selection, GetUsedRange())
    If objSelection Is Nothing Then
        Exit Sub
    End If
    
    '列方向の結合セルを含む時
    If IsBorderMerged(objSelection) Then
        Call MsgBox("結合されたセルの一部を選択することはできません。", vbExclamation)
        Exit Sub
    End If
    
    If WorksheetFunction.CountA(objSelection) = 0 Then
        Exit Sub
    End If
    
    '非表示の列を対象とするか確認？
    Dim objVisible    As Range
    Dim objNonVisible As Range
    Set objVisible = GetVisibleCells(objSelection)
    Set objNonVisible = MinusRange(objSelection, objVisible)
    If Not (objNonVisible Is Nothing) Then
        If WorksheetFunction.CountA(objNonVisible) > 0 Then
            If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
                Select Case MsgBox("非表示のセルを対象としますか？", vbYesNoCancel + vbQuestion + vbDefaultButton2)
                Case vbNo
                    Set objSelection = objVisible
                Case vbCancel
                    Exit Sub
                End Select
            Else
                Set objSelection = objVisible
            End If
        End If
    End If
    
    If objSelection.MergeCells = False Then
        
        'アンドゥ用に元のサイズを保存する
        Call SaveUndoInfo(E_ColSize, objSelection, GetSameWidthAddresses(objSelection))
        
        Call objSelection.Columns.AutoFit
        
        Call SetOnUndo
'        Call SetOnRepeat
    Exit Sub
    End If

    '横方向の結合セルの、一番結合幅が広いセルを取得
    Set objWorkRange = GetDoRange(objSelection, colColumns)
    If objWorkRange Is Nothing Then
        Call MsgBox("横方向の結合が統一されていないため、実行できません", vbExclamation)
        Exit Sub
    End If
    
    If Not (MinusRange(objSelection, objWorkRange) Is Nothing) Then
        Call objWorkRange.Select
        If MsgBox("現在選択しているセルを対象として実行します。" & vbLf & "よろしいですか？", vbOKCancel + vbQuestion) = vbCancel Then
            Exit Sub
        End If
    End If
    
    If WorksheetFunction.CountA(objWorkRange) = 0 Then
        Exit Sub
    End If
    
    '***********************************************
    '実行
    '***********************************************
    Application.ScreenUpdating = False
    
    'アンドゥ用に元のサイズを保存する
    Call SaveUndoInfo(E_ColSize, objSelection, GetSameWidthAddresses(objWorkRange))
    
    '列毎にループ
    For i = 1 To colColumns.Count - 1
        Call SetColumnWidth(colColumns(i))
    Next
    
    '横方向の結合がないセルを一括設定
    If Not (colColumns(colColumns.Count) Is Nothing) Then
        Call colColumns(colColumns.Count).Columns.AutoFit
    End If
    
    Call SetOnUndo
'    Call SetOnRepeat
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ 関数名 ]　GetDoRange
'[ 概  要 ]　横方向の結合セルがある時、一番結合幅が広いセルのみ実行対象とする
'[ 引  数 ]　objSelection：選択されたセル
'            colColumns：実行対象を列毎に取得した、列毎のRangeの配列(戻り値)
'[ 戻り値 ]　実行対象のセル
'*****************************************************************************
Private Function GetDoRange(ByRef objSelection As Range, ByRef colColumns As Collection) As Range
    Dim i  As Long
    Dim objColumns    As Range
    Dim objSingleCols As Range
    Dim objArea       As Range
    Dim objLeftCol    As Range
    Dim objRightCol   As Range
    Dim objMeregCol   As Range
    Dim objWkRange(1 To 2) As Range
    
    '選択範囲のColumnsの和集合を取り重複列を排除する
    Set objColumns = Union(objSelection.EntireColumn, objSelection.EntireColumn)
    
    Set colColumns = New Collection
    
    For Each objArea In objColumns.Areas
        i = 1
        While i <= objArea.Columns.Count
            With GetMergeCol(objArea.Columns(i), objSelection)
                Set objLeftCol = .Columns(1)
                Set objRightCol = .Columns(.Columns.Count)
                i = i + .Columns.Count
            End With
            Set objWkRange(1) = ArrangeRange(IntersectRange(objLeftCol, objSelection))
            Set objWkRange(2) = ArrangeRange(IntersectRange(objRightCol, objSelection))
            
            Set objMeregCol = IntersectRange(objWkRange(1), objWkRange(2))
            If Not (objMeregCol Is Nothing) Then
                Set GetDoRange = UnionRange(GetDoRange, objMeregCol)
                If WorksheetFunction.CountA(objMeregCol) > 0 Then
                    If objMeregCol.Columns.Count = 1 Then
                        Set objSingleCols = UnionRange(objSingleCols, objMeregCol)
                    Else
                        Call colColumns.Add(objMeregCol)
                    End If
                End If
            End If
        Wend
    Next
    
    '一番最後には、横方向の結合がないセルを設定
    Call colColumns.Add(objSingleCols)
End Function

'*****************************************************************************
'[ 関数名 ]　GetMergeCol
'[ 概  要 ]　横方向の結合の最大幅の列を取得する
'[ 引  数 ]　調査する列、選択された列
'[ 戻り値 ]　最大幅の列
'*****************************************************************************
Private Function GetMergeCol(ByRef objColumn As Range, ByRef objSelection As Range) As Range
    Dim i          As Long
    Dim objRange   As Range
    Dim objWkRange As Range
    
    Set objWkRange = ArrangeRange(IntersectRange(objColumn, objSelection))
    
    '選択範囲のColumnsの和集合を取り重複列を排除する
    Set objRange = Union(objWkRange.EntireColumn, objWkRange.EntireColumn)

    '万一無限ループにならないようにループ回数に上限を与える
    For i = 1 To Columns.Count
        Set objWkRange = ArrangeRange(IntersectRange(objRange, objSelection))
        
        '選択範囲のColumnsの和集合を取り重複列を排除する
        Set GetMergeCol = Union(objWkRange.EntireColumn, objWkRange.EntireColumn)
            
        If GetMergeCol.Address = objRange.Address Then
            Exit Function
        End If
        Set objRange = GetMergeCol
    Next i
    
    '無限ループにおちいる時
    Call Err.Raise(C_CheckErrMsg, , "横方向の結合が統一されていないため、実行できません")
End Function

'*****************************************************************************
'[ 関数名 ]　SetColumnWidth
'[ 概  要 ]　列の幅を設定
'[ 引  数 ]　対象の列
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SetColumnWidth(ByRef objColumns As Range)
    Dim colAddress As Collection
    Dim i As Long
    Dim lngPixel    As Long
    Dim lngOldPixel As Long
    Dim lngNewPixel As Long
    
    lngOldPixel = objColumns.EntireColumn.Width / DPIRatio
    lngNewPixel = GetNewPixel(objColumns)
    
    '同じ幅のセルを取得する
    Set colAddress = GetSameWidthAddresses(objColumns)
    
    '同じ幅の塊ごとに幅を設定する
    For i = 1 To colAddress.Count
        With Range(colAddress(i))
            If lngOldPixel = 0 Then
                .ColumnWidth = PixelToWidth(lngNewPixel / .Columns.Count)
            Else
                lngPixel = .Width / DPIRatio * lngNewPixel / lngOldPixel
                .ColumnWidth = PixelToWidth(lngPixel / .Columns.Count)
            End If
        End With
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　GetNewPixel
'[ 概  要 ]　WorkSheetを利用し、文字列の幅を取得
'[ 引  数 ]　対象の列
'[ 戻り値 ]　新しい幅
'*****************************************************************************
Private Function GetNewPixel(ByRef objColumns As Range) As Long
On Error GoTo ErrHandle
    Dim objWorksheet As Worksheet
        
    Set objWorksheet = ThisWorkbook.Worksheets("Workarea1")
    Call DeleteSheet(objWorksheet)
    objWorksheet.Columns(1).ColumnWidth = PixelToWidth(objColumns.Width / DPIRatio)
    
    'Workarea1シートに対象セルをコピー
    Call objColumns.Copy(objWorksheet.Cells(1, 1))
            
    '先頭列の結合を解除
    Call objWorksheet.Columns(1).UnMerge
        
    'セル参照があると#ERRとなるケースがあるため値をコピーする
    If objWorksheet.Columns(1).HasFormula = False Then
    Else
        Call objColumns.Copy
        Call objWorksheet.Cells(1, 1).PasteSpecial(xlPasteValues)
        Application.CutCopyMode = False
    End If
    
    '先頭列の幅を設定
    Call objWorksheet.Columns(1).AutoFit
    
    GetNewPixel = objWorksheet.Columns(1).Width / DPIRatio
ErrHandle:
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
End Function

'*****************************************************************************
'[ 関数名 ]　PixelToWidth
'[ 概  要 ]　幅の単位を変換
'[ 引  数 ]　lngPixel : 幅(単位:ピクセル)
'[ 戻り値 ]　Width
'*****************************************************************************
Public Function PixelToWidth(ByVal lngPixel As Long) As Double
    'ピクセル情報を設定する
    Call SetPixelInfo 'Undoできなくなります
    
    If lngPixel <= x1 Then
        PixelToWidth = lngPixel / x1
    Else
        PixelToWidth = (lngPixel - x1) / (x2 - x1) + 1
    End If
    PixelToWidth = WorksheetFunction.RoundDown(PixelToWidth, 3)
End Function

'*****************************************************************************
'[ 関数名 ]　SetPixelInfo
'[ 概  要 ]　標準スタイルのフォントの1文字と2文字のピクセルを求める
'            x1：1文字のピクセル、x2：2文字のピクセル
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetPixelInfo()
On Error GoTo ErrHandle
    Static udtFont As TFont
    Dim objWorkbook As Workbook
        
    Set objWorkbook = ActiveWorkbook
    
    '標準スタイルのフォントが変更されたか判定
    With ActiveWorkbook.Styles("Normal").Font
        If udtFont.Name = .Name And udtFont.size = .size And _
           udtFont.Bold = .Bold And udtFont.Italic = .Italic Then
            Exit Sub
        Else
            'フォント情報を保存する
            udtFont.Name = .Name
            udtFont.size = .size
            udtFont.Bold = .Bold
            udtFont.Italic = .Italic
        End If
    End With
    
    'アクティブなブックしか標準スタイルのフォントを変更出来ないため
    Call ThisWorkbook.Activate
    
    'マクロのブックの標準スタイルのフォントを変更
    Dim blnScreenUpdating  As Boolean
    blnScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    With ThisWorkbook.Styles("Normal").Font
        If .Name <> udtFont.Name Then
            .Name = udtFont.Name
        End If
        If .size <> udtFont.size Then
            .size = udtFont.size
        End If
        If .Bold <> udtFont.Bold Then
            .Bold = udtFont.Bold
        End If
        If .Italic <> udtFont.Italic Then
            .Italic = udtFont.Italic
        End If
    End With
    Application.ScreenUpdating = blnScreenUpdating
    
    'サイズ情報を保存する
    With ThisWorkbook.Worksheets("Commands")
        .Range("O:O").ColumnWidth = 1
        .Range("P:P").ColumnWidth = 2
        x1 = .Range("O:O").Width / DPIRatio
        x2 = .Range("P:P").Width / DPIRatio
    End With
    
    Call objWorkbook.Activate
Exit Sub
ErrHandle:
    x1 = 13
    x2 = 21
    Call objWorkbook.Activate
End Sub

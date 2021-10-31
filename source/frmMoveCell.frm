VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMoveCell 
   Caption         =   "領域の操作"
   ClientHeight    =   2970
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5124
   OleObjectBlob   =   "frmMoveCell.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMoveCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Enum EModeType
    E_Move
    E_Copy
    E_Exchange
    E_CutInsert
End Enum

Private Enum EDirection
    E_ERR
    E_Non
    E_UP
    E_DOWN
    E_LEFT
    E_RIGHT
End Enum

Private Type TShape
    ID          As Long
    TopRow      As Long
    Top         As Double
    LeftColumn  As Long
    Left        As Double
    BottomRow   As Long
    Height      As Double
    RightColumn As Long
    Width       As Double
    Placement   As Byte
End Type

Private Type TRect
    Top      As Long
    Height   As Long
    Left     As Long
    Width    As Long
End Type

Private blnCheck As Boolean

Private enmModeType  As EModeType
Private strFromRange As String '元の領域
Private strToRange   As String '移動先
Private objFromSheet As Worksheet '元の領域
Private objTextbox   As Shape  '移動先を視覚的に表現する
Private lngDisplayObjects As Long
Private lngZoom      As Long

'*****************************************************************************
'[概要] 初期設定情報を設定
'[引数] enmType:作業タイプ
'       objFromRange:移動(コピー元)の領域
'       objToRange:選択中の領域
'[戻値] なし
'*****************************************************************************
Public Sub Initialize(ByVal enmType As EModeType, ByRef objFromRange As Range, ByRef objToRange As Range)
    
    lngDisplayObjects = ActiveWorkbook.DisplayDrawingObjects
    enmModeType = enmType
    Set objFromSheet = objFromRange.Worksheet
    
    strFromRange = objFromRange.Address
    strToRange = objToRange.Address
    If strFromRange <> strToRange Then
        strToRange = GetToRange(objFromRange, objToRange).Address
    End If
    
    '移動(コピー)先のシートをActivateにする
    Call objToRange.Worksheet.Activate
    
    '元と先のシートが同一の時、コピー元のセルを選択することで可視できるようにする
     Call Range(strFromRange).Select
'    If blnSameSheet = True Then
'        Call Range(strFromRange).Select
'    Else
'        Call Range(strToRange).Select
'    End If
    
    'テキストボックス作成
    ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
    With Range(strToRange)
        Set objTextbox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
    End With
    With objTextbox.TextFrame2.TextRange.Font
        .NameComplexScript = ActiveWorkbook.Styles("Normal").Font.Name
        .NameFarEast = ActiveWorkbook.Styles("Normal").Font.Name
        .Name = ActiveWorkbook.Styles("Normal").Font.Name
        .Size = ActiveWorkbook.Styles("Normal").Font.Size
    End With
    'テキストボックスの背景を変更
    With objTextbox.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.SchemeColor = 65
        .Transparency = 0.12  '背景を透けさせる
    End With
    
    'テキストボックスの罫線を変更
    With objTextbox.Line
        .Weight = 2#
        .Style = msoLineSingle
        .Transparency = 0#
        .Visible = msoTrue
        .ForeColor.SchemeColor = 64
        .BackColor.Rgb = Rgb(255, 255, 255)
        .Pattern = msoPattern50Percent
    End With
    
    'テキストボックスに選択セルのアドレスを表示
    With objTextbox.TextFrame.Characters
        .Text = Replace$(strToRange, "$", "")
        lblCellAddress.Caption = " " & .Text
    End With
    
    '選択領域が画面から消えている時
    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '画面分割のない時
        If IntersectRange(ActiveWindow.VisibleRange, objToRange) Is Nothing Then
            With objToRange
                Call ActiveWindow.ScrollIntoView(.Left / DPIRatio, .Top / DPIRatio, .Width / DPIRatio, .Height / DPIRatio)
            End With
        End If
    End If
    
'    'チェックボタンの制御
'    If blnSameSheet = False Then
'        chkExchange.Enabled = False
'        chkCopy.Enabled = False
'        chkExchange.Enabled = False
'        chkCutInsert.Enabled = False
'    End If
    
    'チェックボックスを設定
    Call ChangeMode
End Sub

'*****************************************************************************
'[ 関数名 ]　GetToRange
'[ 概  要 ]  選択中の領域から、移動先の領域の初期表示エリアを計算する
'[ 引  数 ]  objFromRange:移動(コピー元)の領域
'            objToRange:選択中の領域
'[ 戻り値 ]  移動先の領域の初期表示エリア
'*****************************************************************************
Private Function GetToRange(ByVal objFromRange As Range, ByVal objToRange As Range) As Range
    Select Case True
    Case objFromRange.Columns.Count = Columns.Count
        Set GetToRange = objToRange.EntireRow
    Case objFromRange.Rows.Count = Rows.Count
        Set GetToRange = objToRange.EntireColumn
    Case objToRange.Rows.Count = 1 And objToRange.Columns.Count = 1
        With Range(objFromRange.Address)
            Set GetToRange = objToRange.Range(Cells(1, 1), Cells(.Rows.Count, .Columns.Count))
        End With
    Case objToRange.Rows.Count = 1
        With Range(objFromRange.Address)
            Set GetToRange = objToRange.Range(Cells(1, 1), Cells(.Rows.Count, objToRange.Columns.Count))
        End With
    Case objToRange.Columns.Count = 1
        With Range(objFromRange.Address)
            Set GetToRange = objToRange.Range(Cells(1, 1), Cells(objToRange.Rows.Count, .Columns.Count))
        End With
    Case Else
        Set GetToRange = objToRange
    End Select
End Function

'*****************************************************************************
'[イベント]　UserForm_Initialize
'[ 概  要 ]　フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()
    '呼び元に通知する
    blnFormLoad = True
    lngZoom = ActiveWindow.Zoom
End Sub

'*****************************************************************************
'[イベント]　UserForm_Terminate
'[ 概  要 ]　フォームアンロード時
'*****************************************************************************
Private Sub UserForm_Terminate()
    '呼び元に通知する
    blnFormLoad = False
    
    If Not (objTextbox Is Nothing) Then
        Call objTextbox.Delete
    End If
    
    ActiveWorkbook.DisplayDrawingObjects = lngDisplayObjects
    ActiveWindow.Zoom = lngZoom
End Sub

'*****************************************************************************
'[イベント]　cmdOK_Click
'[ 概  要 ]　ＯＫボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrHandle
    Dim blnCopyObjectsWithCells  As Boolean
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells
    Application.CopyObjectsWithCells = False '呼び元で復元するため当モジュールでは復元しない
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False 'コメントがある時警告が出る時がある
    
    Call objTextbox.Delete
    Set objTextbox = Nothing
    
    'アンドゥ用に元の状態を保存する
    Select Case enmModeType
    Case E_Copy
         Call SaveUndoInfo(E_CopyCell, FromRange)
'        If blnSameSheet = True Then
'            Call SaveUndoInfo(E_CopyCell, FromRange)
'        Else
'            Call SaveUndoInfo(E_CopyCell, ToRange)
'        End If
    Case E_Move, E_Exchange, E_CutInsert
        Call SaveUndoInfo(E_MoveCell, FromRange)
    End Select
    
    'セルを移動・コピーする
    Select Case enmModeType
    Case E_Move
        Call MoveCell
    Case E_Copy
        Call CopyCell
    Case E_Exchange
        Call ExchangeCell
    Case E_CutInsert
        Call CutInsertCell
    End Select
    
    '図形を移動・コピーする
'    If blnCopyObjectsWithCells = True Then
        If chkOnlyValue.Value = False Then
            If ActiveSheet.Shapes.Count > 0 And lngDisplayObjects <> xlHide Then
                Select Case enmModeType
                Case E_Move, E_Copy
                    Call MoveShape
                Case E_Exchange
                    Call ExchangeShape
                Case E_CutInsert
                    Call CutInsertShape
                End Select
            End If
        End If
'    End If
        
    If enmModeType = E_CutInsert Then
        Select Case GetDirection()
        Case E_DOWN
            Call OffsetRange(ToRange, -FromRange.Rows.Count).Select
        Case E_RIGHT
            Call OffsetRange(ToRange, , -FromRange.Columns.Count).Select
        Case Else
            Call ToRange.Select
        End Select
    Else
        Call ToRange.Select
    End If
    
    Application.DisplayAlerts = True
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea2"))
    Call Unload(Me)
    Call SetOnUndo
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Call MsgBox(Err.Description, vbExclamation)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea2"))
    Call Unload(Me)
End Sub

'*****************************************************************************
'[ 関数名 ]　CopyCell
'[ 概  要 ]  領域をコピーする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub CopyCell()
    Dim objWkRange As Range
    
    '領域をワークシートにコピーする
    Set objWkRange = CopyToWkSheet(1, FromRange, ToRange)
    
    'ワークシートで領域のサイズを変更する
    Set objWkRange = ReSizeArea(objWkRange, ToRange.Rows.Count, ToRange.Columns.Count)
    
    If chkOnlyValue.Value = True Then
        '値のみコピー
        Call CopyOnlyValue(objWkRange, ToRange)
    Else
        Call objWkRange.Copy(ToRange)
    End If
End Sub
    
'*****************************************************************************
'[ 関数名 ]　MoveCell
'[ 概  要 ]  領域を移動する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveCell()
    Dim objWkRange As Range
    
    '領域をワークシートにコピーする
    Set objWkRange = CopyToWkSheet(1, FromRange, ToRange)
    
    'ワークシートで領域のサイズを変更する
    Set objWkRange = ReSizeArea(objWkRange, ToRange.Rows.Count, ToRange.Columns.Count)
    
    '元の領域をクリア
    With FromRange
        Call .Clear
        If .Worksheet.Cells(Rows.Count - 2, Columns.Count - 2).MergeCells = False Then
            'シート上の標準的な書式に設定
            Call .Worksheet.Cells(Rows.Count - 2, Columns.Count - 2).Copy(.Cells)
            Call .ClearContents
        End If
    End With
    
    Call objWkRange.Copy(ToRange)
End Sub

'*****************************************************************************
'[ 関数名 ]　ExchangeCell
'[ 概  要 ]  領域を入換える
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ExchangeCell()
    Dim objWkRange(1 To 2) As Range
    
    '｢WorkArea｣シートを使用して編集する
    Set objWkRange(1) = CopyToWkSheet(1, FromRange, ToRange)
    Set objWkRange(2) = CopyToWkSheet(2, ToRange, FromRange)
    
    'ワークシートで領域のサイズを変更する
    Set objWkRange(1) = ReSizeArea(objWkRange(1), ToRange.Rows.Count, ToRange.Columns.Count)
    Set objWkRange(2) = ReSizeArea(objWkRange(2), FromRange.Rows.Count, FromRange.Columns.Count)
    
    If chkOnlyValue.Value = True Then
        '値のみコピー
        Call CopyOnlyValue(objWkRange(1), ToRange)
        Call CopyOnlyValue(objWkRange(2), FromRange)
    Else
        Call objWkRange(1).Copy(ToRange)
        Call objWkRange(2).Copy(FromRange)
    End If
End Sub
    
'*****************************************************************************
'[ 関数名 ]　CutInsertCell
'[ 概  要 ]  切り取ったセルの挿入
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub CutInsertCell()
    Dim objActionRange As Range
    Dim strActionRange As String
    Dim objWkRange     As Range
    
    Set objActionRange = GetActionRange()
    strActionRange = objActionRange.Address

    'ワークのシートに作業域をコピー
    Set objWkRange = CopyToWkSheet(1, objActionRange, objActionRange)

    '切り取ったセルの挿入
    With objWkRange.Worksheet
        Call .Range(strFromRange).Cut
        Call .Range(strToRange).Insert
    End With
    
    With objWkRange.Worksheet
        Call .Range(strActionRange).Copy(Range(strActionRange))
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　CopyOnlyValue
'[ 概  要 ]  値のみコピー
'[ 引  数 ]　コピー元、コピー先
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub CopyOnlyValue(ByRef objFromRange As Range, ByRef objToRange As Range)
    'セル結合のコピー
    Call objToRange.UnMerge
    Call CopyMergeRange(objFromRange, objToRange)
    
    '値のみコピー
    Call objFromRange.Copy
    Call objToRange.PasteSpecial(xlPasteFormulas, xlNone, False, False)
End Sub

'*****************************************************************************
'[ 関数名 ]　CopyMergeRange
'[ 概  要 ]  領域の結合状態のみを複写する
'[ 引  数 ]　コピー元、コピー先
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub CopyMergeRange(ByRef objFromRange As Range, ByRef objToRange As Range)
    Dim objRange As Range
    Dim i As Long
    Dim j As Long
    
    '行の数だけループ
    For i = 1 To objFromRange.Rows.Count
        '列の数だけループ
        For j = 1 To objFromRange.Columns.Count
            With objFromRange(i, j).MergeArea
                If .Count > 1 Then
                    '結合セルの左上のセルなら
                    If objFromRange(i, j).Address = .Cells(1, 1).Address Then
                        '複写先のセルを結合
                        Call objToRange(i, j).Resize(.Rows.Count, .Columns.Count).Merge
                    End If
                End If
            End With
        Next j
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　MoveShape
'[ 概  要 ]  図形を移動またはコピーする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveShape()
    Dim objShapes As ShapeRange
    
    '回転している図形をグループ化する
    Call GroupSelection(GetSheeetShapeRange(ActiveSheet))
    
    '移動元の領域の図形を取得
    Set objShapes = SelectShapes(FromRange)
    
    '移動元→移動先
    If Not (objShapes Is Nothing) Then
        Call MoveShapes(objShapes, FromRange, ToRange, (enmModeType = E_Copy))
    End If
    
    'グループ化した図形の解除
    Call UnGroupSelection(GetSheeetShapeRange(ActiveSheet))
End Sub

'*****************************************************************************
'[ 関数名 ]　ExchangeShape
'[ 概  要 ]  図形を交換する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ExchangeShape()
    Dim objShapes(1 To 2) As ShapeRange '1:移動元の図形、2:移動先の図形
    
    '回転している図形をグループ化する
    Call GroupSelection(GetSheeetShapeRange(ActiveSheet))
    
    '移動元と移動先の領域の図形を取得
    Set objShapes(1) = SelectShapes(FromRange)
    Set objShapes(2) = SelectShapes(ToRange)
    
    '移動元→移動先
    If Not (objShapes(1) Is Nothing) Then
        Call MoveShapes(objShapes(1), FromRange, ToRange)
    End If
    
    '移動元←移動先
    If Not (objShapes(2) Is Nothing) Then
        Call MoveShapes(objShapes(2), ToRange, FromRange)
    End If
    
    'グループ化した図形の解除
    Call UnGroupSelection(GetSheeetShapeRange(ActiveSheet))
End Sub

'*****************************************************************************
'[ 関数名 ]　CutInsertShape
'[ 概  要 ]  切り取ったセルの挿入時の図形を交換する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub CutInsertShape()
    Dim objShapes(1 To 2) As ShapeRange '1:移動元の図形、2:以外の図形
    
    '回転している図形をグループ化する
    Call GroupSelection(GetSheeetShapeRange(ActiveSheet))
    
    '移動元と以外の領域の図形を取得
    Set objShapes(1) = SelectShapes(FromRange)
    Set objShapes(2) = SelectShapes(MinusRange(GetActionRange(), FromRange))
    
    '移動元→移動先
    If Not (objShapes(1) Is Nothing) Then
        Call MoveShapes(objShapes(1), FromRange, GetSlideRange(True))
    End If
    
    '移動元←移動先
    If Not (objShapes(2) Is Nothing) Then
        Call MoveShapes(objShapes(2), MinusRange(GetActionRange(), FromRange), GetSlideRange(False))
    End If
    
    'グループ化した図形の解除
    Call UnGroupSelection(GetSheeetShapeRange(ActiveSheet))
End Sub

'*****************************************************************************
'[ 関数名 ]　GetSlideRange
'[ 概　要 ]　切り取ったセルの挿入時、図形をスライドさせる領域を取得する
'[ 引　数 ]　True:移動元のスライド領域、False:以外のスライド領域
'[ 戻り値 ]　図形をスライドさせる領域
'*****************************************************************************
Private Function GetSlideRange(ByVal blnFrom As Boolean) As Range
    If blnFrom Then
        Select Case GetDirection()
        Case E_UP, E_LEFT
            Set GetSlideRange = ToRange
        Case E_DOWN
            Set GetSlideRange = OffsetRange(ToRange, -FromRange.Rows.Count)
        Case E_RIGHT
            Set GetSlideRange = OffsetRange(ToRange, , -FromRange.Columns.Count)
        End Select
    Else
        Select Case GetDirection()
        Case E_UP
            Set GetSlideRange = OffsetRange(MinusRange(GetActionRange(), FromRange), FromRange.Rows.Count)
        Case E_DOWN
            Set GetSlideRange = OffsetRange(MinusRange(GetActionRange(), FromRange), -FromRange.Rows.Count)
        Case E_LEFT
            Set GetSlideRange = OffsetRange(MinusRange(GetActionRange(), FromRange), , FromRange.Columns.Count)
        Case E_RIGHT
            Set GetSlideRange = OffsetRange(MinusRange(GetActionRange(), FromRange), , -FromRange.Columns.Count)
        End Select
    End If
End Function

'*****************************************************************************
'[ 関数名 ]　SelectShapes
'[ 概  要 ]  選択されている領域の中に含まれる図形を取得
'[ 引  数 ]　対象領域
'[ 戻り値 ]　領域に含まれる図形
'*****************************************************************************
Private Function SelectShapes(ByRef objRange As Range) As ShapeRange
    Dim i           As Long
    Dim objShape    As Shape
    Dim udtShape    As TRect
    Dim udtRange    As TRect
    ReDim lngIDArray(1 To ActiveSheet.Shapes.Count) As Variant
    
    '選択されたエリアに含まれる図形を取得
    For Each objShape In ActiveSheet.Shapes
        'Excel2007対応DPIRatioの倍数に補正
        With objShape
            udtShape.Height = .Height / DPIRatio
            udtShape.Left = .Left / DPIRatio
            udtShape.Top = .Top / DPIRatio
            udtShape.Width = .Width / DPIRatio
        End With
        With objRange
            udtRange.Height = .Height / DPIRatio
            udtRange.Left = .Left / DPIRatio
            udtRange.Top = .Top / DPIRatio
            udtRange.Width = .Width / DPIRatio
        End With
            
        If udtRange.Left <= udtShape.Left And _
           udtShape.Left + udtShape.Width <= udtRange.Left + udtRange.Width And _
           udtRange.Top <= udtShape.Top And _
           udtShape.Top + udtShape.Height <= udtRange.Top + udtRange.Height Then
            i = i + 1
            lngIDArray(i) = objShape.ID
        End If
    Next objShape

    If i > 0 Then
        ReDim Preserve lngIDArray(1 To i)
        Set SelectShapes = GetShapeRangeFromID(lngIDArray)
    End If
End Function

'*****************************************************************************
'[ 関数名 ]　MoveShapes
'[ 概  要 ]  図形を移動する
'[ 引  数 ]　移動元の領域の図形、移動元のセル領域、移動先のセル領域、True:コピーする
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveShapes(ByRef objShapes As ShapeRange, ByRef objFromRange As Range, _
                       ByRef objToRange As Range, Optional ByVal blnCopy As Boolean = False)
    Dim objShape    As Shape
    Dim i           As Long
    ReDim udtShapes(1 To objShapes.Count) As TShape
    
    '選択されたエリアに含まれる図形を取得
    For Each objShape In objShapes
        With objShape
            i = i + 1
            udtShapes(i).ID = .ID
            udtShapes(i).Placement = .Placement
            udtShapes(i).TopRow = .TopLeftCell.Row - objFromRange.Row + 1
            udtShapes(i).LeftColumn = .TopLeftCell.Column - objFromRange.Column + 1
            udtShapes(i).BottomRow = .BottomRightCell.Row - objFromRange.Row + 1
            udtShapes(i).RightColumn = .BottomRightCell.Column - objFromRange.Column + 1
            udtShapes(i).Height = .Top + .Height - .BottomRightCell.Top
            udtShapes(i).Width = .Left + .Width - .BottomRightCell.Left
            udtShapes(i).Top = .Top - .TopLeftCell.Top
            udtShapes(i).Left = .Left - .TopLeftCell.Left
        End With
    Next objShape
    
    For i = 1 To objShapes.Count
'        If udtShapes(i).Placement <> xlFreeFloating Then
            If blnCopy Then
                '複写する
                Call SetRect(GetShapeFromID(udtShapes(i).ID).Duplicate, objToRange, udtShapes(i))
            Else
                '移動する
                Call SetRect(GetShapeFromID(udtShapes(i).ID), objToRange, udtShapes(i))
            End If
'        End If
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　SetRect
'[ 概  要 ]  図形の位置を移動先の領域に設定する
'[ 引  数 ]　対象の図形、移動先の領域、位置情報
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SetRect(ByRef objShape As Shape, ByRef objRange As Range, ByRef udtShape As TShape)
    '左上の位置を設定する
'    If udtShape.Placement <> xlFreeFloating Then
        objShape.Top = objRange.Rows(udtShape.TopRow).Top + udtShape.Top
        objShape.Left = objRange.Columns(udtShape.LeftColumn).Left + udtShape.Left
'    End If

    '幅と高さを設定する
'    If udtShape.Placement = xlMoveAndSize Then
        objShape.Height = objRange.Rows(udtShape.BottomRow).Top + udtShape.Height - objShape.Top
        objShape.Width = objRange.Columns(udtShape.RightColumn).Left + udtShape.Width - objShape.Left
'    End If
End Sub

'*****************************************************************************
'[イベント]　cmdCancel_Click
'[ 概  要 ]　キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[イベント]　cmdHelp_Click
'[ 概  要 ]　ヘルプボタン押下時
'*****************************************************************************
Private Sub cmdHelp_Click()
    Call OpenHelpPage("MoveCell.htm")
End Sub

'*****************************************************************************
'[イベント]　チェックボタン_Change
'[ 概  要 ]　チェックボタン変更時
'*****************************************************************************
Private Sub chkCopy_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkCopy.Value Then
        enmModeType = E_Copy
    Else
        enmModeType = E_Move
    End If
    Call ChangeMode
End Sub
Private Sub chkCutInsert_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkCutInsert.Value Then
        enmModeType = E_CutInsert
    Else
        enmModeType = E_Move
    End If
    Call ChangeMode
End Sub
Private Sub chkExchange_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkExchange.Value Then
        enmModeType = E_Exchange
    Else
        enmModeType = E_Move
    End If
    Call ChangeMode
End Sub

'*****************************************************************************
'[ 関数名 ]　ChangeMode
'[ 概  要 ]  ｢移動｣･｢コピー｣･｢入替え｣のモードを変更する
'[ 引  数 ]　なし
'[ 戻り値 ]　True:キャンセル時
'*****************************************************************************
Private Sub ChangeMode()
    
    blnCheck = True
    Select Case enmModeType
    Case E_Move
        chkCopy.Value = False
        chkExchange.Value = False
        chkCutInsert.Value = False
        chkOnlyValue.Value = False
        chkOnlyValue.Enabled = False
    Case E_Copy
        chkCopy.Value = True
        chkExchange.Value = False
        chkCutInsert.Value = False
        chkOnlyValue.Enabled = True
    Case E_Exchange
        chkCopy.Value = False
        chkExchange.Value = True
        chkCutInsert.Value = False
        chkOnlyValue.Enabled = True
    Case E_CutInsert
        chkCopy.Value = False
        chkExchange.Value = False
        chkCutInsert.Value = True
        chkOnlyValue.Value = False
        chkOnlyValue.Enabled = False
    End Select
    blnCheck = False
    
    Select Case enmModeType
    Case E_Move
        lblTitle.Caption = "移動先"
    Case E_Copy
        lblTitle.Caption = "コピー先"
    Case E_Exchange
        lblTitle.Caption = "入替え先"
    Case E_CutInsert
        lblTitle.Caption = "挿入先"
    End Select

    'テキストボックスを編集
    Call EditTextbox
End Sub

'*****************************************************************************
'[イベント]　KeyDown
'[ 概  要 ]　カーソルキーで移動先を変更させる
'*****************************************************************************
Private Sub cmdOK_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub cmdCancel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub cmdHelp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub chkCopy_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub chkExchange_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub chkCutInsert_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub chkOnlyValue_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub fraKeyCapture_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub

'*****************************************************************************
'[イベント]　UserForm_KeyDown
'[ 概  要 ]　カーソルキーで移動先を変更させる
'*****************************************************************************
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i         As Long
    Dim lngTop    As Long
    Dim lngLeft   As Long
    Dim lngBottom As Long
    Dim lngRight  As Long
    Dim blnChk     As Boolean
    Dim objToRange As Range
    
    If enmModeType = E_Move Then
        Select Case (KeyCode)
        Case vbKeyInsert, vbKeyDelete
            If InsertOrDeleteCell(KeyCode) = True Then
                Exit Sub
            End If
        End Select
    End If
    
    Select Case (KeyCode)
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown, vbKeyHome
        Call fraKeyCapture.SetFocus
    Case Else
        Exit Sub
    End Select
    
    'Altが押されていればスクロール
    If GetKeyState(vbKeyMenu) < 0 Then
        Select Case (KeyCode)
        Case vbKeyLeft
            Call ActiveWindow.SmallScroll(, , , 1)
        Case vbKeyRight
            Call ActiveWindow.SmallScroll(, , 1)
        Case vbKeyUp
            Call ActiveWindow.SmallScroll(, 1)
        Case vbKeyDown
            Call ActiveWindow.SmallScroll(1)
        End Select
        Exit Sub
    End If
    
    'Zoom
    Select Case (KeyCode)
    Case vbKeyHome, vbKeyPageUp, vbKeyPageDown
        Select Case (KeyCode)
        Case vbKeyHome
            ActiveWindow.Zoom = lngZoom
        Case vbKeyPageUp
            ActiveWindow.Zoom = WorksheetFunction.Min(ActiveWindow.Zoom + 10, 400)
        Case vbKeyPageDown
            ActiveWindow.Zoom = WorksheetFunction.Max(ActiveWindow.Zoom - 10, 10)
        End Select
        
        With ToRange
            lngLeft = WorksheetFunction.Max(.Left / DPIRatio - 1, 0) * ActiveWindow.Zoom / 100
            lngTop = WorksheetFunction.Max(.Top / DPIRatio - 1, 0) * ActiveWindow.Zoom / 100
            Call ActiveWindow.ScrollIntoView(lngLeft, lngTop, 1, 1)
        End With
        Exit Sub
    End Select
    
    '選択領域の四方の位置を待避
    With Range(strToRange)
        lngTop = .Row
        lngBottom = .Row + .Rows.Count - 1
        lngLeft = .Column
        lngRight = .Column + .Columns.Count - 1
    End With
    
    Select Case (Shift)
    Case 0
        '選択領域を移動
        Select Case (KeyCode)
        Case vbKeyLeft
            lngLeft = lngLeft - 1
            lngRight = lngRight - 1
        Case vbKeyRight
            lngLeft = lngLeft + 1
            lngRight = lngRight + 1
        Case vbKeyUp
            lngTop = lngTop - 1
            lngBottom = lngBottom - 1
        Case vbKeyDown
            lngTop = lngTop + 1
            lngBottom = lngBottom + 1
        End Select
    Case 1
        '選択領域の大きさを変更
        If GetKeyState(vbKeyZ) < 0 Then
            Select Case (KeyCode)
            Case vbKeyLeft
                lngLeft = lngLeft - 1
            Case vbKeyRight
                lngLeft = lngLeft + 1
            Case vbKeyUp
                lngTop = lngTop - 1
            Case vbKeyDown
                lngTop = lngTop + 1
            End Select
        Else
            Select Case (KeyCode)
            Case vbKeyLeft
                lngRight = lngRight - 1
            Case vbKeyRight
                lngRight = lngRight + 1
            Case vbKeyUp
                lngBottom = lngBottom - 1
            Case vbKeyDown
                lngBottom = lngBottom + 1
            End Select
        End If
    Case Else
        Exit Sub
    End Select
    
    'チェック
    If (lngLeft <= lngRight And lngTop <= lngBottom) And _
       (1 <= lngLeft And lngRight <= Columns.Count) And _
       (1 <= lngTop And lngBottom <= Rows.Count) Then
        Set objToRange = Range(Cells(lngTop, lngLeft), Cells(lngBottom, lngRight))
    Else
        Exit Sub
    End If
    
    '新しい選択領域を示すRange
    strToRange = objToRange.Address
    
    'テキストボックスを編集
    Call EditTextbox
    
    'テキストボックスを移動
    With Range(strToRange)
        objTextbox.Left = .Left
        objTextbox.Top = .Top
        objTextbox.Width = .Width
        objTextbox.Height = .Height
    End With
    
    '移動先(コピー先)の編集
    lblCellAddress.Caption = " " & Replace$(strToRange, "$", "")

    '選択領域が画面から消えたら画面をスクロール
    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '画面分割のない時
        Select Case (KeyCode)
        Case vbKeyLeft
            i = WorksheetFunction.Max(ToRange.Column - 1, 1)
            If IntersectRange(ActiveWindow.VisibleRange, Columns(i)) Is Nothing Then
                Call ActiveWindow.SmallScroll(, , , 1)
            End If
        Case vbKeyRight
            i = WorksheetFunction.Min(ToRange.Column + ToRange.Columns.Count, Columns.Count)
            If IntersectRange(ActiveWindow.VisibleRange, Columns(i)) Is Nothing Then
                Call ActiveWindow.SmallScroll(, , 1)
            End If
        Case vbKeyUp
            i = WorksheetFunction.Max(ToRange.Row - 1, 1)
            If IntersectRange(ActiveWindow.VisibleRange, Rows(i)) Is Nothing Then
                Call ActiveWindow.SmallScroll(, 1)
            End If
        Case vbKeyDown
            i = WorksheetFunction.Min(ToRange.Row + ToRange.Rows.Count, Rows.Count)
            If IntersectRange(ActiveWindow.VisibleRange, Rows(i)) Is Nothing Then
                Call ActiveWindow.SmallScroll(1)
            End If
        End Select
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　InsertOrDeleteCell
'[ 概  要 ]  セルを挿入または削除する
'[ 引  数 ]　なし
'[ 戻り値 ]　True:キャンセル時
'*****************************************************************************
Private Function InsertOrDeleteCell(ByVal lngKeyCode As Long) As Boolean
    Dim strMsg      As String
    Dim lngSelect   As Long
    
    Select Case (lngKeyCode)
    Case vbKeyDelete
        strMsg = "列方向を削除しますか？" & vbCrLf
        strMsg = strMsg & "　「 はい 」････ 左方向にシフトする" & vbCrLf
        strMsg = strMsg & "　「いいえ」････ 上方向にシフトする"
    Case vbKeyInsert
        strMsg = "列方向に挿入しますか？" & vbCrLf
        strMsg = strMsg & "　「 はい 」････ 右方向にシフトする" & vbCrLf
        strMsg = strMsg & "　「いいえ」････ 下方向にシフトする"
    End Select
    
    Select Case True
    Case strFromRange = strToRange
        lngSelect = MsgBox(strMsg, vbYesNoCancel + vbDefaultButton1, "選択して下さい")
    Case FromRange.EntireRow.Address = ToRange.EntireRow.Address
        lngSelect = vbYes
    Case FromRange.EntireColumn.Address = ToRange.EntireColumn.Address
        lngSelect = vbNo
    Case Else
        lngSelect = MsgBox(strMsg, vbYesNoCancel + vbDefaultButton1, "選択して下さい")
    End Select
    If lngSelect = vbCancel Then
        InsertOrDeleteCell = True
        Exit Function
    End If
    
    If CheckInsertOrDelete(lngSelect = vbYes) = False Then
        Call MsgBox("結合されたセルの一部を変更することはできません", vbExclamation)
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    Call objTextbox.Delete
    Set objTextbox = Nothing

On Error GoTo ErrHandle
    'アンドゥ用に元の状態を保存する
    Call SaveUndoInfo(E_MoveCell, FromRange, Cells)
    
    Select Case (lngKeyCode)
    Case vbKeyDelete
        Select Case (lngSelect)
        Case vbYes
            Call ToRange.Delete(xlToLeft)
        Case vbNo
            Call ToRange.Delete(xlUp)
        End Select
    Case vbKeyInsert
        Select Case (lngSelect)
        Case vbYes
            Call ToRange.Insert(xlToRight)
        Case vbNo
            Call ToRange.Insert(xlDown)
        End Select
    End Select
    
    Call ToRange.Select
    Call Unload(Me)
    Call SetOnUndo
Exit Function
ErrHandle:
    Call ToRange.Select
    Call Unload(Me)
End Function

'*****************************************************************************
'[ 関数名 ]　CheckInsertOrDelete
'[ 概  要 ]  Ins/Delキー押下時の結合セルのチェック
'[ 引  数 ]　True：列方向のシフト、False：行方向のシフト
'[ 戻り値 ]　False:エラーあり
'*****************************************************************************
Private Function CheckInsertOrDelete(ByVal blnSelect As Boolean) As Boolean
    Dim objWkRange   As Range
    Dim objActRange  As Range
    Dim objActArea   As Range
    Dim objChkRange  As Range

    Set objWkRange = ArrangeRange(ToRange)
    If blnSelect Then
        '列方向のシフトの時、選択領域が行方向に結合がはみ出していないか判定
        Set objChkRange = MinusRange(objWkRange, ToRange.EntireRow)
    Else
        '行方向のシフトの時、選択領域が列方向に結合がはみ出していないか判定
        Set objChkRange = MinusRange(objWkRange, ToRange.EntireColumn)
    End If
    If Not (objChkRange Is Nothing) Then
        CheckInsertOrDelete = False
        Exit Function
    End If
        
    If blnSelect Then
        '列方向のシフトの時、選択領域から最後の列の範囲を取得
        Set objActRange = Range(ToRange, Cells(objWkRange.Row, Columns.Count))
        Set objActArea = objActRange.EntireColumn
    Else
        '行方向のシフトの時、選択領域から最後の行の範囲を取得
        Set objActRange = Range(ToRange, Cells(Rows.Count, objWkRange.Column))
        Set objActArea = objActRange.EntireRow
    End If
    'シートの最後の行または列の範囲で､結合セルがはみ出していないか判定
    Set objChkRange = IntersectRange(MinusRange(ArrangeRange(objActRange), objActRange), objActArea)
    CheckInsertOrDelete = (objChkRange Is Nothing)
End Function
    
'*****************************************************************************
'[ 関数名 ]　EditTextbox
'[ 概  要 ]  テキストボックスを編集する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub EditTextbox()
On Error GoTo ErrHandle
'    If blnSameSheet Then
    Select Case enmModeType
    Case E_Copy, E_Exchange
        If Not (IntersectRange(FromRange, ToRange) Is Nothing) Then
            Select Case enmModeType
            Case E_Copy
                Call Err.Raise(C_CheckErrMsg, , "元の領域にはコピーできません")
            Case E_Exchange
                Call Err.Raise(C_CheckErrMsg, , "元の領域とは入替えできません")
            End Select
        End If
    End Select
'    End If

    Select Case enmModeType
    Case E_Move, E_Copy, E_Exchange
        If CheckBorder() = True Then
            Call Err.Raise(C_CheckErrMsg, , "結合されたセルの一部を変更することはできません")
        End If
        If FromRange.Columns.Count = Columns.Count And _
           ToRange.Columns.Count <> Columns.Count Then
            Call Err.Raise(C_CheckErrMsg, , "すべての列の選択時は列の幅を変更できません")
        End If
    Case E_CutInsert
        Call CheckCutInsertRange
    End Select
    
    With objTextbox.TextFrame.Characters
        .Text = Replace$(strToRange, "$", "")
        .Font.ColorIndex = 0
        cmdOK.Enabled = True  'OKボタンを使用可にする
    End With
Exit Sub
ErrHandle:
    If Err.Number = C_CheckErrMsg Then
        With objTextbox.TextFrame.Characters
            .Text = Err.Description
            .Font.ColorIndex = 3
            cmdOK.Enabled = False 'OKボタンを使用不可にする
        End With
    Else
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　CheckCutInsertRange
'[ 概  要 ]  切り取った領域の挿入のときの領域のチェック
'[ 引  数 ]　なし
'[ 戻り値 ]　なし（エラーの時は例外）
'*****************************************************************************
Private Sub CheckCutInsertRange()
    Dim enmDirection   As EDirection
    Dim objActionRange As Range
    
    If FromRange.Columns.Count <> ToRange.Columns.Count Or _
       FromRange.Rows.Count <> ToRange.Rows.Count Then
        Call Err.Raise(C_CheckErrMsg, , "挿入時は大きさを変更できません")
    End If
    
    enmDirection = GetDirection()
    Select Case enmDirection
    Case E_Non
        Call Err.Raise(C_CheckErrMsg, , "挿入先に移動して下さい")
    Case E_ERR
        Call Err.Raise(C_CheckErrMsg, , "真上・真下・真横以外には挿入できません")
    End Select
        
    Set objActionRange = GetActionRange()
    If objActionRange.Address = strFromRange Then
        Call Err.Raise(C_CheckErrMsg, , "元の領域には挿入できません")
    End If
    
    If IsBorderMerged(objActionRange) Then
        Call Err.Raise(C_CheckErrMsg, , "結合されたセルの一部を変更することはできません")
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　CheckBorder
'[ 概  要 ]  移動先の四方が結合セルをまたいでいるかどうか
'            移動の時は、元の領域の結合セルは対象としない
'[ 引  数 ]　なし
'[ 戻り値 ]　True:結合セルをまたいでいる、False:いない
'*****************************************************************************
Private Function CheckBorder() As Boolean
    Dim objChkRange As Range
    
    Select Case enmModeType
    Case E_Move
        Set objChkRange = MinusRange(ArrangeRange(ToRange), UnionRange(ToRange, FromRange))
    Case E_Copy, E_Exchange
        Set objChkRange = MinusRange(ArrangeRange(ToRange), ToRange)
    End Select
    
    CheckBorder = Not (objChkRange Is Nothing)
End Function

'*****************************************************************************
'[ 関数名 ]　CopyToWkSheet
'[ 概  要 ]  ワークシートに領域をコピーする
'[ 引  数 ]　Workarea1を使用するかWorkarea2を使用するか判別
'            コピー元のアドレス、コピー先のアドレス
'[ 戻り値 ]  ワークシートのコピーされたRange(サイズはコピー元のサイズ)
'*****************************************************************************
Private Function CopyToWkSheet(ByVal bytNo As Byte, ByRef objFromRange As Range, ByRef objToRange As Range) As Range
    Dim objWorksheet As Worksheet
    
    Set objWorksheet = ThisWorkbook.Worksheets("Workarea" & bytNo)
    Call DeleteSheet(objWorksheet)
    
    '「Workarea」シートに選択領域を複写する
    With objFromRange
        Set CopyToWkSheet = objWorksheet.Range(objToRange.Address).Resize(.Rows.Count, .Columns.Count)
        Call .Copy(CopyToWkSheet)
    End With
End Function
    
'*****************************************************************************
'[ 関数名 ]　ReSizeArea
'[ 概  要 ]  ワークエリアを使用し領域のサイズを変更する
'[ 引  数 ]　objRange：変更する領域
'            lngNewRow、lngNewCol：新しいサイズ
'[ 戻り値 ]  サイズを変更したRange
'*****************************************************************************
Private Function ReSizeArea(ByRef objRange As Range, ByVal lngNewRow As Long, ByVal lngNewCol As Long) As Range
    Dim objWkRange   As Range
    Dim lngDiff      As Long
    Dim lngOldRow    As Long
    Dim lngOldCol    As Long
    
    '変更前のサイズを取得
    With objRange
        lngOldRow = .Rows.Count
        lngOldCol = .Columns.Count
    End With
    
    '行のサイズ変更
    lngDiff = lngNewRow - lngOldRow
    If lngDiff <> 0 Then
        Call ChangeRangeRowSize(objRange, lngOldRow, lngOldCol, lngDiff)
    End If
    
    '列のサイズ変更
    lngDiff = lngNewCol - lngOldCol
    If lngDiff <> 0 Then
        Call ChangeRangeColSize(objRange, lngOldRow, lngOldCol, lngDiff)
    End If
    
    Set ReSizeArea = objRange.Resize(lngNewRow, lngNewCol)
End Function


'*****************************************************************************
'[ 関数名 ]　ChangeRangeRowSize
'[ 概  要 ]  領域のサイズを変更する
'[ 引  数 ]　objWkRange:サイズを変更する領域､元の行数と列数、lngDiff:行の変更サイズ
'[ 戻り値 ]  なし
'*****************************************************************************
Private Sub ChangeRangeRowSize(ByRef objWkRange As Range, ByVal lngRowCount As Long, ByVal lngColCount As Long, ByVal lngDiff As Long)
    Dim i            As Long
    Dim udtBorder    As TBorder '罫線の種類
    Dim objWorksheet As Worksheet
    
    Set objWorksheet = objWkRange.Worksheet
    
    Select Case (lngDiff)
    Case Is < 0 '縮小
        '下端の罫線を保存
        With objWkRange.Borders
            udtBorder = GetBorder(.Item(xlEdgeBottom))
        End With
        '削除
        With objWkRange
            Call objWorksheet.Range(.Rows(lngRowCount + lngDiff + 1), _
                                    .Rows(lngRowCount)).Delete(xlShiftUp)
        End With
    Case Is > 0 '拡張
        '挿入
        With objWkRange.Rows
            Call .Item(.Count).Copy(objWorksheet.Range(.Item(lngRowCount + 1), .Item(lngRowCount + lngDiff)))
            Call objWorksheet.Range(.Item(lngRowCount + 1), .Item(lngRowCount + lngDiff)).ClearContents
        End With
        
        '列の数だけループ
        For i = 1 To lngColCount
            'セルを結合
            With objWkRange.Rows(lngRowCount)
                Call objWorksheet.Range(.Cells(1, i), .Cells(lngDiff + 1, i).MergeArea).Merge
            End With
        Next i
    End Select
    
    Set objWkRange = objWkRange.Resize(lngRowCount + lngDiff)
    
    '削除の時下端の罫線を再設定
    If lngDiff < 0 Then
        '下端の罫線を再設定
        With objWkRange.Borders
            Call SetBorder(udtBorder, .Item(xlEdgeBottom))
        End With
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　ChangeRangeColSize
'[ 概  要 ]  領域のサイズを変更する
'[ 引  数 ]　objWkRange:サイズを変更する領域､元の行数と列数、lngDiff:列の変更サイズ
'[ 戻り値 ]  なし
'*****************************************************************************
Private Sub ChangeRangeColSize(ByRef objWkRange As Range, ByVal lngRowCount As Long, ByVal lngColCount As Long, ByVal lngDiff As Long)
    Dim i            As Long
    Dim udtBorder    As TBorder '罫線の種類
    Dim objWorksheet As Worksheet
    
    Set objWorksheet = objWkRange.Worksheet
    
    Select Case (lngDiff)
    Case Is < 0 '縮小
        '右端の罫線を保存
        With objWkRange.Borders
            udtBorder = GetBorder(.Item(xlEdgeRight))
        End With
        '削除
        With objWkRange
            Call objWorksheet.Range(.Columns(lngColCount + lngDiff + 1), _
                                    .Columns(lngColCount)).Delete(xlShiftToLeft)
        End With
    Case Is > 0 '拡張
        '挿入
        With objWkRange.Columns
            Call .Item(.Count).Copy(objWorksheet.Range(.Item(lngColCount + 1), .Item(lngColCount + lngDiff)))
            Call objWorksheet.Range(.Item(lngColCount + 1), .Item(lngColCount + lngDiff)).ClearContents
        End With
        
        '行の数だけループ
        For i = 1 To lngRowCount
            'セルを結合
            With objWkRange.Columns(lngColCount)
                Call objWorksheet.Range(.Cells(i, 1), .Cells(i, lngDiff + 1).MergeArea).Merge
            End With
        Next i
    End Select
    
    Set objWkRange = objWkRange.Resize(, lngColCount + lngDiff)
    
    '削除の時右端の罫線を再設定
    If lngDiff < 0 Then
        '右端の罫線を再設定
        With objWkRange.Borders
            Call SetBorder(udtBorder, .Item(xlEdgeRight))
        End With
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　GetDirection
'[ 概  要 ]  移動方向を取得
'[ 引  数 ]　なし
'[ 戻り値 ]  移動方向：移動なし・上・下・左・右・以外(エラー)
'*****************************************************************************
Private Function GetDirection() As EDirection
    Select Case True
    Case ToRange.Row = FromRange.Row And ToRange.Column = FromRange.Column
        GetDirection = E_Non
    Case ToRange.Row <> FromRange.Row And ToRange.Column <> FromRange.Column
        GetDirection = E_ERR
    Case ToRange.Row < FromRange.Row
        GetDirection = E_UP
    Case ToRange.Row > FromRange.Row
        GetDirection = E_DOWN
    Case ToRange.Column < FromRange.Column
        GetDirection = E_LEFT
    Case ToRange.Column > FromRange.Column
        GetDirection = E_RIGHT
    End Select
End Function

'*****************************************************************************
'[ 関数名 ]　GetActionRange
'[ 概  要 ]　切り取ったセルの挿入の対象となる領域
'[ 引  数 ]　なし
'[ 戻り値 ]  対象領域
'*****************************************************************************
Private Function GetActionRange() As Range
    Select Case GetDirection()
    Case E_UP, E_LEFT
        '作業元から移動先で囲まれた領域を対象
        Set GetActionRange = Range(FromRange, ToRange)
    Case E_DOWN
        '作業元から移動先の1行前までを対象
        Set GetActionRange = Range(FromRange, ToRange(0, 1))
    Case E_RIGHT
        '作業元から移動先の1列前までを対象
        Set GetActionRange = Range(FromRange, ToRange(1, 0))
    End Select
End Function

'*****************************************************************************
'[プロパティ]　FromRange
'[ 概  要 ]　移動元の領域
'*****************************************************************************
Private Property Get FromRange() As Range
    Set FromRange = objFromSheet.Range(strFromRange)
End Property

'*****************************************************************************
'[プロパティ]　ToRange
'[ 概  要 ]　移動先の領域
'*****************************************************************************
Private Property Get ToRange() As Range
    Set ToRange = Range(strToRange)
End Property

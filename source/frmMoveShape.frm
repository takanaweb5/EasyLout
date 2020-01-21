VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMoveShape 
   Caption         =   "図形の移動"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   OleObjectBlob   =   "frmMoveShape.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMoveShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type TRect
    Top      As Double
    Height   As Double
    Left     As Double
    Width    As Double
End Type

Private Type TShapes  'Undo情報
    Shapes() As TRect
End Type

Private udtShapes(1 To 100) As TShapes
Private lngUndoCount   As Long
Private objTmpBar      As CommandBar
Private objShapeRange  As ShapeRange
Private blnChange      As Boolean
Private blnOK          As Boolean
Private blnCancel      As Boolean
Private lngZoom        As Long
Private objDummy       As Shape

'*****************************************************************************
'[イベント]　UserForm_Initialize
'[ 概  要 ]　フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()
    Dim i         As Long
    
    '呼び元に通知する
    blnFormLoad = True
    lngZoom = ActiveWindow.Zoom
    
    Set objShapeRange = Selection.ShapeRange
    
    '「配置/整列」のポップアップメニュー作成
    Set objTmpBar = CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
    With objTmpBar.Controls
        .Add(, 664).BeginGroup = False '左揃え
        .Add(, 668).BeginGroup = False '左右中央揃え
        .Add(, 665).BeginGroup = False '右揃え
        .Add(, 666).BeginGroup = True  '上揃え
        .Add(, 669).BeginGroup = False '上下中央揃え
        .Add(, 667).BeginGroup = False '下揃え
        
        With .Add(, 408)
            .BeginGroup = True
            .Caption = "左右に等間隔で整列(&H)"
        End With
        With .Add(, 465)
            .BeginGroup = False
            .Caption = "上下に等間隔で整列(&V)"
        End With
        With .Add(, 408)
            .BeginGroup = True
            .Caption = "横方向に連結(&J)"
        End With
        With .Add(, 465)
            .BeginGroup = False
            .Caption = "縦方向に連結(&K)"
        End With
        With .Add(, 542)
            .BeginGroup = True
            .FaceId = 2067
            .Caption = "幅を揃える(&W)"
        End With
        With .Add(, 541)
            .BeginGroup = False
            .FaceId = 2068
            .Caption = "高さを揃える(&E)"
        End With
        With .Add(, 549)
            .BeginGroup = True
            .FaceId = 550
            .Caption = "グリッドに合せる(&G)"
        End With
    End With

    For i = 1 To objTmpBar.Controls.Count
        objTmpBar.Controls(i).OnAction = "OnPopupClick"
        objTmpBar.Controls(i).Tag = i
    Next i
    
    Select Case objShapeRange.Count
    Case 1
        objTmpBar.Controls(1).Enabled = False  '左揃え
        objTmpBar.Controls(2).Enabled = False  '左右中央揃え
        objTmpBar.Controls(3).Enabled = False  '右揃え
        objTmpBar.Controls(4).Enabled = False  '上揃え
        objTmpBar.Controls(5).Enabled = False  '上下中央揃え
        objTmpBar.Controls(6).Enabled = False  '下揃え
        objTmpBar.Controls(7).Enabled = False  '左右に等間隔で整列
        objTmpBar.Controls(8).Enabled = False  '上下に等間隔で整列
        objTmpBar.Controls(9).Enabled = False  '横方向に連結
        objTmpBar.Controls(10).Enabled = False '縦方向に連結
        objTmpBar.Controls(11).Enabled = False '幅を揃える
        objTmpBar.Controls(12).Enabled = False '高さを揃える
    Case 2
        objTmpBar.Controls(7).Enabled = False  '左右に整列
        objTmpBar.Controls(8).Enabled = False  '上下に整列
    End Select

    '「グリッドにあわせる」をチェックするか判定
    If CommandBars.ActionControl.Caption = "図形をグリッドに合せる" Then
        chkGrid.Value = True
    Else
        If objTmpBar.FindControl(, 549).State = True Then
            chkGrid.Value = True
        Else
            chkGrid.Value = False
        End If
    End If
    objTmpBar.FindControl(, 549).State = False
End Sub

'*****************************************************************************
'[イベント]　UserForm_Terminate
'[ 概  要 ]　フォームアンロード時
'*****************************************************************************
Private Sub UserForm_Terminate()
    '呼び元に通知する
    blnFormLoad = False
    
    '「配置/整列」のポップアップメニュー削除
    Call objTmpBar.Delete
End Sub

'*****************************************************************************
'[イベント]　UserForm_QueryClose
'[ 概  要 ]　×ボタンでフォームを閉じる時、変更を元に戻す
'*****************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Call objShapeRange.Select
    
    '倍率を元に戻す
    If ActiveWindow.Zoom <> lngZoom Then
        ActiveWindow.Zoom = lngZoom
    End If
    
    '変更がなければフォームを閉じる
    If blnChange = False Then
        Exit Sub
    End If
        
    '×ボタンでフォームを閉じる時、フォームを閉じない
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Exit Sub
    End If
    
    If Not (objDummy Is Nothing) Then
        Call objDummy.Delete
    End If
    
    'グループ化した図形の解除
    Call UnGroupSelection(objShapeRange).Select
    
    Select Case True
    Case blnOK
        Call SetOnUndo
    Case blnCancel
        Call ExecUndo
    End Select
End Sub

'*****************************************************************************
'[イベント]　cmdOK_Click
'[ 概  要 ]　ＯＫボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
    blnOK = True
    Call Unload(Me)
End Sub

'*****************************************************************************
'[イベント]　cmdCancel_Click
'[ 概  要 ]　キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    blnCancel = True
    Call Unload(Me)
End Sub

'*****************************************************************************
'[イベント]　cmdAlign_Click
'[ 概  要 ]　「配置/整列」ボタンクリック
'*****************************************************************************
Private Sub cmdAlign_Click()
    Call objTmpBar.ShowPopup
    Call fraKeyCapture.SetFocus
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
Private Sub chkGrid_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub cmdAlign_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub fraKeyCapture_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub

'*****************************************************************************
'[イベント]　MouseDown
'[ 概  要 ]　右クリックでポップアップメニューを表示する
'*****************************************************************************
Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        Call objTmpBar.ShowPopup
    End If
End Sub
Private Sub cmdOK_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call UserForm_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub cmdCancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call UserForm_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub chkGrid_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call UserForm_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub cmdAlign_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call UserForm_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub fraKeyCapture_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call UserForm_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub lblLabel1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call UserForm_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub lblLabel2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call UserForm_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub lblLabel3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call UserForm_MouseDown(Button, Shift, X, Y)
End Sub

'*****************************************************************************
'[イベント]　UserForm_KeyDown
'[ 概  要 ]　カーソルキーで移動先を変更させる
'*****************************************************************************
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim blnGrid As Boolean
        
    '[Ctrl]+Z が押下されている時、Undoを行う
    If (Shift = 2) And (KeyCode = vbKeyZ) Then
        Call PopUndoInfo
        Call fraKeyCapture.SetFocus
        KeyCode = 0
        Exit Sub
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
        Call objShapeRange.Select
        Select Case (KeyCode)
        Case vbKeyHome
            If ActiveWindow.Zoom = lngZoom Then
                'Excelの機能を利用して、図形を表示できる位置に画面をスクロールさせるため
                If lngZoom > 100 Then
                    ActiveWindow.Zoom = ActiveWindow.Zoom - 10
                Else
                    ActiveWindow.Zoom = ActiveWindow.Zoom + 10
                End If
            End If
            ActiveWindow.Zoom = lngZoom
        Case vbKeyPageUp
            ActiveWindow.Zoom = WorksheetFunction.Min(ActiveWindow.Zoom + 10, 400)
        Case vbKeyPageDown
            ActiveWindow.Zoom = WorksheetFunction.Max(ActiveWindow.Zoom - 10, 10)
        End Select
        If Not (objDummy Is Nothing) Then
            Call objDummy.Select
        End If
        Exit Sub
    End Select
    
    '変更前の情報を保存
    Call SaveBeforeChange
    
    '[Ctrl]Keyが押下されている or グリッドにあわせるがチェックされていない
    If (GetKeyState(vbKeyControl) < 0) Or chkGrid.Value = False Then
        blnGrid = False
    Else
        blnGrid = True
    End If
    
On Error GoTo ErrHandle
    Dim dblSave    As Double
    Dim strMove    As String
    
    If GetKeyState(vbKeyShift) < 0 Then
        '図形の大きさを変更
        If GetKeyState(vbKeyZ) < 0 Then
            Select Case (KeyCode)
            Case vbKeyLeft
                Call ChangeShapesWidth(objShapeRange, 1, blnGrid, True)
                strMove = "Left"
            Case vbKeyRight
                Call ChangeShapesWidth(objShapeRange, -1, blnGrid, True)
            Case vbKeyUp
                Call ChangeShapesHeight(objShapeRange, 1, blnGrid, True)
                strMove = "Up"
            Case vbKeyDown
                Call ChangeShapesHeight(objShapeRange, -1, blnGrid, True)
            End Select
        Else
            Select Case (KeyCode)
            Case vbKeyLeft
                Call ChangeShapesWidth(objShapeRange, -1, blnGrid, False)
            Case vbKeyRight
                Call ChangeShapesWidth(objShapeRange, 1, blnGrid, False)
                strMove = "Right"
            Case vbKeyUp
                Call ChangeShapesHeight(objShapeRange, -1, blnGrid, False)
            Case vbKeyDown
                Call ChangeShapesHeight(objShapeRange, 1, blnGrid, False)
                strMove = "Down"
            End Select
        End If
    Else
        '図形を移動
        Select Case (KeyCode)
        Case vbKeyLeft
            Call MoveShapesLR(objShapeRange, -1, blnGrid)
            strMove = "Left"
        Case vbKeyRight
            Call MoveShapesLR(objShapeRange, 1, blnGrid)
            strMove = "Right"
        Case vbKeyUp
            Call MoveShapesUD(objShapeRange, -1, blnGrid)
            strMove = "Up"
        Case vbKeyDown
            Call MoveShapesUD(objShapeRange, 1, blnGrid)
            strMove = "Down"
        End Select
    End If

    '選択領域が画面から消えたら画面をスクロール
    Dim i As Long
    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '画面分割のない時
        If strMove <> "" Then
            With GetShapeRangeRange(objShapeRange)
                Select Case (strMove)
                Case "Left"
                    i = WorksheetFunction.Max(.Column - 1, 1)
                    If IntersectRange(ActiveWindow.VisibleRange, Columns(i)) Is Nothing Then
                        Call ActiveWindow.SmallScroll(, , , 1)
                    End If
                Case "Right"
                    i = WorksheetFunction.Min(.Column + .Columns.Count, Columns.Count)
                    If IntersectRange(ActiveWindow.VisibleRange, Columns(i)) Is Nothing Then
                        Call ActiveWindow.SmallScroll(, , 1)
                    End If
                Case "Up"
                    i = WorksheetFunction.Max(.Row - 1, 1)
                    If IntersectRange(ActiveWindow.VisibleRange, Rows(i)) Is Nothing Then
                        Call ActiveWindow.SmallScroll(, 1)
                    End If
                Case "Down"
                    i = WorksheetFunction.Min(.Row + .Rows.Count, Rows.Count)
                    If IntersectRange(ActiveWindow.VisibleRange, Rows(i)) Is Nothing Then
                        Call ActiveWindow.SmallScroll(1)
                    End If
                End Select
            End With
            
            '描画の残像を消すため
            Application.ScreenUpdating = True
        End If
    End If
ErrHandle:
End Sub

'*****************************************************************************
'[ 関数名 ]　GetShapeRangeRange
'[ 概  要 ]  ShapeRangeの四方のセル範囲を取得する
'[ 引  数 ]　ShapeRangeオブジェクト
'[ 戻り値 ]　セル範囲
'*****************************************************************************
Private Function GetShapeRangeRange(ByRef objShapeRange As ShapeRange) As Range
    Dim i As Long
    Set GetShapeRangeRange = GetNearlyRange(objShapeRange(1))
    
    For i = 2 To objShapeRange.Count
        Set GetShapeRangeRange = Range(GetShapeRangeRange, GetNearlyRange(objShapeRange(i)))
    Next
End Function

'*****************************************************************************
'[ 関数名 ]　SaveBeforeChange
'[ 概  要 ]　変更前の情報を保存する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SaveBeforeChange()
    If blnChange = False Then
        'アンドゥ用に元の状態を保存する
        Call SaveUndoInfo(E_ShapeSize, objShapeRange)
        Set objShapeRange = GroupSelection(objShapeRange)
        
        If Val(Application.Version) >= 12 Then
            Set objDummy = ActiveSheet.Shapes.AddLine(0, 0, 0, 0)
            Call objDummy.Select
        Else
            Call objShapeRange.Select
        End If

        blnChange = True
        '閉じるボタンを無効にする
        Call EnableMenuItem(GetSystemMenu(FindWindow("ThunderDFrame", Me.Caption), False), SC_CLOSE, (MF_BYCOMMAND Or MF_GRAYED))
    End If

    Call PushUndoInfo
End Sub
    
'*****************************************************************************
'[ 関数名 ]　PushUndoInfo
'[ 概  要 ]　位置情報を保存する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub PushUndoInfo()
    Dim i  As Long
    
    'Undo保存数の最大を超えた時
    If lngUndoCount = UBound(udtShapes) Then
        For i = 2 To UBound(udtShapes)
            udtShapes(i - 1) = udtShapes(i)
        Next
        lngUndoCount = lngUndoCount - 1
    End If
    
    lngUndoCount = lngUndoCount + 1
    With udtShapes(lngUndoCount)
        ReDim .Shapes(1 To objShapeRange.Count)
        For i = 1 To objShapeRange.Count
            .Shapes(i).Height = objShapeRange(i).Height
            .Shapes(i).Width = objShapeRange(i).Width
            .Shapes(i).Top = objShapeRange(i).Top
            .Shapes(i).Left = objShapeRange(i).Left
        Next
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　PopUndoInfo
'[ 概  要 ]　位置情報を復元する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub PopUndoInfo()
    Dim i  As Long
    
    If lngUndoCount = 0 Then
        Exit Sub
    End If
    
    With udtShapes(lngUndoCount)
        For i = 1 To objShapeRange.Count
            objShapeRange(i).Height = .Shapes(i).Height
            objShapeRange(i).Width = .Shapes(i).Width
            objShapeRange(i).Top = .Shapes(i).Top
            objShapeRange(i).Left = .Shapes(i).Left
        Next
    End With

    lngUndoCount = lngUndoCount - 1
End Sub

'*****************************************************************************
'[ 関数名 ]　OnPopupClick
'[ 概  要 ]　ポップアップメニュークリック時
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub OnPopupClick()
On Error GoTo ErrHandle
    '状態の保存
    Call SaveBeforeChange
    
    With objShapeRange
        Select Case CommandBars.ActionControl.Tag
        Case 1
            Call .Align(msoAlignLefts, False)
        Case 2
            Call .Align(msoAlignCenters, False)
        Case 3
            Call .Align(msoAlignRights, False)
        Case 4
            Call .Align(msoAlignTops, False)
        Case 5
            Call .Align(msoAlignMiddles, False)
        Case 6
            Call .Align(msoAlignBottoms, False)
        Case 7
            Call .Distribute(msoDistributeHorizontally, False)
        Case 8
            Call .Distribute(msoDistributeVertically, False)
        Case 9
            Call ConnectShapesH(objShapeRange)
        Case 10
            Call ConnectShapesV(objShapeRange)
        Case 11
            Call DistributeShapesWidth(objShapeRange)
        Case 12
            Call DistributeShapesHeight(objShapeRange)
        Case 13
            Call FitShapesGrid(objShapeRange)
        End Select
    End With
ErrHandle:
'    Call objShapeRange.Select
End Sub

'*****************************************************************************
'[ 関数名 ]　ConnectShapesH
'[ 概  要 ]　図形を左右に連結する
'[ 引  数 ]　objShapes:図形
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ConnectShapesH(ByRef objShapes As ShapeRange)
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
'[ 関数名 ]　ConnectShapesV
'[ 概  要 ]　図形を上下に連結する
'[ 引  数 ]　objShapes:図形
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ConnectShapesV(ByRef objShapes As ShapeRange)
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
'[ 関数名 ]　MoveShapesLR
'[ 概  要 ]　図形を左右に移動する
'[ 引  数 ]　objShapes:図形
'            lngSize:変更サイズ(Pixel)
'            blnFitGrid:枠線にあわせるか
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveShapesLR(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean)
    Dim objShape    As Shape   '図形
    Dim lngLeft      As Long
    Dim lngRight     As Long
    Dim lngNewLeft   As Long
    Dim lngNewRight  As Long
        
    '枠線にあわせるか
    If blnFitGrid = True Then
        '図形の数だけループ
        For Each objShape In objShapes
            lngLeft = Round(objShape.Left / DPIRatio)
            lngRight = Round((objShape.Left + objShape.Width) / DPIRatio)
            
            If lngSize < 0 Then
                lngNewLeft = GetLeftGrid(lngLeft, objShape.TopLeftCell.EntireColumn)
                If lngNewLeft < lngLeft Then
                   objShape.Left = lngNewLeft * DPIRatio
                End If
            Else
                lngNewRight = GetRightGrid(lngRight, objShape.BottomRightCell.EntireColumn)
                If lngNewRight > lngRight Then
                   objShape.Left = (lngLeft + lngNewRight - lngRight) * DPIRatio
                End If
            End If
        Next objShape
    Else
        'ピクセル単位の移動を行う
        Call objShapes.IncrementLeft(lngSize * DPIRatio)
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　MoveShapesUD
'[ 概  要 ]　図形を上下に移動する
'[ 引  数 ]　objShapes:図形
'            lngSize:変更サイズ(Pixel)
'            blnFitGrid:枠線にあわせるか
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub MoveShapesUD(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean)
    Dim objShape     As Shape   '図形
    Dim lngTop       As Long
    Dim lngBottom    As Long
    Dim lngNewTop    As Long
    Dim lngNewBottom As Long
    
    '枠線にあわせるか
    If blnFitGrid = True Then
        '図形の数だけループ
        For Each objShape In objShapes
            lngTop = Round(objShape.Top / DPIRatio)
            lngBottom = Round((objShape.Top + objShape.Height) / DPIRatio)
            
            If lngSize < 0 Then
                lngNewTop = GetTopGrid(lngTop, objShape.TopLeftCell.EntireRow)
                If lngNewTop < lngTop Then
                   objShape.Top = lngNewTop * DPIRatio
                End If
            Else
                lngNewBottom = GetBottomGrid(lngBottom, objShape.BottomRightCell.EntireRow)
                If lngNewBottom > lngBottom Then
                   objShape.Top = (lngTop + lngNewBottom - lngBottom) * DPIRatio
                End If
            End If
        Next objShape
    Else
        'ピクセル単位の移動を行う
        Call objShapes.IncrementTop(lngSize * DPIRatio)
    End If
End Sub

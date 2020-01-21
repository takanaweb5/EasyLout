VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMoveShape 
   Caption         =   "�}�`�̈ړ�"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   OleObjectBlob   =   "frmMoveShape.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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

Private Type TShapes  'Undo���
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
'[�C�x���g]�@UserForm_Initialize
'[ �T  �v ]�@�t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    Dim i         As Long
    
    '�Ăь��ɒʒm����
    blnFormLoad = True
    lngZoom = ActiveWindow.Zoom
    
    Set objShapeRange = Selection.ShapeRange
    
    '�u�z�u/����v�̃|�b�v�A�b�v���j���[�쐬
    Set objTmpBar = CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
    With objTmpBar.Controls
        .Add(, 664).BeginGroup = False '������
        .Add(, 668).BeginGroup = False '���E��������
        .Add(, 665).BeginGroup = False '�E����
        .Add(, 666).BeginGroup = True  '�㑵��
        .Add(, 669).BeginGroup = False '�㉺��������
        .Add(, 667).BeginGroup = False '������
        
        With .Add(, 408)
            .BeginGroup = True
            .Caption = "���E�ɓ��Ԋu�Ő���(&H)"
        End With
        With .Add(, 465)
            .BeginGroup = False
            .Caption = "�㉺�ɓ��Ԋu�Ő���(&V)"
        End With
        With .Add(, 408)
            .BeginGroup = True
            .Caption = "�������ɘA��(&J)"
        End With
        With .Add(, 465)
            .BeginGroup = False
            .Caption = "�c�����ɘA��(&K)"
        End With
        With .Add(, 542)
            .BeginGroup = True
            .FaceId = 2067
            .Caption = "���𑵂���(&W)"
        End With
        With .Add(, 541)
            .BeginGroup = False
            .FaceId = 2068
            .Caption = "�����𑵂���(&E)"
        End With
        With .Add(, 549)
            .BeginGroup = True
            .FaceId = 550
            .Caption = "�O���b�h�ɍ�����(&G)"
        End With
    End With

    For i = 1 To objTmpBar.Controls.Count
        objTmpBar.Controls(i).OnAction = "OnPopupClick"
        objTmpBar.Controls(i).Tag = i
    Next i
    
    Select Case objShapeRange.Count
    Case 1
        objTmpBar.Controls(1).Enabled = False  '������
        objTmpBar.Controls(2).Enabled = False  '���E��������
        objTmpBar.Controls(3).Enabled = False  '�E����
        objTmpBar.Controls(4).Enabled = False  '�㑵��
        objTmpBar.Controls(5).Enabled = False  '�㉺��������
        objTmpBar.Controls(6).Enabled = False  '������
        objTmpBar.Controls(7).Enabled = False  '���E�ɓ��Ԋu�Ő���
        objTmpBar.Controls(8).Enabled = False  '�㉺�ɓ��Ԋu�Ő���
        objTmpBar.Controls(9).Enabled = False  '�������ɘA��
        objTmpBar.Controls(10).Enabled = False '�c�����ɘA��
        objTmpBar.Controls(11).Enabled = False '���𑵂���
        objTmpBar.Controls(12).Enabled = False '�����𑵂���
    Case 2
        objTmpBar.Controls(7).Enabled = False  '���E�ɐ���
        objTmpBar.Controls(8).Enabled = False  '�㉺�ɐ���
    End Select

    '�u�O���b�h�ɂ��킹��v���`�F�b�N���邩����
    If CommandBars.ActionControl.Caption = "�}�`���O���b�h�ɍ�����" Then
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
'[�C�x���g]�@UserForm_Terminate
'[ �T  �v ]�@�t�H�[���A�����[�h��
'*****************************************************************************
Private Sub UserForm_Terminate()
    '�Ăь��ɒʒm����
    blnFormLoad = False
    
    '�u�z�u/����v�̃|�b�v�A�b�v���j���[�폜
    Call objTmpBar.Delete
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_QueryClose
'[ �T  �v ]�@�~�{�^���Ńt�H�[������鎞�A�ύX�����ɖ߂�
'*****************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Call objShapeRange.Select
    
    '�{�������ɖ߂�
    If ActiveWindow.Zoom <> lngZoom Then
        ActiveWindow.Zoom = lngZoom
    End If
    
    '�ύX���Ȃ���΃t�H�[�������
    If blnChange = False Then
        Exit Sub
    End If
        
    '�~�{�^���Ńt�H�[������鎞�A�t�H�[������Ȃ�
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Exit Sub
    End If
    
    If Not (objDummy Is Nothing) Then
        Call objDummy.Delete
    End If
    
    '�O���[�v�������}�`�̉���
    Call UnGroupSelection(objShapeRange).Select
    
    Select Case True
    Case blnOK
        Call SetOnUndo
    Case blnCancel
        Call ExecUndo
    End Select
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdOK_Click
'[ �T  �v ]�@�n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
    blnOK = True
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdCancel_Click
'[ �T  �v ]�@�L�����Z���{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    blnCancel = True
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdAlign_Click
'[ �T  �v ]�@�u�z�u/����v�{�^���N���b�N
'*****************************************************************************
Private Sub cmdAlign_Click()
    Call objTmpBar.ShowPopup
    Call fraKeyCapture.SetFocus
End Sub

'*****************************************************************************
'[�C�x���g]�@KeyDown
'[ �T  �v ]�@�J�[�\���L�[�ňړ����ύX������
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
'[�C�x���g]�@MouseDown
'[ �T  �v ]�@�E�N���b�N�Ń|�b�v�A�b�v���j���[��\������
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
'[�C�x���g]�@UserForm_KeyDown
'[ �T  �v ]�@�J�[�\���L�[�ňړ����ύX������
'*****************************************************************************
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim blnGrid As Boolean
        
    '[Ctrl]+Z ����������Ă��鎞�AUndo���s��
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
    
    'Alt��������Ă���΃X�N���[��
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
                'Excel�̋@�\�𗘗p���āA�}�`��\���ł���ʒu�ɉ�ʂ��X�N���[�������邽��
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
    
    '�ύX�O�̏���ۑ�
    Call SaveBeforeChange
    
    '[Ctrl]Key����������Ă��� or �O���b�h�ɂ��킹�邪�`�F�b�N����Ă��Ȃ�
    If (GetKeyState(vbKeyControl) < 0) Or chkGrid.Value = False Then
        blnGrid = False
    Else
        blnGrid = True
    End If
    
On Error GoTo ErrHandle
    Dim dblSave    As Double
    Dim strMove    As String
    
    If GetKeyState(vbKeyShift) < 0 Then
        '�}�`�̑傫����ύX
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
        '�}�`���ړ�
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

    '�I��̈悪��ʂ�����������ʂ��X�N���[��
    Dim i As Long
    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '��ʕ����̂Ȃ���
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
            
            '�`��̎c������������
            Application.ScreenUpdating = True
        End If
    End If
ErrHandle:
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetShapeRangeRange
'[ �T  �v ]  ShapeRange�̎l���̃Z���͈͂��擾����
'[ ��  �� ]�@ShapeRange�I�u�W�F�N�g
'[ �߂�l ]�@�Z���͈�
'*****************************************************************************
Private Function GetShapeRangeRange(ByRef objShapeRange As ShapeRange) As Range
    Dim i As Long
    Set GetShapeRangeRange = GetNearlyRange(objShapeRange(1))
    
    For i = 2 To objShapeRange.Count
        Set GetShapeRangeRange = Range(GetShapeRangeRange, GetNearlyRange(objShapeRange(i)))
    Next
End Function

'*****************************************************************************
'[ �֐��� ]�@SaveBeforeChange
'[ �T  �v ]�@�ύX�O�̏���ۑ�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub SaveBeforeChange()
    If blnChange = False Then
        '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
        Call SaveUndoInfo(E_ShapeSize, objShapeRange)
        Set objShapeRange = GroupSelection(objShapeRange)
        
        If Val(Application.Version) >= 12 Then
            Set objDummy = ActiveSheet.Shapes.AddLine(0, 0, 0, 0)
            Call objDummy.Select
        Else
            Call objShapeRange.Select
        End If

        blnChange = True
        '����{�^���𖳌��ɂ���
        Call EnableMenuItem(GetSystemMenu(FindWindow("ThunderDFrame", Me.Caption), False), SC_CLOSE, (MF_BYCOMMAND Or MF_GRAYED))
    End If

    Call PushUndoInfo
End Sub
    
'*****************************************************************************
'[ �֐��� ]�@PushUndoInfo
'[ �T  �v ]�@�ʒu����ۑ�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub PushUndoInfo()
    Dim i  As Long
    
    'Undo�ۑ����̍ő�𒴂�����
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
'[ �֐��� ]�@PopUndoInfo
'[ �T  �v ]�@�ʒu���𕜌�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
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
'[ �֐��� ]�@OnPopupClick
'[ �T  �v ]�@�|�b�v�A�b�v���j���[�N���b�N��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub OnPopupClick()
On Error GoTo ErrHandle
    '��Ԃ̕ۑ�
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
'[ �֐��� ]�@ConnectShapesH
'[ �T  �v ]�@�}�`�����E�ɘA������
'[ ��  �� ]�@objShapes:�}�`
'[ �߂�l ]�@�Ȃ�
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

    'Let,Width�̏��Ń\�[�g����
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
'[ �֐��� ]�@ConnectShapesV
'[ �T  �v ]�@�}�`���㉺�ɘA������
'[ ��  �� ]�@objShapes:�}�`
'[ �߂�l ]�@�Ȃ�
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

    'Top,Height�̏��Ń\�[�g����
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
'[ �֐��� ]�@MoveShapesLR
'[ �T  �v ]�@�}�`�����E�Ɉړ�����
'[ ��  �� ]�@objShapes:�}�`
'            lngSize:�ύX�T�C�Y(Pixel)
'            blnFitGrid:�g���ɂ��킹�邩
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveShapesLR(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean)
    Dim objShape    As Shape   '�}�`
    Dim lngLeft      As Long
    Dim lngRight     As Long
    Dim lngNewLeft   As Long
    Dim lngNewRight  As Long
        
    '�g���ɂ��킹�邩
    If blnFitGrid = True Then
        '�}�`�̐��������[�v
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
        '�s�N�Z���P�ʂ̈ړ����s��
        Call objShapes.IncrementLeft(lngSize * DPIRatio)
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@MoveShapesUD
'[ �T  �v ]�@�}�`���㉺�Ɉړ�����
'[ ��  �� ]�@objShapes:�}�`
'            lngSize:�ύX�T�C�Y(Pixel)
'            blnFitGrid:�g���ɂ��킹�邩
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveShapesUD(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean)
    Dim objShape     As Shape   '�}�`
    Dim lngTop       As Long
    Dim lngBottom    As Long
    Dim lngNewTop    As Long
    Dim lngNewBottom As Long
    
    '�g���ɂ��킹�邩
    If blnFitGrid = True Then
        '�}�`�̐��������[�v
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
        '�s�N�Z���P�ʂ̈ړ����s��
        Call objShapes.IncrementTop(lngSize * DPIRatio)
    End If
End Sub

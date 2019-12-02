VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMoveShape 
   Caption         =   "�}�`�̈ړ�"
   ClientHeight    =   2670
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4812
   OleObjectBlob   =   "frmMoveShape.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmMoveShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TPos
    TopLeft  As Long
    Length   As Long
    No       As Long
End Type

Private Type TRect
    Top      As Double
    Height   As Double
    Left     As Double
    Width    As Double
End Type
Private Type TShapes  'Undo���
    Shapes()       As TRect
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
        .Add(, 408).BeginGroup = True  '���E�ɐ���
        .Add(, 465).BeginGroup = False '�㉺�ɐ���

        With .Add(, 408)
            .BeginGroup = False
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
            .BeginGroup = False
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
        objTmpBar.Controls(7).Enabled = False  '���E�ɐ���
        objTmpBar.Controls(8).Enabled = False  '�㉺�ɐ���
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
    
    If GetKeyState(vbKeyShift) < 0 Then
        '�}�`�̑傫����ύX
        If GetKeyState(vbKeyZ) < 0 Then
            Select Case (KeyCode)
            Case vbKeyLeft
                Call ChangeShapesWidth(objShapeRange, 1, blnGrid, True)
            Case vbKeyRight
                Call ChangeShapesWidth(objShapeRange, -1, blnGrid, True)
            Case vbKeyUp
                Call ChangeShapesHeight(objShapeRange, 1, blnGrid, True)
            Case vbKeyDown
                Call ChangeShapesHeight(objShapeRange, -1, blnGrid, True)
            End Select
        Else
            Select Case (KeyCode)
            Case vbKeyLeft
                Call ChangeShapesWidth(objShapeRange, -1, blnGrid, False)
            Case vbKeyRight
                Call ChangeShapesWidth(objShapeRange, 1, blnGrid, False)
            Case vbKeyUp
                Call ChangeShapesHeight(objShapeRange, -1, blnGrid, False)
            Case vbKeyDown
                Call ChangeShapesHeight(objShapeRange, 1, blnGrid, False)
            End Select
        End If
    Else
        '�}�`���ړ�
        Select Case (KeyCode)
        Case vbKeyLeft
            Call MoveShapesLR(objShapeRange, -1, blnGrid)
        Case vbKeyRight
            Call MoveShapesLR(objShapeRange, 1, blnGrid)
        Case vbKeyUp
            Call MoveShapesUD(objShapeRange, -1, blnGrid)
        Case vbKeyDown
            Call MoveShapesUD(objShapeRange, 1, blnGrid)
        End Select
    End If
ErrHandle:
End Sub

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
    
    ReDim udtPosArray(1 To objShapes.Count) As TPos
    For i = 1 To objShapes.Count
        With udtPosArray(i)
            .No = i
            .TopLeft = objShapes(i).Left / 0.75
            .Length = objShapes(i).Width / 0.75
        End With
    Next

    Call SortPosArray(udtPosArray())

    Dim lngLeft    As Long
    lngLeft = udtPosArray(1).TopLeft
    
    For i = 2 To objShapes.Count
        If udtPosArray(i).TopLeft > udtPosArray(i - 1).TopLeft Then
            lngLeft = lngLeft + udtPosArray(i - 1).Length
        End If
        
        objShapes(udtPosArray(i).No).Left = lngLeft * 0.75
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
    
    ReDim udtPosArray(1 To objShapes.Count) As TPos
    For i = 1 To objShapes.Count
        With udtPosArray(i)
            .No = i
            .TopLeft = objShapes(i).Top / 0.75
            .Length = objShapes(i).Height / 0.75
        End With
    Next

    Call SortPosArray(udtPosArray())

    Dim lngLeft    As Long
    lngLeft = udtPosArray(1).TopLeft
    
    For i = 2 To objShapes.Count
        If udtPosArray(i).TopLeft > udtPosArray(i - 1).TopLeft Then
            lngLeft = lngLeft + udtPosArray(i - 1).Length
        End If
        
        objShapes(udtPosArray(i).No).Top = lngLeft * 0.75
    Next
End Sub

'*****************************************************************************
'[ �֐��� ]�@SortPosArray
'[ �T  �v ]�@PosArray�z���Worksheet���g���ă\�[�g����
'[ ��  �� ]�@PosArray�z��
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub SortPosArray(ByRef udtPosArray() As TPos)
On Error GoTo ErrHandle
    Dim objWorksheet As Worksheet
    Dim i As Long
    
    Set objWorksheet = ThisWorkbook.Worksheets("Workarea1")
    Call DeleteSheet(objWorksheet)
    For i = 1 To UBound(udtPosArray)
        With udtPosArray(i)
            objWorksheet.Cells(i, 1) = .TopLeft
            objWorksheet.Cells(i, 2) = .Length
            objWorksheet.Cells(i, 3) = .No
        End With
    Next
    
    With objWorksheet.Cells(1, 1).CurrentRegion
        'Key1:TopLeft�AKey2:Length �Ń\�[�g����
        Call .Sort(Key1:=.Columns(1), Order1:=xlAscending, _
                   Key2:=.Columns(2), Order2:=xlAscending, _
                   Header:=xlNo, OrderCustom:=1, Orientation:=xlTopToBottom)
    End With
    
    For i = 1 To UBound(udtPosArray)
        With udtPosArray(i)
            .TopLeft = objWorksheet.Cells(i, 1)
            .Length = objWorksheet.Cells(i, 2)
            .No = objWorksheet.Cells(i, 3)
        End With
    Next
ErrHandle:
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
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
            lngLeft = Round(objShape.Left / 0.75)
            lngRight = Round((objShape.Left + objShape.Width) / 0.75)
            
            If lngSize < 0 Then
                lngNewLeft = GetLeftGrid(lngLeft, objShape.TopLeftCell.EntireColumn)
                If lngNewLeft < lngLeft Then
                   objShape.Left = lngNewLeft * 0.75
                End If
            Else
                lngNewRight = GetRightGrid(lngRight, objShape.BottomRightCell.EntireColumn)
                If lngNewRight > lngRight Then
                   objShape.Left = (lngLeft + lngNewRight - lngRight) * 0.75
                End If
            End If
        Next objShape
    Else
        '�s�N�Z���P�ʂ̈ړ����s��
        Call objShapes.IncrementLeft(lngSize * 0.75)
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
            lngTop = Round(objShape.Top / 0.75)
            lngBottom = Round((objShape.Top + objShape.Height) / 0.75)
            
            If lngSize < 0 Then
                lngNewTop = GetTopGrid(lngTop, objShape.TopLeftCell.EntireRow)
                If lngNewTop < lngTop Then
                   objShape.Top = lngNewTop * 0.75
                End If
            Else
                lngNewBottom = GetBottomGrid(lngBottom, objShape.BottomRightCell.EntireRow)
                If lngNewBottom > lngBottom Then
                   objShape.Top = (lngTop + lngNewBottom - lngBottom) * 0.75
                End If
            End If
        Next objShape
    Else
        '�s�N�Z���P�ʂ̈ړ����s��
        Call objShapes.IncrementTop(lngSize * 0.75)
    End If
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMoveCell 
   Caption         =   "�̈�̑���"
   ClientHeight    =   2970
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5124
   OleObjectBlob   =   "frmMoveCell.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
Private strFromRange As String '���̗̈�
Private strToRange   As String '�ړ���
Private objFromSheet As Worksheet '���̗̈�
Private objTextbox   As Shape  '�ړ�������o�I�ɕ\������
Private lngDisplayObjects As Long
Private lngZoom      As Long

'*****************************************************************************
'[�T�v] �����ݒ����ݒ�
'[����] enmType:��ƃ^�C�v
'       objFromRange:�ړ�(�R�s�[��)�̗̈�
'       objToRange:�I�𒆂̗̈�
'[�ߒl] �Ȃ�
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
    
    '�ړ�(�R�s�[)��̃V�[�g��Activate�ɂ���
    Call objToRange.Worksheet.Activate
    
    '���Ɛ�̃V�[�g������̎��A�R�s�[���̃Z����I�����邱�Ƃŉ��ł���悤�ɂ���
     Call Range(strFromRange).Select
'    If blnSameSheet = True Then
'        Call Range(strFromRange).Select
'    Else
'        Call Range(strToRange).Select
'    End If
    
    '�e�L�X�g�{�b�N�X�쐬
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
    '�e�L�X�g�{�b�N�X�̔w�i��ύX
    With objTextbox.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.SchemeColor = 65
        .Transparency = 0.12  '�w�i�𓧂�������
    End With
    
    '�e�L�X�g�{�b�N�X�̌r����ύX
    With objTextbox.Line
        .Weight = 2#
        .Style = msoLineSingle
        .Transparency = 0#
        .Visible = msoTrue
        .ForeColor.SchemeColor = 64
        .BackColor.Rgb = Rgb(255, 255, 255)
        .Pattern = msoPattern50Percent
    End With
    
    '�e�L�X�g�{�b�N�X�ɑI���Z���̃A�h���X��\��
    With objTextbox.TextFrame.Characters
        .Text = Replace$(strToRange, "$", "")
        lblCellAddress.Caption = " " & .Text
    End With
    
    '�I��̈悪��ʂ�������Ă��鎞
    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '��ʕ����̂Ȃ���
        If IntersectRange(ActiveWindow.VisibleRange, objToRange) Is Nothing Then
            With objToRange
                Call ActiveWindow.ScrollIntoView(.Left / DPIRatio, .Top / DPIRatio, .Width / DPIRatio, .Height / DPIRatio)
            End With
        End If
    End If
    
'    '�`�F�b�N�{�^���̐���
'    If blnSameSheet = False Then
'        chkExchange.Enabled = False
'        chkCopy.Enabled = False
'        chkExchange.Enabled = False
'        chkCutInsert.Enabled = False
'    End If
    
    '�`�F�b�N�{�b�N�X��ݒ�
    Call ChangeMode
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetToRange
'[ �T  �v ]  �I�𒆂̗̈悩��A�ړ���̗̈�̏����\���G���A���v�Z����
'[ ��  �� ]  objFromRange:�ړ�(�R�s�[��)�̗̈�
'            objToRange:�I�𒆂̗̈�
'[ �߂�l ]  �ړ���̗̈�̏����\���G���A
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
'[�C�x���g]�@UserForm_Initialize
'[ �T  �v ]�@�t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    '�Ăь��ɒʒm����
    blnFormLoad = True
    lngZoom = ActiveWindow.Zoom
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_Terminate
'[ �T  �v ]�@�t�H�[���A�����[�h��
'*****************************************************************************
Private Sub UserForm_Terminate()
    '�Ăь��ɒʒm����
    blnFormLoad = False
    
    If Not (objTextbox Is Nothing) Then
        Call objTextbox.Delete
    End If
    
    ActiveWorkbook.DisplayDrawingObjects = lngDisplayObjects
    ActiveWindow.Zoom = lngZoom
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdOK_Click
'[ �T  �v ]�@�n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrHandle
    Dim blnCopyObjectsWithCells  As Boolean
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells
    Application.CopyObjectsWithCells = False '�Ăь��ŕ������邽�ߓ����W���[���ł͕������Ȃ�
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False '�R�����g�����鎞�x�����o�鎞������
    
    Call objTextbox.Delete
    Set objTextbox = Nothing
    
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
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
    
    '�Z�����ړ��E�R�s�[����
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
    
    '�}�`���ړ��E�R�s�[����
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
'[ �֐��� ]�@CopyCell
'[ �T  �v ]  �̈���R�s�[����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub CopyCell()
    Dim objWkRange As Range
    
    '�̈�����[�N�V�[�g�ɃR�s�[����
    Set objWkRange = CopyToWkSheet(1, FromRange, ToRange)
    
    '���[�N�V�[�g�ŗ̈�̃T�C�Y��ύX����
    Set objWkRange = ReSizeArea(objWkRange, ToRange.Rows.Count, ToRange.Columns.Count)
    
    If chkOnlyValue.Value = True Then
        '�l�̂݃R�s�[
        Call CopyOnlyValue(objWkRange, ToRange)
    Else
        Call objWkRange.Copy(ToRange)
    End If
End Sub
    
'*****************************************************************************
'[ �֐��� ]�@MoveCell
'[ �T  �v ]  �̈���ړ�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveCell()
    Dim objWkRange As Range
    
    '�̈�����[�N�V�[�g�ɃR�s�[����
    Set objWkRange = CopyToWkSheet(1, FromRange, ToRange)
    
    '���[�N�V�[�g�ŗ̈�̃T�C�Y��ύX����
    Set objWkRange = ReSizeArea(objWkRange, ToRange.Rows.Count, ToRange.Columns.Count)
    
    '���̗̈���N���A
    With FromRange
        Call .Clear
        If .Worksheet.Cells(Rows.Count - 2, Columns.Count - 2).MergeCells = False Then
            '�V�[�g��̕W���I�ȏ����ɐݒ�
            Call .Worksheet.Cells(Rows.Count - 2, Columns.Count - 2).Copy(.Cells)
            Call .ClearContents
        End If
    End With
    
    Call objWkRange.Copy(ToRange)
End Sub

'*****************************************************************************
'[ �֐��� ]�@ExchangeCell
'[ �T  �v ]  �̈���������
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ExchangeCell()
    Dim objWkRange(1 To 2) As Range
    
    '�WorkArea��V�[�g���g�p���ĕҏW����
    Set objWkRange(1) = CopyToWkSheet(1, FromRange, ToRange)
    Set objWkRange(2) = CopyToWkSheet(2, ToRange, FromRange)
    
    '���[�N�V�[�g�ŗ̈�̃T�C�Y��ύX����
    Set objWkRange(1) = ReSizeArea(objWkRange(1), ToRange.Rows.Count, ToRange.Columns.Count)
    Set objWkRange(2) = ReSizeArea(objWkRange(2), FromRange.Rows.Count, FromRange.Columns.Count)
    
    If chkOnlyValue.Value = True Then
        '�l�̂݃R�s�[
        Call CopyOnlyValue(objWkRange(1), ToRange)
        Call CopyOnlyValue(objWkRange(2), FromRange)
    Else
        Call objWkRange(1).Copy(ToRange)
        Call objWkRange(2).Copy(FromRange)
    End If
End Sub
    
'*****************************************************************************
'[ �֐��� ]�@CutInsertCell
'[ �T  �v ]  �؂������Z���̑}��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub CutInsertCell()
    Dim objActionRange As Range
    Dim strActionRange As String
    Dim objWkRange     As Range
    
    Set objActionRange = GetActionRange()
    strActionRange = objActionRange.Address

    '���[�N�̃V�[�g�ɍ�ƈ���R�s�[
    Set objWkRange = CopyToWkSheet(1, objActionRange, objActionRange)

    '�؂������Z���̑}��
    With objWkRange.Worksheet
        Call .Range(strFromRange).Cut
        Call .Range(strToRange).Insert
    End With
    
    With objWkRange.Worksheet
        Call .Range(strActionRange).Copy(Range(strActionRange))
    End With
End Sub

'*****************************************************************************
'[ �֐��� ]�@CopyOnlyValue
'[ �T  �v ]  �l�̂݃R�s�[
'[ ��  �� ]�@�R�s�[���A�R�s�[��
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub CopyOnlyValue(ByRef objFromRange As Range, ByRef objToRange As Range)
    '�Z�������̃R�s�[
    Call objToRange.UnMerge
    Call CopyMergeRange(objFromRange, objToRange)
    
    '�l�̂݃R�s�[
    Call objFromRange.Copy
    Call objToRange.PasteSpecial(xlPasteFormulas, xlNone, False, False)
End Sub

'*****************************************************************************
'[ �֐��� ]�@CopyMergeRange
'[ �T  �v ]  �̈�̌�����Ԃ݂̂𕡎ʂ���
'[ ��  �� ]�@�R�s�[���A�R�s�[��
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub CopyMergeRange(ByRef objFromRange As Range, ByRef objToRange As Range)
    Dim objRange As Range
    Dim i As Long
    Dim j As Long
    
    '�s�̐��������[�v
    For i = 1 To objFromRange.Rows.Count
        '��̐��������[�v
        For j = 1 To objFromRange.Columns.Count
            With objFromRange(i, j).MergeArea
                If .Count > 1 Then
                    '�����Z���̍���̃Z���Ȃ�
                    If objFromRange(i, j).Address = .Cells(1, 1).Address Then
                        '���ʐ�̃Z��������
                        Call objToRange(i, j).Resize(.Rows.Count, .Columns.Count).Merge
                    End If
                End If
            End With
        Next j
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ]�@MoveShape
'[ �T  �v ]  �}�`���ړ��܂��̓R�s�[����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveShape()
    Dim objShapes As ShapeRange
    
    '��]���Ă���}�`���O���[�v������
    Call GroupSelection(GetSheeetShapeRange(ActiveSheet))
    
    '�ړ����̗̈�̐}�`���擾
    Set objShapes = SelectShapes(FromRange)
    
    '�ړ������ړ���
    If Not (objShapes Is Nothing) Then
        Call MoveShapes(objShapes, FromRange, ToRange, (enmModeType = E_Copy))
    End If
    
    '�O���[�v�������}�`�̉���
    Call UnGroupSelection(GetSheeetShapeRange(ActiveSheet))
End Sub

'*****************************************************************************
'[ �֐��� ]�@ExchangeShape
'[ �T  �v ]  �}�`����������
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ExchangeShape()
    Dim objShapes(1 To 2) As ShapeRange '1:�ړ����̐}�`�A2:�ړ���̐}�`
    
    '��]���Ă���}�`���O���[�v������
    Call GroupSelection(GetSheeetShapeRange(ActiveSheet))
    
    '�ړ����ƈړ���̗̈�̐}�`���擾
    Set objShapes(1) = SelectShapes(FromRange)
    Set objShapes(2) = SelectShapes(ToRange)
    
    '�ړ������ړ���
    If Not (objShapes(1) Is Nothing) Then
        Call MoveShapes(objShapes(1), FromRange, ToRange)
    End If
    
    '�ړ������ړ���
    If Not (objShapes(2) Is Nothing) Then
        Call MoveShapes(objShapes(2), ToRange, FromRange)
    End If
    
    '�O���[�v�������}�`�̉���
    Call UnGroupSelection(GetSheeetShapeRange(ActiveSheet))
End Sub

'*****************************************************************************
'[ �֐��� ]�@CutInsertShape
'[ �T  �v ]  �؂������Z���̑}�����̐}�`����������
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub CutInsertShape()
    Dim objShapes(1 To 2) As ShapeRange '1:�ړ����̐}�`�A2:�ȊO�̐}�`
    
    '��]���Ă���}�`���O���[�v������
    Call GroupSelection(GetSheeetShapeRange(ActiveSheet))
    
    '�ړ����ƈȊO�̗̈�̐}�`���擾
    Set objShapes(1) = SelectShapes(FromRange)
    Set objShapes(2) = SelectShapes(MinusRange(GetActionRange(), FromRange))
    
    '�ړ������ړ���
    If Not (objShapes(1) Is Nothing) Then
        Call MoveShapes(objShapes(1), FromRange, GetSlideRange(True))
    End If
    
    '�ړ������ړ���
    If Not (objShapes(2) Is Nothing) Then
        Call MoveShapes(objShapes(2), MinusRange(GetActionRange(), FromRange), GetSlideRange(False))
    End If
    
    '�O���[�v�������}�`�̉���
    Call UnGroupSelection(GetSheeetShapeRange(ActiveSheet))
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetSlideRange
'[ �T�@�v ]�@�؂������Z���̑}�����A�}�`���X���C�h������̈���擾����
'[ ���@�� ]�@True:�ړ����̃X���C�h�̈�AFalse:�ȊO�̃X���C�h�̈�
'[ �߂�l ]�@�}�`���X���C�h������̈�
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
'[ �֐��� ]�@SelectShapes
'[ �T  �v ]  �I������Ă���̈�̒��Ɋ܂܂��}�`���擾
'[ ��  �� ]�@�Ώۗ̈�
'[ �߂�l ]�@�̈�Ɋ܂܂��}�`
'*****************************************************************************
Private Function SelectShapes(ByRef objRange As Range) As ShapeRange
    Dim i           As Long
    Dim objShape    As Shape
    Dim udtShape    As TRect
    Dim udtRange    As TRect
    ReDim lngIDArray(1 To ActiveSheet.Shapes.Count) As Variant
    
    '�I�����ꂽ�G���A�Ɋ܂܂��}�`���擾
    For Each objShape In ActiveSheet.Shapes
        'Excel2007�Ή�DPIRatio�̔{���ɕ␳
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
'[ �֐��� ]�@MoveShapes
'[ �T  �v ]  �}�`���ړ�����
'[ ��  �� ]�@�ړ����̗̈�̐}�`�A�ړ����̃Z���̈�A�ړ���̃Z���̈�ATrue:�R�s�[����
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveShapes(ByRef objShapes As ShapeRange, ByRef objFromRange As Range, _
                       ByRef objToRange As Range, Optional ByVal blnCopy As Boolean = False)
    Dim objShape    As Shape
    Dim i           As Long
    ReDim udtShapes(1 To objShapes.Count) As TShape
    
    '�I�����ꂽ�G���A�Ɋ܂܂��}�`���擾
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
                '���ʂ���
                Call SetRect(GetShapeFromID(udtShapes(i).ID).Duplicate, objToRange, udtShapes(i))
            Else
                '�ړ�����
                Call SetRect(GetShapeFromID(udtShapes(i).ID), objToRange, udtShapes(i))
            End If
'        End If
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetRect
'[ �T  �v ]  �}�`�̈ʒu���ړ���̗̈�ɐݒ肷��
'[ ��  �� ]�@�Ώۂ̐}�`�A�ړ���̗̈�A�ʒu���
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub SetRect(ByRef objShape As Shape, ByRef objRange As Range, ByRef udtShape As TShape)
    '����̈ʒu��ݒ肷��
'    If udtShape.Placement <> xlFreeFloating Then
        objShape.Top = objRange.Rows(udtShape.TopRow).Top + udtShape.Top
        objShape.Left = objRange.Columns(udtShape.LeftColumn).Left + udtShape.Left
'    End If

    '���ƍ�����ݒ肷��
'    If udtShape.Placement = xlMoveAndSize Then
        objShape.Height = objRange.Rows(udtShape.BottomRow).Top + udtShape.Height - objShape.Top
        objShape.Width = objRange.Columns(udtShape.RightColumn).Left + udtShape.Width - objShape.Left
'    End If
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdCancel_Click
'[ �T  �v ]�@�L�����Z���{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdHelp_Click
'[ �T  �v ]�@�w���v�{�^��������
'*****************************************************************************
Private Sub cmdHelp_Click()
    Call OpenHelpPage("MoveCell.htm")
End Sub

'*****************************************************************************
'[�C�x���g]�@�`�F�b�N�{�^��_Change
'[ �T  �v ]�@�`�F�b�N�{�^���ύX��
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
'[ �֐��� ]�@ChangeMode
'[ �T  �v ]  ��ړ�����R�s�[������ւ���̃��[�h��ύX����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@True:�L�����Z����
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
        lblTitle.Caption = "�ړ���"
    Case E_Copy
        lblTitle.Caption = "�R�s�[��"
    Case E_Exchange
        lblTitle.Caption = "���ւ���"
    Case E_CutInsert
        lblTitle.Caption = "�}����"
    End Select

    '�e�L�X�g�{�b�N�X��ҏW
    Call EditTextbox
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
'[�C�x���g]�@UserForm_KeyDown
'[ �T  �v ]�@�J�[�\���L�[�ňړ����ύX������
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
    
    '�I��̈�̎l���̈ʒu��Ҕ�
    With Range(strToRange)
        lngTop = .Row
        lngBottom = .Row + .Rows.Count - 1
        lngLeft = .Column
        lngRight = .Column + .Columns.Count - 1
    End With
    
    Select Case (Shift)
    Case 0
        '�I��̈���ړ�
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
        '�I��̈�̑傫����ύX
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
    
    '�`�F�b�N
    If (lngLeft <= lngRight And lngTop <= lngBottom) And _
       (1 <= lngLeft And lngRight <= Columns.Count) And _
       (1 <= lngTop And lngBottom <= Rows.Count) Then
        Set objToRange = Range(Cells(lngTop, lngLeft), Cells(lngBottom, lngRight))
    Else
        Exit Sub
    End If
    
    '�V�����I��̈������Range
    strToRange = objToRange.Address
    
    '�e�L�X�g�{�b�N�X��ҏW
    Call EditTextbox
    
    '�e�L�X�g�{�b�N�X���ړ�
    With Range(strToRange)
        objTextbox.Left = .Left
        objTextbox.Top = .Top
        objTextbox.Width = .Width
        objTextbox.Height = .Height
    End With
    
    '�ړ���(�R�s�[��)�̕ҏW
    lblCellAddress.Caption = " " & Replace$(strToRange, "$", "")

    '�I��̈悪��ʂ�����������ʂ��X�N���[��
    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '��ʕ����̂Ȃ���
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
'[ �֐��� ]�@InsertOrDeleteCell
'[ �T  �v ]  �Z����}���܂��͍폜����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@True:�L�����Z����
'*****************************************************************************
Private Function InsertOrDeleteCell(ByVal lngKeyCode As Long) As Boolean
    Dim strMsg      As String
    Dim lngSelect   As Long
    
    Select Case (lngKeyCode)
    Case vbKeyDelete
        strMsg = "��������폜���܂����H" & vbCrLf
        strMsg = strMsg & "�@�u �͂� �v���� �������ɃV�t�g����" & vbCrLf
        strMsg = strMsg & "�@�u�������v���� ������ɃV�t�g����"
    Case vbKeyInsert
        strMsg = "������ɑ}�����܂����H" & vbCrLf
        strMsg = strMsg & "�@�u �͂� �v���� �E�����ɃV�t�g����" & vbCrLf
        strMsg = strMsg & "�@�u�������v���� �������ɃV�t�g����"
    End Select
    
    Select Case True
    Case strFromRange = strToRange
        lngSelect = MsgBox(strMsg, vbYesNoCancel + vbDefaultButton1, "�I�����ĉ�����")
    Case FromRange.EntireRow.Address = ToRange.EntireRow.Address
        lngSelect = vbYes
    Case FromRange.EntireColumn.Address = ToRange.EntireColumn.Address
        lngSelect = vbNo
    Case Else
        lngSelect = MsgBox(strMsg, vbYesNoCancel + vbDefaultButton1, "�I�����ĉ�����")
    End Select
    If lngSelect = vbCancel Then
        InsertOrDeleteCell = True
        Exit Function
    End If
    
    If CheckInsertOrDelete(lngSelect = vbYes) = False Then
        Call MsgBox("�������ꂽ�Z���̈ꕔ��ύX���邱�Ƃ͂ł��܂���", vbExclamation)
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    Call objTextbox.Delete
    Set objTextbox = Nothing

On Error GoTo ErrHandle
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
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
'[ �֐��� ]�@CheckInsertOrDelete
'[ �T  �v ]  Ins/Del�L�[�������̌����Z���̃`�F�b�N
'[ ��  �� ]�@True�F������̃V�t�g�AFalse�F�s�����̃V�t�g
'[ �߂�l ]�@False:�G���[����
'*****************************************************************************
Private Function CheckInsertOrDelete(ByVal blnSelect As Boolean) As Boolean
    Dim objWkRange   As Range
    Dim objActRange  As Range
    Dim objActArea   As Range
    Dim objChkRange  As Range

    Set objWkRange = ArrangeRange(ToRange)
    If blnSelect Then
        '������̃V�t�g�̎��A�I��̈悪�s�����Ɍ������͂ݏo���Ă��Ȃ�������
        Set objChkRange = MinusRange(objWkRange, ToRange.EntireRow)
    Else
        '�s�����̃V�t�g�̎��A�I��̈悪������Ɍ������͂ݏo���Ă��Ȃ�������
        Set objChkRange = MinusRange(objWkRange, ToRange.EntireColumn)
    End If
    If Not (objChkRange Is Nothing) Then
        CheckInsertOrDelete = False
        Exit Function
    End If
        
    If blnSelect Then
        '������̃V�t�g�̎��A�I��̈悩��Ō�̗�͈̔͂��擾
        Set objActRange = Range(ToRange, Cells(objWkRange.Row, Columns.Count))
        Set objActArea = objActRange.EntireColumn
    Else
        '�s�����̃V�t�g�̎��A�I��̈悩��Ō�̍s�͈̔͂��擾
        Set objActRange = Range(ToRange, Cells(Rows.Count, objWkRange.Column))
        Set objActArea = objActRange.EntireRow
    End If
    '�V�[�g�̍Ō�̍s�܂��͗�͈̔͂Ť�����Z�����͂ݏo���Ă��Ȃ�������
    Set objChkRange = IntersectRange(MinusRange(ArrangeRange(objActRange), objActRange), objActArea)
    CheckInsertOrDelete = (objChkRange Is Nothing)
End Function
    
'*****************************************************************************
'[ �֐��� ]�@EditTextbox
'[ �T  �v ]  �e�L�X�g�{�b�N�X��ҏW����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub EditTextbox()
On Error GoTo ErrHandle
'    If blnSameSheet Then
    Select Case enmModeType
    Case E_Copy, E_Exchange
        If Not (IntersectRange(FromRange, ToRange) Is Nothing) Then
            Select Case enmModeType
            Case E_Copy
                Call Err.Raise(C_CheckErrMsg, , "���̗̈�ɂ̓R�s�[�ł��܂���")
            Case E_Exchange
                Call Err.Raise(C_CheckErrMsg, , "���̗̈�Ƃ͓��ւ��ł��܂���")
            End Select
        End If
    End Select
'    End If

    Select Case enmModeType
    Case E_Move, E_Copy, E_Exchange
        If CheckBorder() = True Then
            Call Err.Raise(C_CheckErrMsg, , "�������ꂽ�Z���̈ꕔ��ύX���邱�Ƃ͂ł��܂���")
        End If
        If FromRange.Columns.Count = Columns.Count And _
           ToRange.Columns.Count <> Columns.Count Then
            Call Err.Raise(C_CheckErrMsg, , "���ׂĂ̗�̑I�����͗�̕���ύX�ł��܂���")
        End If
    Case E_CutInsert
        Call CheckCutInsertRange
    End Select
    
    With objTextbox.TextFrame.Characters
        .Text = Replace$(strToRange, "$", "")
        .Font.ColorIndex = 0
        cmdOK.Enabled = True  'OK�{�^�����g�p�ɂ���
    End With
Exit Sub
ErrHandle:
    If Err.Number = C_CheckErrMsg Then
        With objTextbox.TextFrame.Characters
            .Text = Err.Description
            .Font.ColorIndex = 3
            cmdOK.Enabled = False 'OK�{�^�����g�p�s�ɂ���
        End With
    Else
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@CheckCutInsertRange
'[ �T  �v ]  �؂������̈�̑}���̂Ƃ��̗̈�̃`�F�b�N
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ��i�G���[�̎��͗�O�j
'*****************************************************************************
Private Sub CheckCutInsertRange()
    Dim enmDirection   As EDirection
    Dim objActionRange As Range
    
    If FromRange.Columns.Count <> ToRange.Columns.Count Or _
       FromRange.Rows.Count <> ToRange.Rows.Count Then
        Call Err.Raise(C_CheckErrMsg, , "�}�����͑傫����ύX�ł��܂���")
    End If
    
    enmDirection = GetDirection()
    Select Case enmDirection
    Case E_Non
        Call Err.Raise(C_CheckErrMsg, , "�}����Ɉړ����ĉ�����")
    Case E_ERR
        Call Err.Raise(C_CheckErrMsg, , "�^��E�^���E�^���ȊO�ɂ͑}���ł��܂���")
    End Select
        
    Set objActionRange = GetActionRange()
    If objActionRange.Address = strFromRange Then
        Call Err.Raise(C_CheckErrMsg, , "���̗̈�ɂ͑}���ł��܂���")
    End If
    
    If IsBorderMerged(objActionRange) Then
        Call Err.Raise(C_CheckErrMsg, , "�������ꂽ�Z���̈ꕔ��ύX���邱�Ƃ͂ł��܂���")
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@CheckBorder
'[ �T  �v ]  �ړ���̎l���������Z�����܂����ł��邩�ǂ���
'            �ړ��̎��́A���̗̈�̌����Z���͑ΏۂƂ��Ȃ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@True:�����Z�����܂����ł���AFalse:���Ȃ�
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
'[ �֐��� ]�@CopyToWkSheet
'[ �T  �v ]  ���[�N�V�[�g�ɗ̈���R�s�[����
'[ ��  �� ]�@Workarea1���g�p���邩Workarea2���g�p���邩����
'            �R�s�[���̃A�h���X�A�R�s�[��̃A�h���X
'[ �߂�l ]  ���[�N�V�[�g�̃R�s�[���ꂽRange(�T�C�Y�̓R�s�[���̃T�C�Y)
'*****************************************************************************
Private Function CopyToWkSheet(ByVal bytNo As Byte, ByRef objFromRange As Range, ByRef objToRange As Range) As Range
    Dim objWorksheet As Worksheet
    
    Set objWorksheet = ThisWorkbook.Worksheets("Workarea" & bytNo)
    Call DeleteSheet(objWorksheet)
    
    '�uWorkarea�v�V�[�g�ɑI��̈�𕡎ʂ���
    With objFromRange
        Set CopyToWkSheet = objWorksheet.Range(objToRange.Address).Resize(.Rows.Count, .Columns.Count)
        Call .Copy(CopyToWkSheet)
    End With
End Function
    
'*****************************************************************************
'[ �֐��� ]�@ReSizeArea
'[ �T  �v ]  ���[�N�G���A���g�p���̈�̃T�C�Y��ύX����
'[ ��  �� ]�@objRange�F�ύX����̈�
'            lngNewRow�AlngNewCol�F�V�����T�C�Y
'[ �߂�l ]  �T�C�Y��ύX����Range
'*****************************************************************************
Private Function ReSizeArea(ByRef objRange As Range, ByVal lngNewRow As Long, ByVal lngNewCol As Long) As Range
    Dim objWkRange   As Range
    Dim lngDiff      As Long
    Dim lngOldRow    As Long
    Dim lngOldCol    As Long
    
    '�ύX�O�̃T�C�Y���擾
    With objRange
        lngOldRow = .Rows.Count
        lngOldCol = .Columns.Count
    End With
    
    '�s�̃T�C�Y�ύX
    lngDiff = lngNewRow - lngOldRow
    If lngDiff <> 0 Then
        Call ChangeRangeRowSize(objRange, lngOldRow, lngOldCol, lngDiff)
    End If
    
    '��̃T�C�Y�ύX
    lngDiff = lngNewCol - lngOldCol
    If lngDiff <> 0 Then
        Call ChangeRangeColSize(objRange, lngOldRow, lngOldCol, lngDiff)
    End If
    
    Set ReSizeArea = objRange.Resize(lngNewRow, lngNewCol)
End Function


'*****************************************************************************
'[ �֐��� ]�@ChangeRangeRowSize
'[ �T  �v ]  �̈�̃T�C�Y��ύX����
'[ ��  �� ]�@objWkRange:�T�C�Y��ύX����̈椌��̍s���Ɨ񐔁AlngDiff:�s�̕ύX�T�C�Y
'[ �߂�l ]  �Ȃ�
'*****************************************************************************
Private Sub ChangeRangeRowSize(ByRef objWkRange As Range, ByVal lngRowCount As Long, ByVal lngColCount As Long, ByVal lngDiff As Long)
    Dim i            As Long
    Dim udtBorder    As TBorder '�r���̎��
    Dim objWorksheet As Worksheet
    
    Set objWorksheet = objWkRange.Worksheet
    
    Select Case (lngDiff)
    Case Is < 0 '�k��
        '���[�̌r����ۑ�
        With objWkRange.Borders
            udtBorder = GetBorder(.Item(xlEdgeBottom))
        End With
        '�폜
        With objWkRange
            Call objWorksheet.Range(.Rows(lngRowCount + lngDiff + 1), _
                                    .Rows(lngRowCount)).Delete(xlShiftUp)
        End With
    Case Is > 0 '�g��
        '�}��
        With objWkRange.Rows
            Call .Item(.Count).Copy(objWorksheet.Range(.Item(lngRowCount + 1), .Item(lngRowCount + lngDiff)))
            Call objWorksheet.Range(.Item(lngRowCount + 1), .Item(lngRowCount + lngDiff)).ClearContents
        End With
        
        '��̐��������[�v
        For i = 1 To lngColCount
            '�Z��������
            With objWkRange.Rows(lngRowCount)
                Call objWorksheet.Range(.Cells(1, i), .Cells(lngDiff + 1, i).MergeArea).Merge
            End With
        Next i
    End Select
    
    Set objWkRange = objWkRange.Resize(lngRowCount + lngDiff)
    
    '�폜�̎����[�̌r�����Đݒ�
    If lngDiff < 0 Then
        '���[�̌r�����Đݒ�
        With objWkRange.Borders
            Call SetBorder(udtBorder, .Item(xlEdgeBottom))
        End With
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@ChangeRangeColSize
'[ �T  �v ]  �̈�̃T�C�Y��ύX����
'[ ��  �� ]�@objWkRange:�T�C�Y��ύX����̈椌��̍s���Ɨ񐔁AlngDiff:��̕ύX�T�C�Y
'[ �߂�l ]  �Ȃ�
'*****************************************************************************
Private Sub ChangeRangeColSize(ByRef objWkRange As Range, ByVal lngRowCount As Long, ByVal lngColCount As Long, ByVal lngDiff As Long)
    Dim i            As Long
    Dim udtBorder    As TBorder '�r���̎��
    Dim objWorksheet As Worksheet
    
    Set objWorksheet = objWkRange.Worksheet
    
    Select Case (lngDiff)
    Case Is < 0 '�k��
        '�E�[�̌r����ۑ�
        With objWkRange.Borders
            udtBorder = GetBorder(.Item(xlEdgeRight))
        End With
        '�폜
        With objWkRange
            Call objWorksheet.Range(.Columns(lngColCount + lngDiff + 1), _
                                    .Columns(lngColCount)).Delete(xlShiftToLeft)
        End With
    Case Is > 0 '�g��
        '�}��
        With objWkRange.Columns
            Call .Item(.Count).Copy(objWorksheet.Range(.Item(lngColCount + 1), .Item(lngColCount + lngDiff)))
            Call objWorksheet.Range(.Item(lngColCount + 1), .Item(lngColCount + lngDiff)).ClearContents
        End With
        
        '�s�̐��������[�v
        For i = 1 To lngRowCount
            '�Z��������
            With objWkRange.Columns(lngColCount)
                Call objWorksheet.Range(.Cells(i, 1), .Cells(i, lngDiff + 1).MergeArea).Merge
            End With
        Next i
    End Select
    
    Set objWkRange = objWkRange.Resize(, lngColCount + lngDiff)
    
    '�폜�̎��E�[�̌r�����Đݒ�
    If lngDiff < 0 Then
        '�E�[�̌r�����Đݒ�
        With objWkRange.Borders
            Call SetBorder(udtBorder, .Item(xlEdgeRight))
        End With
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetDirection
'[ �T  �v ]  �ړ��������擾
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]  �ړ������F�ړ��Ȃ��E��E���E���E�E�E�ȊO(�G���[)
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
'[ �֐��� ]�@GetActionRange
'[ �T  �v ]�@�؂������Z���̑}���̑ΏۂƂȂ�̈�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]  �Ώۗ̈�
'*****************************************************************************
Private Function GetActionRange() As Range
    Select Case GetDirection()
    Case E_UP, E_LEFT
        '��ƌ�����ړ���ň͂܂ꂽ�̈��Ώ�
        Set GetActionRange = Range(FromRange, ToRange)
    Case E_DOWN
        '��ƌ�����ړ����1�s�O�܂ł�Ώ�
        Set GetActionRange = Range(FromRange, ToRange(0, 1))
    Case E_RIGHT
        '��ƌ�����ړ����1��O�܂ł�Ώ�
        Set GetActionRange = Range(FromRange, ToRange(1, 0))
    End Select
End Function

'*****************************************************************************
'[�v���p�e�B]�@FromRange
'[ �T  �v ]�@�ړ����̗̈�
'*****************************************************************************
Private Property Get FromRange() As Range
    Set FromRange = objFromSheet.Range(strFromRange)
End Property

'*****************************************************************************
'[�v���p�e�B]�@ToRange
'[ �T  �v ]�@�ړ���̗̈�
'*****************************************************************************
Private Property Get ToRange() As Range
    Set ToRange = Range(strToRange)
End Property

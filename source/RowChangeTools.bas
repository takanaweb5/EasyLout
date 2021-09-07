Attribute VB_Name = "RowChangeTools"
Option Explicit
Option Private Module

Private Const MaxRowHeight = 409.5  '�����̍ő�T�C�Y

'*****************************************************************************
'[ �֐��� ]�@ChangeHeight
'[ �T  �v ]�@�����̕ύX
'[ ��  �� ]�@lngSize:�ύX�T�C�Y(�P�ʁF�s�N�Z��)
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ChangeHeight(ByVal lngSize As Long)
On Error GoTo ErrHandle
    '[Ctrl]Key����������Ă���΁A�ړ�����5�{�ɂ���
    If FPressKey = E_Ctrl Then
        lngSize = lngSize * 5
    End If
    
    '�I������Ă���I�u�W�F�N�g�𔻒�
    Select Case CheckSelection()
    Case E_Range
        Call ChangeRowsHeight(lngSize)
    Case E_Shape
        Call ChangeShapeHeight(lngSize)
    End Select
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@MoveHorizonBorder
'[ �T  �v ]�@�s�̋��E�����㉺�Ɉړ�����
'[ ��  �� ]�@lngSize:�ύX�T�C�Y(�P�ʁF�s�N�Z��)
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveHorizonBorder(ByVal lngSize As Long)
On Error GoTo ErrHandle
    '[Ctrl]Key����������Ă���΁A�ړ�����5�{�ɂ���
    If FPressKey = E_Ctrl Then
        lngSize = lngSize * 5
    End If
    
    '�I������Ă���I�u�W�F�N�g�𔻒�
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
'[ �֐��� ]�@ChangeRowsHeight
'[ �T  �v ]�@�����s�T�C�Y�ύX
'[ ��  �� ]�@lngSize:�ύX�T�C�Y(�P�ʁF�s�N�Z��)
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ChangeRowsHeight(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objSelection As Range   '�I�����ꂽ���ׂĂ̗�
    Dim strSelection As String
    
    '�I��͈͂�Rows�̘a�W�������d���s��r������
    strSelection = Selection.Address
    Set objSelection = Union(Selection.EntireRow, Selection.EntireRow)
    
    '***********************************************
    '��\���̍s�����邩�ǂ����̔���
    '***********************************************
    Dim objVisible     As Range    '��Range
    '�I��͈͂̉���������o��
    Set objVisible = GetVisibleCells(objSelection)
    If objVisible Is Nothing Then
        If lngSize < 0 Then
            Call MsgBox("����ȏ�k���o���܂���", vbExclamation)
            Exit Sub
        End If
    Else
        '��\���̍s�����鎞
        If objSelection.Address <> objVisible.Address Then
            If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
                If lngSize < 0 And FPressKey = E_Shift Then
                    Set objSelection = objVisible
                Else
                    Select Case MsgBox("��\���̍s��ΏۂƂ��܂����H", vbYesNoCancel + vbQuestion + vbDefaultButton2)
                    Case vbYes
                        If lngSize < 0 Then
                            Call MsgBox("����ȏ�k���o���܂���", vbExclamation)
                            Exit Sub
                        End If
                    Case vbNo
                        '���Z���̂ݑI������
                        Call IntersectRange(Selection, objVisible).Select
                        Set objSelection = objVisible
                    Case vbCancel
                        Exit Sub
                    End Select
                End If
            Else
                '���Z���̂ݑI������
                Call IntersectRange(Selection, objVisible).Select
                Set objSelection = objVisible
            End If
        End If
    End If
    
    '***********************************************
    '���������̉򂲂Ƃ�Address���擾����
    '***********************************************
    Dim colAddress  As New Collection
    If objVisible Is Nothing Then
        Call colAddress.Add(objSelection.Address)
    Else
        Set colAddress = GetSameHeightAddresses(objSelection)
    End If
    
    '***********************************************
    '�ύX��̃T�C�Y�̃`�F�b�N
    '***********************************************
    Dim lngPixel    As Long    '��(�P��:Pixel)
    If lngSize < 0 And FPressKey = E_Shift Then
    Else
        For i = 1 To colAddress.Count
            lngPixel = GetRange(colAddress(i)).Rows(1).Height / DPIRatio + lngSize
            If lngPixel < 0 Then
                Call MsgBox("����ȏ�k���o���܂���", vbExclamation)
                Exit Sub
            End If
        Next i
    End If
    
    '***********************************************
    '�T�C�Y�̕ύX
    '***********************************************
    Dim blnDisplayPageBreaks As Boolean  '���y�[�W�\��
    Application.ScreenUpdating = False
    
    '�������̂��߉��y�[�W���\���ɂ���
    If ActiveSheet.DisplayAutomaticPageBreaks = True Then
        blnDisplayPageBreaks = True
        ActiveSheet.DisplayAutomaticPageBreaks = False
    End If
    
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Call SaveUndoInfo(E_RowSize2, GetRange(strSelection), colAddress)
    
    'SHIFT����������Ă���Ɣ�\���ɂ���
    If lngSize < 0 And FPressKey = E_Shift Then
        objSelection.EntireRow.Hidden = True
    Else
        '���������̉򂲂Ƃɍ�����ݒ肷��
        For i = 1 To colAddress.Count
            lngPixel = GetRange(colAddress(i)).Rows(1).Height / DPIRatio + lngSize
            GetRange(colAddress(i)).RowHeight = PixelToHeight(lngPixel)
        Next i
    End If
    
    '���y�[�W�����ɖ߂�
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
'[�T�v] ���������̉򂲂Ƃ̃A�h���X��z��Ŏ擾����
'[����] �I�����ꂽ�̈�
'[�ߒl] �A�h���X�̔z��
'*****************************************************************************
Public Function GetSameHeightAddresses(ByRef objSelection As Range) As Collection
    Dim i             As Long
    Dim lngLastRow    As Long    '�g�p����Ă���Ō�̍s
    Dim lngLastCell   As Long
    Dim objRange      As Range
    Dim objWkRange    As Range
    Dim objRows       As Range
    Dim objVisible    As Range
    Dim objNonVisible As Range
    
    Set GetSameHeightAddresses = New Collection
    
    '�g�p����Ă���Ō�̍s
    '�g�p����Ă���Ō�̗�
    With Cells.SpecialCells(xlCellTypeLastCell)
        lngLastRow = .Row + .Rows.Count - 1
    End With
    
    '�I��͈͂�Rows�̘a�W�������d���s��r������
    Set objRows = Union(objSelection.EntireRow, objSelection.EntireRow)
    
    '���Z���ƕs���Z���ɕ�����
    Set objVisible = GetVisibleCells(objRows)
    Set objNonVisible = MinusRange(objRows, objVisible)
    
    '***********************************************
    '�g�p���ꂽ�Ō�̍s�ȑO�̗̈�̐ݒ�(���̈�)
    '***********************************************
    Set objWkRange = IntersectRange(Range(Rows(1), Rows(lngLastRow)), objVisible)
    If Not (objWkRange Is Nothing) Then
        '�G���A�̐��������[�v
        For Each objRange In objWkRange.Areas
            i = objRange.Row
            lngLastCell = i + objRange.Rows.Count - 1
                    
            '���������̉򂲂Ƃɍs���𔻒肷��
            While i <= lngLastCell
                '���������̍s�̃A�h���X��ۑ�
                 Call GetSameHeightAddresses.Add(GetSameHeightAddress(i, lngLastCell))
            Wend
        Next
    End If
    
    '***********************************************
    '�g�p���ꂽ�Ō�̍s�ȍ~�̗̈�̐ݒ�(���̈�)
    '***********************************************
    Set objWkRange = IntersectRange(Range(Rows(lngLastRow + 1), Rows(Rows.Count)), objVisible)
    If Not (objWkRange Is Nothing) Then
        Call GetSameHeightAddresses.Add(objWkRange.Address)
    End If
    
    '***********************************************
    '�s���̈�̐ݒ�
    '***********************************************
    If Not (objNonVisible Is Nothing) Then
        Call GetSameHeightAddresses.Add(objNonVisible.Address)
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@GetSameHeightAddress
'[ �T  �v ]�@�A������s��lngRow�Ɠ��������̍s��\�킷�A�h���X���擾����
'[ ��  �� ]�@lngRow:�ŏ��̍s(���s��͎��̍s)�AlngLastCell:�����̍Ō�̍s
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Function GetSameHeightAddress(ByRef lngRow As Long, ByVal lngLastCell As Long) As String
    Dim lngPixel As Long
    Dim i As Long
    lngPixel = Rows(lngRow).Height / DPIRatio
    
    For i = lngRow + 1 To lngLastCell
        If (Rows(i).Height / DPIRatio) <> lngPixel Then
            Exit For
        End If
    Next i
    GetSameHeightAddress = Range(Rows(lngRow), Rows(i - 1)).Address
    lngRow = i
End Function

'*****************************************************************************
'[ �֐��� ]�@ChangeShapeHeight
'[ �T  �v ]�@�}�`�̃T�C�Y�ύX
'[ ��  �� ]�@lngSize:�ύX�T�C�Y
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ChangeShapeHeight(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim objGroups      As ShapeRange
    Dim blnFitGrid     As Boolean
    
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Application.ScreenUpdating = False
    Call SaveUndoInfo(E_ShapeSize2, Selection.ShapeRange)
    
    '��]���Ă�����̂��O���[�v������
    Set objGroups = GroupSelection(Selection.ShapeRange)
    
    '[Shift]Key����������Ă���΁A�g���ɍ��킹�ĕύX����
    If FPressKey = E_Shift Then
        blnFitGrid = True
    End If
    
    '�}�`�̃T�C�Y��ύX
    Call ChangeShapesHeight(objGroups, lngSize, blnFitGrid)
    
    '��]���Ă���}�`�̃O���[�v�������������̐}�`��I������
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@ChangeShapesHeight
'[ �T  �v ]�@�}�`�̃T�C�Y�ύX
'[ ��  �� ]�@objShapes:�}�`
'            lngSize:�ύX�T�C�Y(Pixel)
'            blnFitGrid:�g���ɂ��킹�邩
'            blnTopLeft:���܂��͏�����ɕω�������
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub ChangeShapesHeight(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean, Optional ByVal blnTopLeft As Boolean = False)
    Dim objShape     As Shape
    Dim lngTop       As Long
    Dim lngBottom    As Long
    Dim lngOldHeight As Long
    Dim lngNewHeight As Long
    Dim lngNewTop    As Long
    Dim lngNewBottom As Long
    
    '�}�`�̐��������[�v
    For Each objShape In objShapes
        lngOldHeight = Round(objShape.Height / DPIRatio)
        lngTop = Round(objShape.Top / DPIRatio)
        lngBottom = Round((objShape.Top + objShape.Height) / DPIRatio)
        
        '�g���ɂ��킹�邩
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
            '�s�N�Z���P�ʂ̕ύX������
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
            objShape.Top = (lngBottom - lngNewHeight) * DPIRatio
        End If
        objShape.Height = lngNewHeight * DPIRatio
        
        'Excel2007�̃o�O�Ή�
        If Round(objShape.Height / DPIRatio) <> lngNewHeight Then
            objShape.Height = (lngNewHeight + lngSize) * DPIRatio
        End If
        
        If Round(objShape.Height / DPIRatio) <> lngOldHeight Then
            If blnTopLeft = True Then
                objShape.Top = (lngBottom - lngNewHeight) * DPIRatio
            Else
                objShape.Top = lngTop * DPIRatio
            End If
        End If
    Next objShape
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetTopGrid
'[ �T  �v ]�@���͂̈ʒu�̏�̘g���̈ʒu���擾(�P�ʃs�N�Z��)
'[ ��  �� ]�@lngPos:�ʒu(�P�ʃs�N�Z��)
'            objRow: lngPos���܂ލs
'[ �߂�l ]�@�}�`�̏㑤�̘g���̈ʒu(�P�ʃs�N�Z��)
'*****************************************************************************
Public Function GetTopGrid(ByVal lngPos As Long, ByRef objRow As Range) As Long
    Dim i      As Long
    Dim lngTop As Long
    
    If lngPos <= Round(Rows(2).Top / DPIRatio) Then
        GetTopGrid = 0
        Exit Function
    End If
    
    For i = objRow.Row To 1 Step -1
        lngTop = Round(Rows(i).Top / DPIRatio)
        If lngTop < lngPos Then
            GetTopGrid = lngTop
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[ �֐��� ]�@GetBottomGrid
'[ �T  �v ]�@���͂̈ʒu�̉��̘g���̈ʒu���擾(�P�ʃs�N�Z��)
'[ ��  �� ]�@lngPos:�ʒu(�P�ʃs�N�Z��)
'            objRow: lngPos���܂ލs
'[ �߂�l ]�@�}�`�̉����̘g���̈ʒu(�P�ʃs�N�Z��)
'*****************************************************************************
Public Function GetBottomGrid(ByVal lngPos As Long, ByRef objRow As Range) As Long
    Dim i         As Long
    Dim lngBottom As Long
    Dim lngMax    As Long
    
    lngMax = Round((Rows(Rows.Count).Top + Rows(Rows.Count).Height) / DPIRatio)
    
    If lngPos >= Round(Rows(Rows.Count).Top / DPIRatio) Then
        GetBottomGrid = lngMax
        Exit Function
    End If
    
    For i = objRow.Row + 1 To Rows.Count
        lngBottom = Round(Rows(i).Top / DPIRatio)
        If lngBottom > lngPos Then
            GetBottomGrid = lngBottom
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[ �֐��� ]�@GetVisibleCells
'[ �T  �v ]�@���Z�����擾
'[ ��  �� ]�@�I���Z��
'[ �߂�l ]�@���Z��
'*****************************************************************************
Private Function GetVisibleCells(ByRef objRange As Range) As Range
On Error GoTo ErrHandle
    Dim objCells As Range
    Set objCells = objRange.SpecialCells(xlCellTypeVisible)
    
    '��̔�\���͑I������
    Set GetVisibleCells = Union(objCells.EntireRow, objCells.EntireRow)
    Set GetVisibleCells = IntersectRange(GetVisibleCells, objRange)
Exit Function
ErrHandle:
    Set GetVisibleCells = Nothing
End Function

'*****************************************************************************
'[ �֐��� ]�@MoveBorder
'[ �T  �v ]�@�s�̋��E�����㉺�Ɉړ�����
'[ ��  �� ]�@lngSize : �ړ��T�C�Y(�P��:�s�N�Z��)
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveBorder(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim strSelection      As String
    Dim objRange          As Range
    Dim lngPixel(1 To 2)  As Long   '�擪�s�ƍŏI�s�̃T�C�Y
    Dim k                 As Long   '�ŏI�s�̍s�ԍ�
    
    strSelection = Selection.Address
    Set objRange = Selection

    '�I���G���A�������Ȃ�ΏۊO
    If objRange.Areas.Count <> 1 Then
        Call MsgBox("���̃R�}���h�͕����̑I��͈͂ɑ΂��Ď��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If

    '�I���s���P�s�Ȃ�ΏۊO
    If objRange.Rows.Count = 1 Then
        Exit Sub
    End If
    
    '�ŏI�s�̍s�ԍ�
    k = objRange.Rows.Count
    
    '�ύX��̃T�C�Y
    lngPixel(1) = objRange.Rows(1).Height / DPIRatio + lngSize '�擪�s
    lngPixel(2) = objRange.Rows(k).Height / DPIRatio - lngSize '�ŏI�s
    
    '�T�C�Y�̃`�F�b�N
    If lngPixel(1) < 0 Or lngPixel(2) < 0 Then
        Exit Sub
    End If
    
    '***********************************************
    '�T�C�Y�̕ύX
    '***********************************************
    Application.ScreenUpdating = False
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Dim colAddress  As New Collection
    Call colAddress.Add(objRange.Rows(1).Address)
    Call colAddress.Add(objRange.Rows(k).Address)
    Call SaveUndoInfo(E_RowSize2, Selection, colAddress)
    
    '�T�C�Y�̕ύX
    objRange.Rows(1).RowHeight = PixelToHeight(lngPixel(1))
    objRange.Rows(k).RowHeight = PixelToHeight(lngPixel(2))
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

''*****************************************************************************
''[ �֐��� ]�@MoveShape
''[ �T  �v ]�@�}�`���㉺�Ɉړ�����
''[ ��  �� ]�@lngSize�F�ړ��T�C�Y
''[ �߂�l ]�@�Ȃ�
''*****************************************************************************
'Private Sub MoveShape(ByVal lngSize As Long)
'On Error GoTo ErrHandle
'    Dim blnFitGrid  As Boolean
'    Dim objGroups   As ShapeRange
'
'    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
'    Application.ScreenUpdating = False
'    Call SaveUndoInfo(E_ShapeSize2, Selection.ShapeRange)
'
'    '[Shift]Key����������Ă���΁A�g���ɍ��킹�ĕύX����
'    If FPressKey = E_Shift Then
'        blnFitGrid = True
'    End If
'
'    '�g���ɂ��킹�邩
'    If blnFitGrid = True Then
'        '��]���Ă���}�`���O���[�v������
'        Set objGroups = GroupSelection(Selection.ShapeRange)
'
'        '�}�`�����E�Ɉړ�����
'        Call MoveShapesUD(objGroups, lngSize, blnFitGrid)
'
'        '��]���Ă���}�`�̃O���[�v�������������̐}�`��I������
'        Call UnGroupSelection(objGroups).Select
'    Else
'        '�}�`�����E�Ɉړ�����
'        Call MoveShapesUD(Selection.ShapeRange, lngSize, blnFitGrid)
'    End If
'
'    Call SetOnUndo
'Exit Sub
'ErrHandle:
'    Call MsgBox(Err.Description, vbExclamation)
'End Sub

'*****************************************************************************
'[ �֐��� ]�@DistributeRowsHeight
'[ �T  �v ]�@�I�����ꂽ�s�̍����𑵂���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub DistributeRowsHeight()
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objRange     As Range
    Dim lngRowCount  As Long    '�I�����ꂽ�s�̐�
    Dim dblHeight    As Double  '�I�����ꂽ�s�̍����̍��v
    Dim objSelection As Range   '�I�����ꂽ���ׂĂ̍s�̃R���N�V����
    Dim strSelection As String
    Dim objVisible   As Range   '�I�����ꂽ���̍s
    
    '�I������Ă���I�u�W�F�N�g�𔻒�
    Select Case CheckSelection()
    Case E_Other
        Exit Sub
    Case E_Shape
        Call DistributeShapeHeight
        Exit Sub
    End Select
    
    '�I��͈͂�Rows�̘a�W�������d���s��r������
    strSelection = Selection.Address
    Set objSelection = Union(Selection.EntireRow, Selection.EntireRow)
    
    '�I��͈͂̉���������o��
    Set objVisible = GetVisibleCells(objSelection)
    
    '���ׂĔ�\���̎�
    If objVisible Is Nothing Then
        Exit Sub
    End If
    
    '��\���̍s�����鎞
    If objSelection.Address <> objVisible.Address Then
        If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
            Select Case MsgBox("��\���̍s��ΏۂƂ��܂����H", vbYesNoCancel + vbQuestion + vbDefaultButton2)
            Case vbNo
                Set objSelection = objVisible
            Case vbCancel
                Exit Sub
            End Select
        Else
            Set objSelection = objVisible
        End If
    End If
    
    '�G���A�̐��������[�v
    For Each objRange In objSelection.Areas
        dblHeight = dblHeight + GetHeight(objRange)
        lngRowCount = lngRowCount + objRange.Rows.Count
    Next objRange
    
    If lngRowCount = 1 Then
        Exit Sub
    End If
    
    '***********************************************
    '�T�C�Y�̕ύX
    '***********************************************
    Application.ScreenUpdating = False
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Call SaveUndoInfo(E_RowSize, Range(strSelection), GetSameHeightAddresses(objSelection))
    objSelection.RowHeight = dblHeight / lngRowCount
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@DistributeShapeHeight
'[ �T  �v ]�@�I�����ꂽ�}�`�̍����𑵂���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub DistributeShapeHeight()
On Error GoTo ErrHandle
    If Selection.ShapeRange.Count = 1 Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Call SaveUndoInfo(E_ShapeSize, Selection.ShapeRange)
    
    '��]���Ă���}�`���O���[�v������
    Dim objGroups As ShapeRange
    Set objGroups = GroupSelection(Selection.ShapeRange)
    
    Call DistributeShapesHeight(objGroups)
    
    '��]���Ă���}�`�̃O���[�v�������������̐}�`��I������
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@DistributeShapesHeight
'[ �T  �v ]�@�}�`�̍����𑵂���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub DistributeShapesHeight(ByRef objShapeRange As ShapeRange)
    Dim objShape   As Shape
    Dim dblHeight  As Double
    
    '�}�`�̐��������[�v
    For Each objShape In objShapeRange
        dblHeight = dblHeight + objShape.Height
    Next objShape
    With objShapeRange
        .Height = Round(dblHeight / .Count / DPIRatio) * DPIRatio
    End With
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetHeight
'[ �T  �v ]�@�I���G���A�̍������擾
'            Height�v���p�e�B��32767�ȏ�̍������v�Z�o���Ȃ�����
'[ ��  �� ]�@�������擾����G���A
'[ �߂�l ]�@����(Height�v���p�e�B)
'*****************************************************************************
Private Function GetHeight(ByRef objRange As Range) As Double
    With objRange
        GetHeight = .Rows(.Rows.Count).Top - .Top + .Rows(.Rows.Count).Height
    End With
    
'    Dim lngCount   As Long
'    Dim lngHalf    As Long
'    Dim MaxHeight  As Double '�����̍ő�l
'
'    MaxHeight = 32767 * DPIRatio
'    If objRange.Height < MaxHeight Then
'        GetHeight = objRange.Height
'    Else
'        With objRange
'            '�O���{�㔼�̍��������v
'            lngCount = .Rows.Count
'            lngHalf = lngCount / 2
'            GetHeight = GetHeight(Range(.Rows(1), .Rows(lngHalf))) + _
'                        GetHeight(Range(.Rows(lngHalf + 1), .Rows(lngCount)))
'        End With
'    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@MergeCellsAsRow
'[ �T  �v ]�@�c�����Ɍ���(�����s�ɒl�����鎞�͉��s�ŘA��)
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
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
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    strSelection = Selection.Address
    lngCalculation = Application.Calculation
    
    '***********************************************
    '�d���̈�̃`�F�b�N
    '***********************************************
    If CheckDupRange(Range(strSelection)) = True Then
        Call MsgBox("�I������Ă���̈�ɏd��������܂�", vbExclamation)
        Exit Sub
    End If
    
    '***********************************************
    '�ύX
    '***********************************************
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    Call SaveUndoInfo(E_MergeCell, Range(strSelection))
    Call Range(strSelection).UnMerge
    
    '�G���A�̐��������[�v
    For Each objRange In Range(strSelection).Areas
        '��̐��������[�v
        For i = 1 To objRange.Columns.Count
            
            '�����̃Z���ɒl�����鎞�A������
            Set objMergeCell = objRange.Columns(i)
            If WorksheetFunction.CountA(objMergeCell) > 1 Then
                strValues = Replace$(GetRangeText(objMergeCell), vbTab, " ")
            Else
                strValues = ""
            End If
            
            '�Z������������
            Call objMergeCell.Merge
            
            '�A�������l���Đݒ肷��
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
'[ �֐��� ]�@SplitRow
'[ �T  �v ]�@�s�𕪊�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub SplitRow()
On Error GoTo ErrHandle
    Dim i               As Long
    Dim objRange        As Range
    Dim lngPixel        As Double  '�P�s�̍���
    Dim lngSplitCount   As Long    '������
    Dim blnCheckInsert  As Boolean
    Dim objNewRow       As Range    '�V�����s
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objRange = Selection
    
    '�I���G���A�������Ȃ�ΏۊO
    If objRange.Areas.Count <> 1 Then
        Call MsgBox("���̃R�}���h�͕����̑I��͈͂ɑ΂��Ď��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    '�����s�I���Ȃ�ΏۊO
    If objRange.Rows.Count <> 1 Then
        Call MsgBox("���̃R�}���h�͕����̑I���s�ɑ΂��Ď��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    '���̍���
    lngPixel = objRange.EntireRow.Height / DPIRatio
    
    '****************************************
    '��������I��������
    '****************************************
    With frmSplitCount
        Call .SetChkLabel(False)

        '�t�H�[����\��
        Call .Show
    
        '�L�����Z����
        If blnFormLoad = False Then
            Exit Sub
        End If
        
        lngSplitCount = .Count
        blnCheckInsert = .CheckInsert
        Call Unload(frmSplitCount)
    End With
    
    '****************************************
    '�����J�n
    '****************************************
    Dim blnDisplayPageBreaks As Boolean
        
    '�������̂��߉��y�[�W���\���ɂ���
    If ActiveSheet.DisplayAutomaticPageBreaks = True Then
        blnDisplayPageBreaks = True
        ActiveSheet.DisplayAutomaticPageBreaks = False
    End If
    Application.ScreenUpdating = False
    
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    Call SaveUndoInfo(E_SplitRow, objRange, lngSplitCount)
    If blnCheckInsert = False Then
        Call SetPlacement
    End If
    
    '�I���s�̉��ɂP�s�}��
    Call objRange(2, 1).EntireRow.Insert
    
    '�V�����s
    Set objNewRow = objRange(2, 1).EntireRow
    
    '*************************************************
    '�r���𐮂���
    '*************************************************
    '�}���s�̂P�Z�����Ɍr�����R�s�[����
    If blnCheckInsert = True Then
        Call CopyBorder("�㉺���E", objRange.EntireRow, objNewRow)
    Else
        Call CopyBorder("�����E", objRange.EntireRow, objNewRow)
    End If
    
    '*************************************************
    '�c�����Ɍ�������
    '*************************************************
    If blnCheckInsert = False Then
        Call MergeRows(2, objRange.EntireRow, objNewRow)
    Else
        Call MergeRows(3, objRange.EntireRow, objNewRow)
    End If
    
    '*************************************************
    '�������J�Ԃ�
    '*************************************************
    '�����������A�s��}������
    For i = 3 To lngSplitCount
        Call objNewRow.EntireRow.Insert
    Next i
    
    '*************************************************
    '�e�s���������Ɍ�������
    '*************************************************
    If blnCheckInsert = True Then
        For i = 2 To lngSplitCount
            Call MergeRows(4, objRange.EntireRow, objRange(i, 1).EntireRow)
        Next i
    End If

    '*************************************************
    '�����̐���
    '*************************************************
    '�V���������ɐݒ�
    If blnCheckInsert = False Then
        If lngSplitCount = 2 Then
            objRange.EntireRow.RowHeight = PixelToHeight(Round(lngPixel / 2))
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
    '���E��������
    '*************************************************
    If blnCheckInsert = False Then
        With Range(objRange, objRange(lngSplitCount, 1)).EntireRow
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
    End If
    
    '*************************************************
    '�㏈��
    '*************************************************
    Call Range(objRange, objRange(lngSplitCount, 1)).Select
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
'[ �֐��� ]�@EraseRow
'[ �T  �v ]�@�I�����ꂽ�s����������
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub EraseRow()
On Error GoTo ErrHandle
    Dim objSelection      As Range
    Dim strSelection      As String
    Dim objRange          As Range
    Dim lngTopRow         As Long  '��ׂ̍s�ԍ�
    Dim lngBottomRow      As Long  '���ׂ̍s�ԍ�
    Dim objColumn         As Range '�I�����ɑI���������
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = Selection
    strSelection = objSelection.Address
    Set objRange = objSelection.EntireRow
    
    '�I�����ɑI���������
    Set objColumn = objSelection.EntireColumn
    
    '�I���G���A�������Ȃ�ΏۊO
    If objSelection.Areas.Count <> 1 Then
        Call MsgBox("���̃R�}���h�͕����̑I��͈͂ɑ΂��Ď��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    If objSelection.Rows.Count = Rows.Count Then
        Call MsgBox("���̃R�}���h�͂��ׂĂ̍s�̑I���ɑ΂��Ď��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    '��ׂ̍s�ԍ�
    lngTopRow = objRange.Row - 1
    '���ׂ̍s�ԍ�
    lngBottomRow = objRange.Row + objRange.Rows.Count
    
    '****************************************
    '�����̃p�^�[����I��������
    '****************************************
    Dim enmSelectType As ESelectType  '�����p�^�[��
    Dim blnHidden   As Boolean        '��\���Ƃ��邩�ǂ���
    With frmEraseSelect
        '�V�[�g�̂P�s�ڂ��I������Ă��邩
        .TopSelect = (lngTopRow = 0)
        
        '�I���t�H�[����\��
        Call .Show
    
        '�����
        If blnFormLoad = False Then
            Exit Sub
        End If
        
        enmSelectType = .SelectType
        blnHidden = .Hidden
        Call Unload(frmEraseSelect)
    End With
    
    '****************************************
    '�l�̂���s���폜���邩�ǂ�������
    '****************************************
    Dim objValueCell As Range '�폜������Œl�̊܂܂��Z��
    If blnHidden = False Then
        Set objValueCell = SearchValueCell(objRange)
        '�폜�����s�Œl�̊܂܂��Z������������
        If Not (objValueCell Is Nothing) Then
            Call objValueCell.Select
            If MsgBox("�l�̓��͂���Ă���Z�����폜����܂����A��낵���ł����H", vbOKCancel + vbQuestion + vbDefaultButton2) = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    
    '****************************************
    'Undo�p�ɍs����ۑ����邽�߂̏��擾
    '****************************************
    Dim colAddress   As New Collection
    Set colAddress = GetSameHeightAddresses(Range(strSelection))
    
    '****************************************
    '�I���s�̏㉺�̍s��Ҕ�
    '****************************************
    Dim objRow(0 To 1)    As Range  '�s�̍�����ς���s
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
    '�s������
    '****************************************
    Dim lngPixel  As Long   '���������s�̍���(�P��:�s�N�Z��)
    Application.ScreenUpdating = False
    
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    If blnHidden = True Then
        Call SaveUndoInfo(E_RowSize, Range(strSelection), colAddress)
    Else
        Call SaveUndoInfo(E_EraseRow, Range(strSelection), colAddress)
    End If
    
    '�}�`�͈ړ������Ȃ�
    Call SetPlacement
    
    '���������s�̍�����ۑ�
    lngPixel = objRange.Height / DPIRatio
    
    If blnHidden = True Then
        '��\��
        objRange.Hidden = True
    Else
        '���̌r�����R�s�[����
        With objRange
            If .Row > 1 Then
                Call CopyBorder("��", .Rows(.Rows.Count), .Rows(0))
            End If
        End With
        
        '�폜
        Call objRange.Delete(xlShiftUp)
    End If
    
    '****************************************
    '�s�����g��
    '****************************************
    Dim lngWkPixel  As Long
    Select Case enmSelectType
    Case E_Front, E_Back
        objRow(0).RowHeight = WorksheetFunction.Min(MaxRowHeight, PixelToHeight(objRow(0).Height / DPIRatio + lngPixel))
    Case E_Middle
        objRow(0).RowHeight = WorksheetFunction.Min(MaxRowHeight, PixelToHeight(objRow(0).Height / DPIRatio + Int(lngPixel / 2 + 0.5)))
        objRow(1).RowHeight = WorksheetFunction.Min(MaxRowHeight, PixelToHeight(objRow(1).Height / DPIRatio + Int(lngPixel / 2)))
    End Select
    
    '****************************************
    '�Z����I��
    '****************************************
    Select Case enmSelectType
    Case E_Front, E_Back
        IntersectRange(objColumn, objRow(0)).Select
    Case E_Middle
        IntersectRange(objColumn, UnionRange(objRow(0), objRow(1))).Select
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
'[ �֐��� ]�@CopyBorder
'[ �T  �v ]�@�r�����R�s�[����
'[ ��  �� ]�@�r���̃^�C�v(�����w���):�㉺���E
'[ ��  �� ]�@objFromRow�F�R�s�|���̍s�AobjToRow�F�R�s�[��̍s
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub CopyBorder(ByVal strBorderType As String, ByRef objFromRow As Range, ByRef objToRow As Range)
    Dim i          As Long
    Dim udtBorder(0 To 3) As TBorder '�r���̎��(�㉺���E)
    Dim lngLast    As Long
    
    Call ActiveSheet.UsedRange '�Ō�̃Z�����C������ Undo�o���Ȃ��Ȃ�܂�
    If objFromRow.Columns.Count = Columns.Count Then
        '�Ō�̃Z���܂Ő�������ΏI������
        lngLast = Cells.SpecialCells(xlCellTypeLastCell).Column
        If lngLast > MAXROWCOLCNT Then
            lngLast = MAXROWCOLCNT
        End If
    Else
        '�I�����ꂽ���ׂĂ̗�𐮔�����
        lngLast = objFromRow.Columns.Count
    End If
    
    '1�񖈂Ƀ��[�v
    For i = 1 To lngLast
        '�r���̎�ނ�ۑ�
        With objFromRow.Columns(i)
            If InStr(1, strBorderType, "��") <> 0 Then
                udtBorder(0) = GetBorder(.Borders(xlEdgeTop))
            End If
            If InStr(1, strBorderType, "��") <> 0 Then
                udtBorder(1) = GetBorder(.Borders(xlEdgeBottom))
            End If
            If InStr(1, strBorderType, "��") <> 0 Then
                udtBorder(2) = GetBorder(.Borders(xlEdgeLeft))
            End If
            If InStr(1, strBorderType, "�E") <> 0 Then
                udtBorder(3) = GetBorder(.Borders(xlEdgeRight))
            End If
        End With
        
        '�r��������
        With objToRow.Columns(i)
            If InStr(1, strBorderType, "��") <> 0 Then
                Call SetBorder(udtBorder(0), .Borders(xlEdgeTop))
            End If
            If InStr(1, strBorderType, "��") <> 0 Then
                Call SetBorder(udtBorder(1), .Borders(xlEdgeBottom))
            End If
            If InStr(1, strBorderType, "��") <> 0 Then
                Call SetBorder(udtBorder(2), .Borders(xlEdgeLeft))
            End If
            If InStr(1, strBorderType, "�E") <> 0 Then
                Call SetBorder(udtBorder(3), .Borders(xlEdgeRight))
            End If
        End With
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ]�@MergeRows
'[ �T  �v ]�@�擪�s�����̍s�܂ŏc�����Ɍ�������
'[ ��  �� ]�@lngType:�����̃^�C�v�A
'            objTopRow�F�����̐擪�s�AobjBottomRow�F�����̒�̍s
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MergeRows(ByVal lngtype As Long, ByRef objTopRow As Range, ByRef objBottomRow As Range)
    Dim i          As Long
    Dim lngLast    As Long
    Dim objRange As Range
    
    Call ActiveSheet.UsedRange '�Ō�̃Z�����C������
    If objTopRow.Columns.Count = Columns.Count Then
        '�Ō�̃Z���܂Ő�������ΏI������
        lngLast = Cells.SpecialCells(xlCellTypeLastCell).Column
        If lngLast > MAXROWCOLCNT Then
            lngLast = MAXROWCOLCNT
        End If
    Else
        '�I�����ꂽ���ׂĂ̗�𐮔�����
        lngLast = objTopRow.Columns.Count
    End If
    
    '1�񖈂Ƀ��[�v
    For i = 1 To lngLast
        With objBottomRow.Cells(1, i)
            '��̍s�̃Z���������Z����
            If .MergeArea.Count = 1 Then
                Set objRange = GetMergeRowRange(lngtype, objTopRow.Cells(1, i), .Cells)
                If Not (objRange Is Nothing) Then
                    Call objRange.Merge
                End If
            End If
        End With
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetMergeRowRange
'[ �T  �v ]�@�c�����Ɍ�������̈���擾����
'[ ��  �� ]�@lngType:�����̃^�C�v�A
'            objBaseCell:�擪�s�̃Z���AobjTergetCell:�Ώۂ̍s�̃Z��
'[ �߂�l ]�@��������̈�(Nothing:�������Ȃ���)
'*****************************************************************************
Private Function GetMergeRowRange(ByVal lngtype As Long, _
                                  ByRef objBaseCell As Range, _
                                  ByRef objTergetCell As Range) As Range
    Select Case lngtype
    Case 1 '�擪�s����ŏI�̍s�܂ŏc�����Ɍ�������
    Case 2 '�擪�s�������Z���̎��A�擪�s����ŏI�̍s�܂ŏc�����Ɍ�������
        If objBaseCell.MergeArea.Count = 1 Then
            Exit Function
        End If
    Case 3 '�擪�s���c�����Ɍ����Z���̎��A�擪�s����ŏI�̍s�܂ŏc�����Ɍ�������
        If objBaseCell.MergeArea.Rows.Count = 1 Then
            Exit Function
        End If
    Case 4 '�擪�s���������̂݌����Z���̎��A�Ώۂ̍s�̃Z�����������Ɍ�������
        If objBaseCell.MergeArea.Columns.Count = 1 Or _
           objBaseCell.MergeArea.Rows.Count > 1 Then
            Exit Function
        End If
    End Select
    
    Select Case lngtype
    Case 1, 2, 3
        Set GetMergeRowRange = Range(objBaseCell.MergeArea, objTergetCell)
    Case 4
        Set GetMergeRowRange = objTergetCell.Resize(1, objBaseCell.MergeArea.Columns.Count)
    End Select
End Function

'*****************************************************************************
'[ �֐��� ]�@ShowHeight
'[ �T  �v ]�@�G���A���ɑI�����ꂽ�s�̍������ꗗ�ŕ\������
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ShowHeight()
On Error GoTo ErrHandle
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Call frmSizeList.Initialize(E_ROW)
    Call frmSizeList.Show
    Call Application.OnRepeat("", "")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@MoveRowsBorder
'[ �T  �v ]�@�s�̋��E�̃Z�����㉺�Ɉړ�����
'[ ��  �� ]�@-1:��Ɉړ��A1:���Ɉړ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveRowsBorder(ByVal lngUpDown As Long)
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objSelection As Range
    Dim objWkRange   As Range
    Dim lngRowCount  As Long  '�I��̈�̍s��
    Dim blnCopyObjectsWithCells As Boolean
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() = E_Range Then
        Set objSelection = Selection
    Else
        Exit Sub
    End If
    
    '�I���G���A�������Ȃ�ΏۊO
    If objSelection.Areas.Count <> 1 Then
        Call MsgBox("���̃R�}���h�͕����̑I��͈͂ɑ΂��Ď��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    '******************************************************************
    '�k���s�\�ȃZ�����Ȃ����`�F�b�N
    '******************************************************************
    lngRowCount = objSelection.Rows.Count
    Dim objChkRange(0 To 2) As Range

    If objSelection.Rows.Count = 1 Then
        Exit Sub
    End If

    '��Ɉړ����鎞
    If lngUpDown < 0 Then
        Set objChkRange(1) = ArrangeRange(Range(objSelection.Rows(1), objSelection.Rows(2)))
        Set objChkRange(2) = ArrangeRange(objSelection.Rows(2))
    Else '���Ɉړ����鎞
        Set objChkRange(1) = ArrangeRange(Range(objSelection.Rows(lngRowCount - 1), objSelection.Rows(lngRowCount)))
        Set objChkRange(2) = ArrangeRange(objSelection.Rows(lngRowCount - 1))
    End If
    
    '�ړ����鋫�E���Ȃ���
    If MinusRange(objSelection, objChkRange(1)) Is Nothing Then
        Exit Sub
    End If
    
    Set objChkRange(0) = MinusRange(objChkRange(1), objChkRange(2))
    If Not (objChkRange(0) Is Nothing) Then
        Call objChkRange(0).Select
        Call MsgBox("����ȏ�k���ł��Ȃ��Z��������܂�")
        Call objSelection.Select
        Exit Sub
    End If
    
    '****************************************
    '�ړ��J�n
    '****************************************
    '�}�`���R�s�[�̑ΏۊO�ɂ���
    Application.CopyObjectsWithCells = False
    Application.ScreenUpdating = False
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    Call SaveUndoInfo(E_CellBorder, objSelection)
    
    '****************************************
    '���̗̈���A"Workarea1"�V�[�g�ɃR�s�[
    '****************************************
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    With ThisWorkbook.Worksheets("Workarea1")
        '�R�����g���܂Ƃ��Ȉʒu�ɔz�u�o����悤�ɁA
        '���̕��ƍ������R�s�[���邽�߁A�V�[�g�S�̂��R�s�[������N���A
        Call ActiveSheet.Cells.Copy(.Cells)
        Call .Cells.Clear
        
        Set objWkRange = .Range(objSelection.Address)
        
        '�̈���R�s�[
        Call objSelection.Copy(objWkRange)
    End With
    
    '****************************************
    '���E���ړ�����
    '****************************************
    If lngUpDown < 0 Then
        '��Ɉړ�����
        Call CopyBorder("��", objWkRange.Rows(2), objWkRange.Rows(1))
        Call objWkRange.Rows(2).Delete(xlUp)
        Call CopyBorder("�����E", objWkRange.Rows(lngRowCount - 1), objWkRange.Rows(lngRowCount))
        Call MergeRows(1, objWkRange.Rows(lngRowCount - 1), objWkRange.Rows(lngRowCount))
    Else
        '���Ɉړ�����
        Call CopyBorder("��", objWkRange.Rows(lngRowCount), objWkRange.Rows(lngRowCount - 1))
        Call objWkRange.Rows(lngRowCount).Delete(xlUp)
        Call objWkRange.Rows(2).Insert(xlDown)
        Call CopyBorder("�����E", objWkRange.Rows(1), objWkRange.Rows(2))
        Call MergeRows(1, objWkRange.Rows(1), objWkRange.Rows(2))
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
'[ �֐��� ]�@AutoRowFit
'[ �T  �v ]�@�s�̍����𕶎��̍����ɂ��킹��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub AutoRowFit()
On Error GoTo ErrHandle
    Dim objSelection  As Range
    Dim objSelection2 As Range
    Dim objUsedRange  As Range
    Dim blnVisible    As Boolean
    Dim dblNewHeight  As Double
    Dim i As Long
    
    '�I������Ă���I�u�W�F�N�g�𔻒�
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = IntersectRange(Selection, GetUsedRange())
    If objSelection Is Nothing Then
        Exit Sub
    End If
    
    '������̌����Z�����܂ގ�
    If IsBorderMerged(objSelection) Then
        Call MsgBox("�������ꂽ�Z���̈ꕔ��I�����邱�Ƃ͂ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    '��\���̍s��ΏۂƂ��邩�m�F�H
    Dim objVisible    As Range
    Dim objNonVisible As Range
    Set objVisible = GetVisibleCells(objSelection)
    Set objNonVisible = MinusRange(objSelection, objVisible)
    If Not (objNonVisible Is Nothing) Then
        If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
            Select Case MsgBox("��\���̃Z����ΏۂƂ��܂����H", vbYesNoCancel + vbQuestion + vbDefaultButton2)
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
    
    'WorkSheet�̕W���X�^�C�������s�V�[�g�Ƃ��킹��
    Call SetPixelInfo 'Undo�ł��Ȃ��Ȃ�܂�
    
    '***********************************************
    '���s
    '***********************************************
    Application.ScreenUpdating = False
    
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Call SaveUndoInfo(E_RowSize, objSelection, GetSameHeightAddresses(objSelection))
    
    '�W���R�}���h�ō�����K����(�ҏW��Font��ς��Ă������Œ��������悤�ɂ���Ή�)
    Call objSelection2.Rows.AutoFit
    
    If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
    Else
        Call SetOnUndo
        Exit Sub
    End If
    
    '���삪���ɒx���Ȃ邽�߂̑Ή�
    If objSelection.Rows.Count > 100 Then
        Call SetOnUndo
        Exit Sub
    End If
    
    '*********************************************************
    'WorkSheet�𗘗p���A�s�̍�����K��������
    '*********************************************************
    Dim objWorksheet As Worksheet
    Set objWorksheet = ThisWorkbook.Worksheets("Workarea1")
    Call DeleteSheet(objWorksheet)
    With objWorksheet
        .Columns.ColumnWidth = 255
        .Range(.Rows(1), .Rows(objSelection.Rows.Count)).Font.size = 1
        Call objSelection.Copy(.Cells(1, 1))
    End With
    
    '�s�������[�v
    Dim objRow     As Range
    Dim objWorkRow As Range
    For i = 1 To objSelection.Rows.Count
        Set objRow = objSelection.Rows(i)
        Set objWorkRow = objWorksheet.Rows(i)
        
        '��\����ΏۊO�ɂ��邩�ǂ���
        If blnVisible Or objRow.Hidden = False Then
            '�s�����̌���������s�́A�W����AutoFit�̂ݍs���i���ł�AutoFit�͊������Ă���j
            If IsBorderMerged(objRow) = False Then
                'WorkSheet�𗘗p���A�s�̍�����K����
                dblNewHeight = GetFitRow(objRow, objWorkRow)
                '�ҏW��Font��ς��Ă������Œ��������悤���߂ɔ��肷��
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
'[ �֐��� ]�@GetFitRow
'[ �T  �v ]�@WorkSheet�𗘗p���A�s�̍�����K����
'[ ��  �� ]�@�Ώۂ̍s(���̃V�[�g)�A�Ώۂ̍s(���[�N�V�[�g)
'[ �߂�l ]�@�K�������RowHeight
'*****************************************************************************
Private Function GetFitRow(ByRef objRow As Range, ByRef objWorkRow As Range) As Double
    Dim objCell        As Range
    Dim dblColumnWidth As Double
    Dim objValueCells As Range
    
    '�l�̓��͂��ꂽ�Z���̕��𐮔�����
    On Error Resume Next
    Set objValueCells = objWorkRow.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If Not (objValueCells Is Nothing) Then
        For Each objCell In objValueCells
            '��̕����R�s�[����
            dblColumnWidth = WorksheetFunction.Min(PixelToWidth(objRow.Columns(objCell.Column).MergeArea.Width / DPIRatio), 255)
            With objCell
                If .ColumnWidth <> dblColumnWidth Then
                    .ColumnWidth = dblColumnWidth
                End If
                Call .UnMerge
            End With
        Next
    End If
    
    '������ݒ�
    Call objWorkRow.AutoFit
    
    GetFitRow = objWorkRow.RowHeight
End Function

'*****************************************************************************
'[ �֐��� ]�@PixelToHeight
'[ �T  �v ]�@�����̒P�ʂ�ϊ�
'[ ��  �� ]�@lngPixel : ����(�P��:�s�N�Z��)
'[ �߂�l ]�@Height
'*****************************************************************************
Public Function PixelToHeight(ByVal lngPixel As Long) As Double
    PixelToHeight = lngPixel * DPIRatio
End Function

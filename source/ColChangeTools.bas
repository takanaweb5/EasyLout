Attribute VB_Name = "ColChangeTools"
Option Explicit

Private Type TFont  '�W���X�^�C���̃t�H���g�̏��
    Name        As String
    Size        As Long
    Bold        As Boolean
    Italic      As Boolean
End Type

Private x1 As Byte '1�����̃s�N�Z��
Private x2 As Byte '2�����̃s�N�Z��

Private Const MaxColumnWidth = 255  '���̍ő�T�C�Y

Private Sub ReduceColsWidth()
    Call ChangeColsWidth(-1)
End Sub
Private Sub ExpandColsWidth()
    Call ChangeColsWidth(1)
End Sub
Private Sub MoveColBorderL()
    Call MoveColBorder(-1)
End Sub
Private Sub MoveColBorderR()
    Call MoveColBorder(1)
End Sub
Private Sub MoveCellBorderL()
    Call MoveCellBorder(-1)
End Sub
Private Sub MoveCellBorderR()
    Call MoveCellBorder(1)
End Sub

'*****************************************************************************
'[ �֐��� ]�@ChangeColsWidth
'[ �T  �v ]�@������T�C�Y�ύX
'[ ��  �� ]�@lngSize:�ύX�T�C�Y(�P�ʁF�s�N�Z��)
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ChangeColsWidth(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objSelection As Range   '�I�����ꂽ���ׂĂ̗�
    Dim strSelection As String
    
    '[Ctrl]Key����������Ă���΁A�ړ�����5�{�ɂ���
    If GetKeyState(vbKeyControl) < 0 Then
        lngSize = lngSize * 5
    End If
    
    '�I������Ă���I�u�W�F�N�g�𔻒�
    Select Case CheckSelection()
    Case E_Other
        Exit Sub
    Case E_Shape
        Call ChangeShapeWidth(lngSize)
        Exit Sub
    End Select
    
    '�I��͈͂�Columns�̘a�W�������d�����r������
    strSelection = Selection.Address
    Set objSelection = Union(Selection.EntireColumn, Selection.EntireColumn)
    
    '�s�N�Z������ݒ肷��
    Call SetPixelInfo
    
    '***********************************************
    '��\���̗񂪂��邩�ǂ����̔���
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
        '��\���̗񂪂��鎞
        If objSelection.Address <> objVisible.Address Then
            If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
                Select Case MsgBox("��\���̗��ΏۂƂ��܂����H", vbYesNoCancel + vbQuestion + vbDefaultButton2)
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
            Else
                '���Z���̂ݑI������
                Call IntersectRange(Selection, objVisible).Select
                Set objSelection = objVisible
            End If
        End If
    End If
    
    '***********************************************
    '�������̉򂲂Ƃ�Address���擾����
    '***********************************************
    Dim colAddress  As New Collection
    If objVisible Is Nothing Then
        Call colAddress.Add(objSelection.Address)
    Else
        Set colAddress = GetSameWidthAddresses(objSelection)
    End If
    
    '***********************************************
    '�ύX��̃T�C�Y�̃`�F�b�N
    '***********************************************
    Dim lngPixel    As Long    '��(�P��:Pixel)
    For i = 1 To colAddress.Count
        lngPixel = WidthToPixel(Range(colAddress(i)).Columns(1).ColumnWidth) + lngSize
        If lngPixel < 0 Then
            Call MsgBox("����ȏ�k���o���܂���", vbExclamation)
            Exit Sub
        ElseIf lngPixel > WidthToPixel(MaxColumnWidth) Then
            Call MsgBox("����ȏ�g��o���܂���", vbExclamation)
            Exit Sub
        End If
    Next i
    
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
    Call SaveUndoInfo(E_ColSize, Range(strSelection), colAddress)
    
    '�������̉򂲂Ƃɕ���ݒ肷��
    For i = 1 To colAddress.Count
        lngPixel = WidthToPixel(Range(colAddress(i)).Columns(1).ColumnWidth) + lngSize
        Range(colAddress(i)).ColumnWidth = PixelToWidth(lngPixel)
    Next i
    
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
'[ �֐��� ]�@GetSameWidthAddresses
'[ �T  �v ]�@�������̉򂲂Ƃ̃A�h���X��z��Ŏ擾����
'[ ��  �� ]�@�I�����ꂽ�̈�
'[ �߂�l ]�@�A�h���X�̔z��
'*****************************************************************************
Public Function GetSameWidthAddresses(ByRef objSelection As Range) As Collection
    Dim i           As Long
    Dim objRange    As Range
    Dim lngLastCell As Long
    
    Set GetSameWidthAddresses = New Collection
    
    '�G���A�̐��������[�v
    For Each objRange In objSelection.Areas
        i = objRange.Column
        lngLastCell = i + objRange.Columns.Count - 1
        
        '�������̉򂲂Ƃɗ񕝂𔻒肷��
        While i <= lngLastCell
            '�������̗�̃A�h���X��ۑ�
            Call GetSameWidthAddresses.Add(GetSameWidthAddress(i, lngLastCell))
        Wend
    Next objRange
End Function

'*****************************************************************************
'[ �֐��� ]�@GetSameWidthAddress
'[ �T  �v ]�@�A��������lngCol�Ɠ������̗��\�킷�A�h���X���擾����
'[ ��  �� ]�@lngCol:�ŏ��̗�(���s��͎��̗�)�AlngLastCell:�����̍Ō�̗�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Function GetSameWidthAddress(ByRef lngCol As Long, ByVal lngLastCell As Long) As String
    Dim lngPixel As Long
    Dim i As Long
    lngPixel = Columns(lngCol).Width / 0.75
    
    For i = lngCol + 1 To lngLastCell
        If (Columns(i).Width / 0.75) <> lngPixel Then
            Exit For
        End If
    Next i
    GetSameWidthAddress = Range(Columns(lngCol), Columns(i - 1)).Address
    lngCol = i
End Function

'*****************************************************************************
'[ �֐��� ]�@ChangeShapeWidth
'[ �T  �v ]�@�}�`�̃T�C�Y�ύX
'[ ��  �� ]�@lngSize:�ύX�T�C�Y
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ChangeShapeWidth(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim objGroups      As ShapeRange
    Dim blnFitGrid     As Boolean
    
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Application.ScreenUpdating = False
    Call SaveUndoInfo(E_ShapeSize, Selection.ShapeRange)
    
    '��]���Ă���}�`���O���[�v������
    Set objGroups = GroupSelection(Selection.ShapeRange)
    
    '[Shift]Key����������Ă���΁A�g���ɍ��킹�ĕύX����
    If GetKeyState(vbKeyShift) < 0 Then
        blnFitGrid = True
    End If
    
    '�}�`�̃T�C�Y��ύX
    Call ChangeShapesWidth(objGroups, lngSize, blnFitGrid)
    
    '��]���Ă���}�`�̃O���[�v�������������̐}�`��I������
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@ChangeShapesWidth
'[ �T  �v ]�@�}�`�̃T�C�Y�ύX
'[ ��  �� ]�@objShapes:�}�`
'            lngSize:�ύX�T�C�Y(Pixel)
'            blnFitGrid:�g���ɂ��킹�邩
'            blnTopLeft:���܂��͏�����ɕω�������
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub ChangeShapesWidth(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean, Optional ByVal blnTopLeft As Boolean = False)
    Dim objShape     As Shape
    Dim lngLeft      As Long
    Dim lngRight     As Long
    Dim lngOldWidth  As Long
    Dim lngNewWidth  As Long
    Dim lngNewLeft   As Long
    Dim lngNewRight  As Long
    
    '�}�`�̐��������[�v
    For Each objShape In objShapes
        lngOldWidth = Round(objShape.Width / 0.75)
        lngLeft = Round(objShape.Left / 0.75)
        lngRight = Round((objShape.Left + objShape.Width) / 0.75)
        
        '�g���ɂ��킹�邩
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
            '�s�N�Z���P�ʂ̕ύX������
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
            objShape.Left = (lngRight - lngNewWidth) * 0.75
        End If
        objShape.Width = lngNewWidth * 0.75
        
        'Excel2007�̃o�O�Ή�
        If Round(objShape.Width / 0.75) <> lngNewWidth Then
            objShape.Width = (lngNewWidth + lngSize) * 0.75
        End If
        
        If Round(objShape.Width / 0.75) <> lngOldWidth Then
            If blnTopLeft = True Then
                objShape.Left = (lngRight - lngNewWidth) * 0.75
            Else
                objShape.Left = lngLeft * 0.75
            End If
        End If
    Next objShape
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetLeftGrid
'[ �T  �v ]�@���͂̈ʒu�̍����̘g���̈ʒu���擾(�P�ʃs�N�Z��)
'[ ��  �� ]�@lngPos:�ʒu(�P�ʃs�N�Z��)
'            objColumn: lngPos���܂ޗ�
'[ �߂�l ]�@�}�`�̍����̘g���̈ʒu(�P�ʃs�N�Z��)
'*****************************************************************************
Public Function GetLeftGrid(ByVal lngPos As Long, ByRef objColumn As Range) As Long
    Dim i       As Long
    Dim lngLeft As Long
    
    If lngPos <= Round(Columns(2).Left / 0.75) Then
        GetLeftGrid = 0
        Exit Function
    End If
    
    For i = objColumn.Column To 1 Step -1
        lngLeft = Round(Columns(i).Left / 0.75)
        If lngLeft < lngPos Then
            GetLeftGrid = lngLeft
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[ �֐��� ]�@GetRightGrid
'[ �T  �v ]�@���͂̈ʒu�̉E���̘g���̈ʒu���擾(�P�ʃs�N�Z��)
'[ ��  �� ]�@lngPos:�ʒu(�P�ʃs�N�Z��)
'            objColumn: lngPos���܂ޗ�
'[ �߂�l ]�@�}�`�̉E���̘g���̈ʒu(�P�ʃs�N�Z��)
'*****************************************************************************
Public Function GetRightGrid(ByVal lngPos As Long, ByRef objColumn As Range) As Long
    Dim i        As Long
    Dim lngRight As Long
    Dim lngMax   As Long
    
    lngMax = Round((Columns(Columns.Count).Left + Columns(Columns.Count).Width) / 0.75)
    
    If lngPos >= Round(Columns(Columns.Count).Left / 0.75) Then
        GetRightGrid = lngMax
        Exit Function
    End If
    
    For i = objColumn.Column + 1 To Columns.Count
        lngRight = Round(Columns(i).Left / 0.75)
        If lngRight > lngPos Then
            GetRightGrid = lngRight
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
    
    '�s�̔�\���͑I������
    Set GetVisibleCells = Union(objCells.EntireColumn, objCells.EntireColumn)
    Set GetVisibleCells = IntersectRange(GetVisibleCells, objRange)
Exit Function
ErrHandle:
    Set GetVisibleCells = Nothing
End Function

'*****************************************************************************
'[ �֐��� ]�@MoveColBorder
'[ �T  �v ]�@��̋��E�������E�Ɉړ�����
'[ ��  �� ]�@lngSize : �ړ��T�C�Y(�P��:�s�N�Z��)
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveColBorder(ByVal lngSize As Long)
On Error GoTo ErrHandle
    Dim strSelection      As String
    Dim objRange          As Range
    Dim lngPixel(1 To 2)  As Long  '�擪��ƍŏI��̃T�C�Y
    Dim k                 As Long  '�ŏI��̗�ԍ�
    
    '[Ctrl]Key����������Ă���΁A�ړ�����5�{�ɂ���
    If GetKeyState(vbKeyControl) < 0 Then
        lngSize = lngSize * 5
    End If
    
    '�I������Ă���I�u�W�F�N�g�𔻒�
    Select Case CheckSelection()
    Case E_Other
        Exit Sub
    Case E_Shape
'        Call MoveShape(lngSize)
        Exit Sub
    End Select
    
    strSelection = Selection.Address
    Set objRange = Selection

    '�I���G���A�������Ȃ�ΏۊO
    If objRange.Areas.Count <> 1 Then
        Call MsgBox("���̃R�}���h�͕����̑I��͈͂ɑ΂��Ď��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If

    '�I��񂪂P��Ȃ�ΏۊO
    If objRange.Columns.Count = 1 Then
        Exit Sub
    End If
    
    '�s�N�Z������ݒ肷��
    Call SetPixelInfo
    
    '�ŏI��̗�ԍ�
    k = objRange.Columns.Count
    
    '�ύX��̃T�C�Y
    lngPixel(1) = WidthToPixel(objRange.Columns(1).ColumnWidth) + lngSize '�擪��
    lngPixel(2) = WidthToPixel(objRange.Columns(k).ColumnWidth) - lngSize '�ŏI��
    
    '�T�C�Y�̃`�F�b�N
    If (0 <= lngPixel(1) And lngPixel(1) <= WidthToPixel(MaxColumnWidth)) And _
       (0 <= lngPixel(2) And lngPixel(2) <= WidthToPixel(MaxColumnWidth)) Then
    Else
        Exit Sub
    End If
    
    '***********************************************
    '�T�C�Y�̕ύX
    '***********************************************
    Application.ScreenUpdating = False
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Dim colAddress  As New Collection
    Call colAddress.Add(objRange.Columns(1).Address)
    Call colAddress.Add(objRange.Columns(k).Address)
    Call SaveUndoInfo(E_ColSize, Range(strSelection), colAddress)
    
    '�T�C�Y�̕ύX
    objRange.Columns(1).ColumnWidth = PixelToWidth(lngPixel(1))
    objRange.Columns(k).ColumnWidth = PixelToWidth(lngPixel(2))
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

''*****************************************************************************
''[ �֐��� ]�@MoveShape
''[ �T  �v ]�@�}�`�����E�Ɉړ�����
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
'    Call SaveUndoInfo(E_ShapeSize, Selection.ShapeRange)
'
'    '[Shift]Key����������Ă���΁A�g���ɍ��킹�ĕύX����
'    If GetKeyState(vbKeyShift) < 0 Then
'        blnFitGrid = True
'    End If
'
'    '�g���ɂ��킹�邩
'    If blnFitGrid = True Then
'        '��]���Ă���}�`���O���[�v������
'        Set objGroups = GroupSelection(Selection.ShapeRange)
'
'        '�}�`�����E�Ɉړ�����
'        Call MoveShapesLR(objGroups, lngSize, blnFitGrid)
'
'        '��]���Ă���}�`�̃O���[�v�������������̐}�`��I������
'        Call UnGroupSelection(objGroups).Select
'    Else
'        '�}�`�����E�Ɉړ�����
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
'[ �֐��� ]�@DistributeColsWidth
'[ �T  �v ]�@�I�����ꂽ��̕��𑵂���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub DistributeColsWidth()
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objRange     As Range
    Dim lngColCount  As Long    '�I�����ꂽ��̐�
    Dim dblWidth     As Double  '�I�����ꂽ��̕��̍��v
    Dim objSelection As Range   '�I�����ꂽ���ׂĂ̗�̃R���N�V����
    Dim strSelection As String
    Dim objVisible   As Range   '��Range
    
    '�I������Ă���I�u�W�F�N�g�𔻒�
    Select Case CheckSelection()
    Case E_Other
        Exit Sub
    Case E_Shape
        Call DistributeShapeWidth
        Exit Sub
    End Select
    
    '�I��͈͂�Columns�̘a�W�������d�����r������
    strSelection = Selection.Address
    Set objSelection = Union(Selection.EntireColumn, Selection.EntireColumn)
    
    '�s�N�Z������ݒ肷��
    Call SetPixelInfo
    
    '�I��͈͂̉���������o��
    Set objVisible = GetVisibleCells(objSelection)
    
    '���ׂĔ�\���̎�
    If objVisible Is Nothing Then
        Exit Sub
    End If
    
    '��\���̗񂪂��鎞
    If objSelection.Address <> objVisible.Address Then
        If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
            Select Case MsgBox("��\���̗��ΏۂƂ��܂����H", vbYesNoCancel + vbQuestion + vbDefaultButton2)
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
        dblWidth = dblWidth + GetWidth(objRange)
        lngColCount = lngColCount + objRange.Columns.Count
    Next objRange
    
    If lngColCount = 1 Then
        Exit Sub
    End If
    
    '***********************************************
    '�T�C�Y�̕ύX
    '***********************************************
    Application.ScreenUpdating = False
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Call SaveUndoInfo(E_ColSize, Selection, GetSameWidthAddresses(objSelection))
    objSelection.ColumnWidth = PixelToWidth(dblWidth / 0.75 / lngColCount)
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@DistributeShapeWidth
'[ �T  �v ]�@�I�����ꂽ�}�`�̕��𑵂���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub DistributeShapeWidth()
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
    
    Call DistributeShapesWidth(objGroups)
    
    '��]���Ă���}�`�̃O���[�v�������������̐}�`��I������
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@DistributeShapesWidth
'[ �T  �v ]�@�}�`�̕��𑵂���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub DistributeShapesWidth(ByRef objShapeRange As ShapeRange)
    Dim objShape   As Shape
    Dim dblWidth   As Double
    
    '�}�`�̐��������[�v
    For Each objShape In objShapeRange
        dblWidth = dblWidth + objShape.Width
    Next objShape
    With objShapeRange
        .Width = Round(dblWidth / .Count / 0.75) * 0.75
    End With
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetWidth
'[ �T  �v ]�@�I���G���A�̕����擾
'            Width�v���p�e�B��32767�ȏ�̕����v�Z�o���Ȃ�����
'[ ��  �� ]�@�����擾����G���A
'[ �߂�l ]�@��(Width�v���p�e�B)
'*****************************************************************************
Private Function GetWidth(ByRef objRange As Range) As Double
'    With objRange
'        GetWidth = .Columns(.Columns.Count).Left - .Left + .Columns(.Columns.Count).Width
'    End With
    
    Dim lngCount   As Long
    Dim lngHalf    As Long
    Dim MaxWidth   As Double '���̍ő�l

    MaxWidth = 32767 * 0.75
    If objRange.Width < MaxWidth Then
        GetWidth = objRange.Width
    Else
        With objRange
            '�O���{�㔼�̕������v
            lngCount = .Columns.Count
            lngHalf = lngCount / 2
            GetWidth = GetWidth(Range(.Columns(1), .Columns(lngHalf))) + _
                       GetWidth(Range(.Columns(lngHalf + 1), .Columns(lngCount)))
        End With
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@MergeCellsAsColumn
'[ �T  �v ]�@�������Ɍ���(������ɒl�����鎞�͋󔒂ŘA��)
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
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
        '�s�̐��������[�v
        For i = 1 To objRange.Rows.Count
            
            '�����̃Z���ɒl�����鎞�A������
            Set objMergeCell = objRange.Rows(i)
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
'[ �֐��� ]�@SplitColumn
'[ �T  �v ]�@��𕪊�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub SplitColumn()
On Error GoTo ErrHandle
    Dim i               As Long
    Dim objRange        As Range
    Dim lngPixel        As Double  '�P��̕�
    Dim lngSplitCount   As Long    '������
    Dim blnCheckInsert  As Boolean
    Dim objNewCol       As Range   '�V������
    Dim objNewSelection As Range   '������̑I��͈�
    
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
    
    '������I���Ȃ�ΏۊO
    If objRange.Columns.Count <> 1 Then
        Call MsgBox("���̃R�}���h�͕����̑I���ɑ΂��Ď��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    '�s�N�Z������ݒ肷��
    Call SetPixelInfo
    
    '���̕�
    lngPixel = objRange.EntireColumn.Width / 0.75
    
    '****************************************
    '��������I��������
    '****************************************
    With frmSplitCount
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
    Call SaveUndoInfo(E_SplitCol, objRange, lngSplitCount)
    If blnCheckInsert = False Then
        Call SetPlacement
    End If
    
    '�I���̉E���ɂP��}��
    Call objRange(1, 2).EntireColumn.Insert
    
    '�V������
    Set objNewCol = objRange(1, 2).EntireColumn
    
    '*************************************************
    '�r���𐮂���
    '*************************************************
    Dim strMsg  As String
    '�}����̂P�Z�����Ɍr�����R�s�[����
    i = CopyBorders(False, objNewCol)
    If i <> 0 Then
        strMsg = CStr(i + 1) & "�s�ڈȍ~�̃Z���̌r���͐������Ă��܂���B" & vbCrLf
        strMsg = strMsg & "���ʏ�͂��̃��b�Z�[�W�͕\������܂���"
    End If
    
    '���E��������
    With Range(objRange, objNewCol).EntireColumn
        .Borders(xlInsideVertical).LineStyle = xlNone
    End With
    
    '*************************************************
    '�������J�Ԃ�
    '*************************************************
    '�����������A���}������
    For i = 3 To lngSplitCount
        Call objNewCol.EntireColumn.Insert
    Next i
    
    Set objNewSelection = objRange.Resize(, lngSplitCount)
    
    '*************************************************
    '���̐���
    '*************************************************
    '�V�������ɐݒ�
    If blnCheckInsert = False Then
        If lngSplitCount = 2 Then
            objRange.EntireColumn.ColumnWidth = PixelToWidth(Int(lngPixel / 2 + 0.5))
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
    '�㏈��
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
    If strMsg <> "" Then
        Call MsgBox(strMsg, vbExclamation)
    End If
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
'[ �֐��� ]�@EraseColumn
'[ �T  �v ]�@�I�����ꂽ�����������
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub EraseColumn()
On Error GoTo ErrHandle
    Dim i                 As Long
    Dim objSelection      As Range
    Dim strSelection      As String
    Dim objRange          As Range
    Dim lngLeftCol        As Long  '���ׂ̗�ԍ�
    Dim lngRightCol       As Long  '�E�ׂ̗�ԍ�
    Dim objRow            As Range '�I�����ɑI��������s
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = Selection
    strSelection = objSelection.Address
    Set objRange = objSelection.EntireColumn
    
    '�I�����ɑI��������s
    Set objRow = objSelection.EntireRow
    
    '�I���G���A�������Ȃ�ΏۊO
    If objSelection.Areas.Count <> 1 Then
        Call MsgBox("���̃R�}���h�͕����̑I��͈͂ɑ΂��Ď��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    If objSelection.Columns.Count = Columns.Count Then
        Call MsgBox("���̃R�}���h�͂��ׂĂ̗�̑I���ɑ΂��Ď��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    '���ׂ̗�ԍ�
    lngLeftCol = objRange.Column - 1
    '�E�ׂ̗�ԍ�
    lngRightCol = objRange.Column + objRange.Columns.Count
    
    '****************************************
    '�����̃p�^�[����I��������
    '****************************************
    Dim enmSelectType As ESelectType  '�����p�^�[��
    Dim blnHidden   As Boolean        '��\���Ƃ��邩�ǂ���
    With frmEraseSelect
        '�V�[�g�̂P��ڂ��I������Ă��鎞
        If lngLeftCol = 0 Then
            Call .SetEnabled(E_Back)
            Call .SetEnabled(E_Middle)
            .SelectType = E_Front
        '�V�[�g�̍ŏI�񂪑I������Ă��鎞
        ElseIf lngRightCol > Columns.Count Then
            Call .SetEnabled(E_Front)
            Call .SetEnabled(E_Middle)
            .SelectType = E_Back
        Else
            .SelectType = E_Back
        End If
        
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
    '�l�̂������폜���邩�ǂ�������
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
    'Undo�p�ɗ񕝂�ۑ����邽�߂̏��擾
    '****************************************
    Dim colAddress   As New Collection
    Set colAddress = GetSameWidthAddresses(Range(strSelection).EntireColumn)
    
    '****************************************
    '�I���̍��E�̗��Ҕ�
    '****************************************
    Dim objCol(0 To 1)   As Range  '��̕���ς����
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
    '�������
    '****************************************
    Dim strMsg    As String
    Dim lngPixel  As Long   '����������̕�(�P��:�s�N�Z��)
    Application.ScreenUpdating = False
    
    '�s�N�Z������ݒ肷��
    Call SetPixelInfo
    
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    If blnHidden = True Then
        Call SaveUndoInfo(E_ColSize, Range(strSelection), colAddress)
    Else
        Call SaveUndoInfo(E_EraseCol, Range(strSelection), colAddress)
    End If
    
    '�}�`�͈ړ������Ȃ�
    Call SetPlacement
    
    '����������̕���ۑ�
    lngPixel = objRange.Width / 0.75
    
    If blnHidden = True Then
        '��\��
        objRange.Hidden = True
    Else
        i = CopyBorder(objRange)
        If i <> 0 Then
            strMsg = CStr(i + 1) & "�s�ڈȍ~�̃Z���̌r���͐������Ă��܂���B" & vbCrLf
            strMsg = strMsg & "���ʏ�͂��̃��b�Z�[�W�͕\������܂���"
        End If
        '�폜
        Call objRange.Delete(xlShiftToLeft)
    End If
    
    '****************************************
    '�񕝂��g��
    '****************************************
    Dim lngWkPixel  As Long
    Select Case enmSelectType
    Case E_Front, E_Back
        objCol(0).ColumnWidth = WorksheetFunction.Min(MaxColumnWidth, PixelToWidth(objCol(0).Width / 0.75 + lngPixel))
    Case E_Middle
        objCol(0).ColumnWidth = WorksheetFunction.Min(MaxColumnWidth, PixelToWidth(objCol(0).Width / 0.75 + Int(lngPixel / 2 + 0.5)))
        objCol(1).ColumnWidth = WorksheetFunction.Min(MaxColumnWidth, PixelToWidth(objCol(1).Width / 0.75 + Int(lngPixel / 2)))
    End Select
    
    '****************************************
    '�Z����I��
    '****************************************
    Select Case enmSelectType
    Case E_Front, E_Back
        IntersectRange(objRow, objCol(0)).Select
    Case E_Middle
        IntersectRange(objRow, UnionRange(objCol(0), objCol(1))).Select
    End Select
    Call ResetPlacement
    Call SetOnUndo
    
    Application.ScreenUpdating = True
    If strMsg <> "" Then
        Call MsgBox(strMsg, vbExclamation)
    End If
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
    Call ResetPlacement
End Sub

'*****************************************************************************
'[ �֐��� ]�@CopyBorder
'[ �T  �v ]�@�V������Ɍ��̗�̌r�����R�s�[����
'[ ��  �� ]�@objRange:�폜��
'[ �߂�l ]�@0:�����A1�ȏ�:�����𒆒f�����s
'*****************************************************************************
Private Function CopyBorder(ByRef objRange As Range) As Long
    If objRange.Column > 1 Then
        '�P�Z�����Ɍr�����R�s�[����
        With objRange
            CopyBorder = CopyRightBorder(.Columns(.Columns.Count), .Columns(0))
        End With
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@CopyBorders
'[ �T  �v ]�@�V������Ɍ��̗�̌r�����R�s�[����
'[ ��  �� ]�@blnMerge:���ׂĂ̗��O�̗�ƌ������邩�ǂ����AobjRange:�}����
'[ �߂�l ]�@0:�����A1�ȏ�:�����𒆒f�����s
'*****************************************************************************
Private Function CopyBorders(ByVal blnMerge As Boolean, ByRef objRange As Range) As Long
    Dim i                 As Long
    Dim objCell           As Range   '�R�s�[��̃Z��
    Dim objOrgCell        As Range   '�R�s�[���̃Z��
    Dim udtBorder(0 To 2) As TBorder '�r���̎��(�㥉���E)
    Dim lngLast    As Long
    Dim sinTime    As Single  '�o�ߎ��Ԍv�Z�p
    
    sinTime = Timer()
    '���ׂĂ̍s�𐮔����邩�A�Ō�̃Z���܂Ő������邩�A2000�s�܂Ő�������ΏI������
    lngLast = WorksheetFunction.Min(Cells.SpecialCells(xlCellTypeLastCell).Row, _
                              objRange.Row + objRange.Rows.Count - 1, 2000)
    '1�s���Ƀ��[�v
    For i = 1 To lngLast
        Set objCell = objRange.Rows(i)
        
        '�V������̃Z���������Z����
        If objCell.MergeCells = False Then
            '���̃Z������
            Set objOrgCell = objCell.Offset(0, -1)
            
            '�r���̎�ނ�ۑ�(�㥉���E)
            With objOrgCell.MergeArea.Borders
                udtBorder(0) = GetBorder(.Item(xlEdgeTop))
                udtBorder(1) = GetBorder(.Item(xlEdgeBottom))
                udtBorder(2) = GetBorder(.Item(xlEdgeRight))
            End With
            
            '���̃Z���ƌ����Z����
            If blnMerge = True Or objOrgCell.MergeCells = True Then
                With objOrgCell.MergeArea
                    '�}����̃Z��������
                    Call .Resize(, .Columns.Count + 1).Merge
                End With
            End If
            
            With objCell.MergeArea.Borders
                '�r��������(�㥉���E)
                Call SetBorder(udtBorder(0), .Item(xlEdgeTop))
                Call SetBorder(udtBorder(1), .Item(xlEdgeBottom))
                Call SetBorder(udtBorder(2), .Item(xlEdgeRight))
            End With
            
            '5�b�ȏ�o�߂�����
            If Timer() - sinTime > 5 Then
                CopyBorders = objCell.Row
                Exit Function
            End If
        End If
    Next i
    
    If i = 2001 And Cells.SpecialCells(xlCellTypeLastCell).Row <> 2000 Then
        CopyBorders = 2000
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@CopyRightBorder
'[ �T  �v ]�@�폜��̉E�[�̌r�����R�s�[����
'[ ��  �� ]�@objFromRange�F�R�s�|���̗�
'�@�@�@�@�@�@objToRange�F�R�s�[��̗�
'[ �߂�l ]�@0:�����A1�ȏ�:�����𒆒f�����s
'*****************************************************************************
Private Function CopyRightBorder(ByRef objFromRange As Range, ByRef objToRange As Range) As Long
    Dim i          As Long
    Dim udtBorder  As TBorder '�E�[�̌r���̎��
    Dim lngLast    As Long
    Dim sinTime    As Single  '�o�ߎ��Ԍv�Z�p
    sinTime = Timer()
    '���ׂĂ̍s�𐮔����邩�A�Ō�̃Z���܂Ő������邩�A2000�s�܂Ő�������ΏI������
    lngLast = WorksheetFunction.Min(Cells.SpecialCells(xlCellTypeLastCell).Row, _
                              objFromRange.Row + objFromRange.Rows.Count - 1, 2000)
    '1�s���Ƀ��[�v
    For i = 1 To lngLast
        '�R�s�|���̃Z���̌r���̎�ނ�ۑ�
        udtBorder = GetBorder(objFromRange.Rows(i).Borders(xlEdgeRight))
        '�R�s�[��̃Z���ɃR�s�[
        Call SetBorder(udtBorder, objToRange.Rows(i).Borders(xlEdgeRight))
        
        '5�b�ȏ�o�߂�����
        If Timer() - sinTime > 5 Then
            CopyRightBorder = objFromRange.Rows(i).Row
            Exit Function
        End If
    Next i
    If i = 2001 And Cells.SpecialCells(xlCellTypeLastCell).Row <> 2000 Then
        CopyRightBorder = 2000
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@ShowWidth
'[ �T  �v ]�@�G���A���ɑI�����ꂽ��̕����ꗗ�ŕ\������
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ShowWidth()
On Error GoTo ErrHandle
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    '�s�N�Z������ݒ肷��
    Call SetPixelInfo
    
    Call frmSizeList.Initialize(E_Col)
    Call frmSizeList.Show
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@MoveCellBorder
'[ �T  �v ]�@��̋��E�̃Z�������E�Ɉړ�����
'[ ��  �� ]�@-1:���Ɉړ��A1:�E�Ɉړ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveCellBorder(ByVal lngLeftRight As Long)
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objSelection As Range
    Dim objWkRange   As Range
    Dim lngColCount  As Long  '�I��̈�̗�
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
    lngColCount = objSelection.Columns.Count
    Dim objChkRange(0 To 2) As Range

    If objSelection.Columns.Count = 1 Then
        Exit Sub
    End If

    '���Ɉړ����鎞
    If lngLeftRight < 0 Then
        Set objChkRange(1) = ArrangeRange(Range(objSelection.Columns(1), objSelection.Columns(2)))
        Set objChkRange(2) = ArrangeRange(objSelection.Columns(2))
    Else '�E�Ɉړ����鎞
        Set objChkRange(1) = ArrangeRange(Range(objSelection.Columns(lngColCount - 1), objSelection.Columns(lngColCount)))
        Set objChkRange(2) = ArrangeRange(objSelection.Columns(lngColCount - 1))
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
    Call SaveUndoInfo(E_MergeCell, objSelection)
    
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
    If lngLeftRight < 0 Then
        '���Ɉړ�����
        Call CopyRightBorder(objWkRange.Columns(2), objWkRange.Columns(1))
        Call objWkRange.Columns(2).Delete(xlToLeft)
        Call CopyBorders(True, objWkRange.Columns(lngColCount))
    Else
        '�E�Ɉړ�����
        Call CopyRightBorder(objWkRange.Columns(lngColCount), objWkRange.Columns(lngColCount - 1))
        Call objWkRange.Columns(lngColCount).Delete(xlToLeft)
        Call objWkRange.Columns(2).Insert(xlToRight)
        Call CopyBorders(True, objWkRange.Columns(2))
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
'[ �֐��� ]�@AutoColFit
'[ �T  �v ]�@��̕��𕶎���̒����ɂ��킹��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub AutoColFit()
On Error GoTo ErrHandle
    Dim objSelection As Range
    Dim objWorkRange As Range
    Dim colColumns   As Collection '���g�͗��Range
    Dim i As Long
    Dim strErrMsg As String
    
    '�I������Ă���I�u�W�F�N�g�𔻒�
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = Selection
    
    If objSelection.Columns.Count = Columns.Count Then
        Call MsgBox("���̃R�}���h�͂��ׂĂ̗�̑I�����͎��s�ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    '���ׂĂ̍s���I������āA������̌����Z�����܂ގ�
    If IsBorderMerged(objSelection) Then
        Call MsgBox("�������ꂽ�Z���̈ꕔ��I�����邱�Ƃ͂ł��܂���B", vbExclamation)
        Exit Sub
    End If
    
    If WorksheetFunction.CountA(objSelection) = 0 Then
        Exit Sub
    End If
    
    '��\���̗��ΏۂƂ��邩�m�F�H
    Dim objVisible    As Range
    Dim objNonVisible As Range
    Set objVisible = GetVisibleCells(objSelection)
    Set objNonVisible = MinusRange(objSelection, objVisible)
    If Not (objNonVisible Is Nothing) Then
        If WorksheetFunction.CountA(objNonVisible) > 0 Then
            If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
                Select Case MsgBox("��\���̃Z����ΏۂƂ��܂����H", vbYesNoCancel + vbQuestion + vbDefaultButton2)
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
        
        '�s�N�Z������ݒ肷��
        Call SetPixelInfo
        
        '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
        Call SaveUndoInfo(E_ColSize, objSelection, GetSameWidthAddresses(Union(objSelection.EntireColumn, objSelection.EntireColumn)))
        
        Call objSelection.Columns.AutoFit
        
        Call SetOnUndo
        Exit Sub
    End If

    '�������̌����Z���́A��Ԍ��������L���Z�����擾
    Set objWorkRange = GetDoRange(objSelection, colColumns)
    If objWorkRange Is Nothing Then
        Call MsgBox("�������̌��������ꂳ��Ă��Ȃ����߁A���s�ł��܂���", vbExclamation)
        Exit Sub
    End If
    
    If Not (MinusRange(objSelection, objWorkRange) Is Nothing) Then
        Call objWorkRange.Select
        If MsgBox("���ݑI�����Ă���Z����ΏۂƂ��Ď��s���܂��B" & vbLf & "��낵���ł����H", vbOKCancel + vbQuestion) = vbCancel Then
            Exit Sub
        End If
    End If
    
    If WorksheetFunction.CountA(objWorkRange) = 0 Then
        Exit Sub
    End If
    
    '***********************************************
    '���s
    '***********************************************
    Dim objUndoColumn As Range
    
    Application.ScreenUpdating = False
    
    '�s�N�Z������ݒ肷��
    Call SetPixelInfo
    
    '�I��͈͂�Columns�̘a�W�������d�����r������
    Set objUndoColumn = Union(objWorkRange.EntireColumn, objWorkRange.EntireColumn)
    
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Call SaveUndoInfo(E_ColSize, objSelection, GetSameWidthAddresses(objUndoColumn))
    
    '�񖈂Ƀ��[�v
    For i = 1 To colColumns.Count - 1
        Call SetColumnWidth(colColumns(i))
    Next
    
    '�������̌������Ȃ��Z�����ꊇ�ݒ�
    If Not (colColumns(colColumns.Count) Is Nothing) Then
        Call colColumns(colColumns.Count).Columns.AutoFit
    End If
    
    Call SetOnUndo
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetDoRange
'[ �T  �v ]�@�������̌����Z�������鎞�A��Ԍ��������L���Z���̂ݎ��s�ΏۂƂ���
'[ ��  �� ]�@objSelection�F�I�����ꂽ�Z��
'            colColumns�F���s�Ώۂ�񖈂Ɏ擾�����A�񖈂�Range�̔z��(�߂�l)
'[ �߂�l ]�@���s�Ώۂ̃Z��
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
    
    '�I��͈͂�Columns�̘a�W�������d�����r������
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
    
    '��ԍŌ�ɂ́A�������̌������Ȃ��Z����ݒ�
    Call colColumns.Add(objSingleCols)
End Function

'*****************************************************************************
'[ �֐��� ]�@GetMergeCol
'[ �T  �v ]�@�������̌����̍ő啝�̗���擾����
'[ ��  �� ]�@���������A�I�����ꂽ��
'[ �߂�l ]�@�ő啝�̗�
'*****************************************************************************
Private Function GetMergeCol(ByRef objColumn As Range, ByRef objSelection As Range) As Range
    Dim i          As Long
    Dim objRange   As Range
    Dim objWkRange As Range
    
    Set objWkRange = ArrangeRange(IntersectRange(objColumn, objSelection))
    
    '�I��͈͂�Columns�̘a�W�������d�����r������
    Set objRange = Union(objWkRange.EntireColumn, objWkRange.EntireColumn)

    '���ꖳ�����[�v�ɂȂ�Ȃ��悤�Ƀ��[�v�񐔂ɏ����^����
    For i = 1 To Columns.Count
        Set objWkRange = ArrangeRange(IntersectRange(objRange, objSelection))
        
        '�I��͈͂�Columns�̘a�W�������d�����r������
        Set GetMergeCol = Union(objWkRange.EntireColumn, objWkRange.EntireColumn)
            
        If GetMergeCol.Address = objRange.Address Then
            Exit Function
        End If
        Set objRange = GetMergeCol
    Next i
    
    '�������[�v�ɂ������鎞
    Call Err.Raise(C_CheckErrMsg, , "�������̌��������ꂳ��Ă��Ȃ����߁A���s�ł��܂���")
End Function

'*****************************************************************************
'[ �֐��� ]�@SetColumnWidth
'[ �T  �v ]�@��̕���ݒ�
'[ ��  �� ]�@�Ώۂ̗�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub SetColumnWidth(ByRef objColumns As Range)
    Dim colAddress As Collection
    Dim i As Long
    Dim lngPixel    As Long
    Dim lngOldPixel As Long
    Dim lngNewPixel As Long
    
    lngOldPixel = objColumns.EntireColumn.Width / 0.75
    lngNewPixel = GetNewPixel(objColumns)
    
    '�I��͈͂�Columns�̘a�W�������d�����r�����A�������̃Z�����擾����
    Set colAddress = GetSameWidthAddresses(Union(objColumns.EntireColumn, objColumns.EntireColumn))
    
    '�������̉򂲂Ƃɕ���ݒ肷��
    For i = 1 To colAddress.Count
        With Range(colAddress(i))
            If lngOldPixel = 0 Then
                .ColumnWidth = PixelToWidth(lngNewPixel / .Columns.Count)
            Else
                lngPixel = .Width / 0.75 * lngNewPixel / lngOldPixel
                .ColumnWidth = PixelToWidth(lngPixel / .Columns.Count)
            End If
        End With
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetNewPixel
'[ �T  �v ]�@WorkSheet�𗘗p���A������̕����擾
'[ ��  �� ]�@�Ώۂ̗�
'[ �߂�l ]�@�V������
'*****************************************************************************
Private Function GetNewPixel(ByRef objColumns As Range) As Long
On Error GoTo ErrHandle
    Dim objWorksheet As Worksheet
        
    Set objWorksheet = ThisWorkbook.Worksheets("Workarea1")
    Call DeleteSheet(objWorksheet)
    objWorksheet.Columns(1).ColumnWidth = PixelToWidth(objColumns.Width / 0.75)
    
    'Workarea1�V�[�g�ɑΏۃZ�����R�s�[
    Call objColumns.Copy(objWorksheet.Cells(1, 1))
            
    '�擪��̌���������
    Call objWorksheet.Columns(1).UnMerge
        
    '�Z���Q�Ƃ������#ERR�ƂȂ�P�[�X�����邽�ߒl���R�s�[����
    If objWorksheet.Columns(1).HasFormula = False Then
    Else
        Call objColumns.Copy
        Call objWorksheet.Cells(1, 1).PasteSpecial(xlPasteValues)
        Application.CutCopyMode = False
    End If
    
    '�擪��̕���ݒ�
    Call objWorksheet.Columns(1).AutoFit
    
    GetNewPixel = objWorksheet.Columns(1).Width / 0.75
ErrHandle:
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
End Function

'*****************************************************************************
'[ �֐��� ]�@PixelToWidth
'[ �T  �v ]�@���̒P�ʂ�ϊ�
'[ ��  �� ]�@lngPixel : ��(�P��:�s�N�Z��)
'[ �߂�l ]�@Width
'*****************************************************************************
Public Function PixelToWidth(ByVal lngPixel As Long) As Double
    If lngPixel <= x1 Then
        PixelToWidth = lngPixel / x1
    Else
        PixelToWidth = (lngPixel - x1) / (x2 - x1) + 1
    End If
    PixelToWidth = WorksheetFunction.RoundDown(PixelToWidth, 3)
End Function

'*****************************************************************************
'[ �֐��� ]�@WidthToPixel
'[ �T  �v ]�@���̒P�ʂ�ϊ�
'[ ��  �� ]�@dblColumnWidth : ��
'[ �߂�l ]�@��(�P��:�s�N�Z��)
'*****************************************************************************
Private Function WidthToPixel(ByVal dblColumnWidth As Double) As Long
    If dblColumnWidth <= 1 Then
        WidthToPixel = dblColumnWidth * x1
    Else
        WidthToPixel = (dblColumnWidth - 1) * (x2 - x1) + x1
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@SetPixelInfo
'[ �T  �v ]�@�W���X�^�C���̃t�H���g��1������2�����̃s�N�Z�������߂�
'            x1�F1�����̃s�N�Z���Ax2�F2�����̃s�N�Z��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetPixelInfo()
On Error GoTo ErrHandle
    Static udtFont As TFont
    Dim objWorkbook As Workbook
        
    Set objWorkbook = ActiveWorkbook
        
    '�W���X�^�C���̃t�H���g���ύX���ꂽ������
    With ActiveWorkbook.Styles("Normal").Font
        If udtFont.Name = .Name And udtFont.Size = .Size And _
           udtFont.Bold = .Bold And udtFont.Italic = .Italic Then
            Exit Sub
        Else
            '�t�H���g����ۑ�����
            udtFont.Name = .Name
            udtFont.Size = .Size
            udtFont.Bold = .Bold
            udtFont.Italic = .Italic
        End If
    End With
    
    '�A�N�e�B�u�ȃu�b�N�����W���X�^�C���̃t�H���g��ύX�o���Ȃ�����
    Call ThisWorkbook.Activate
    
    '�}�N���̃u�b�N�̕W���X�^�C���̃t�H���g��ύX
    Dim blnScreenUpdating  As Boolean
    blnScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    With ThisWorkbook.Styles("Normal").Font
        .Name = udtFont.Name
        .Size = udtFont.Size
        .Bold = udtFont.Bold
        .Italic = udtFont.Italic
    End With
    Application.ScreenUpdating = blnScreenUpdating
    
    '�T�C�Y����ۑ�����
    With ThisWorkbook.Worksheets("Commands")
        x1 = .Range("J:J").Width / 0.75
        x2 = .Range("K:K").Width / 0.75
    End With
    
    Call objWorkbook.Activate
Exit Sub
ErrHandle:
    x1 = 13
    x2 = 21
End Sub

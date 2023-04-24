Attribute VB_Name = "BookMark"
Option Explicit
Option Private Module

Public Const C_PatternColor = &H800080  '���ł��ǂ�(���߂̖�肾��)
Public FBMarkColor As Long
Public FFillColor As Long '�h��Ԃ��F
    
'*****************************************************************************
'[�T�v] �I���Z��(�܂��͐}�`)�̓h��Ԃ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub FillColor()
On Error GoTo ErrHandle
    Select Case CheckSelection()
    Case E_Range
        Dim objSelection As Range
        Set objSelection = Selection
        '�A���h�D�p�Ɍ��̏���ۑ�����
        Call SaveUndoInfo(E_FillRange, GetAddress(Selection))
        Call objSelection.Select '�I��͈͂�߂�(�Ȃ����S�I������鎖�ۂ��N���邱�Ƃ�����)
    Case E_Shape
        Dim objShapeRange As ShapeRange
        Set objShapeRange = HasInteriorShapes(Selection.ShapeRange)
        If objShapeRange Is Nothing Then
            Exit Sub
        End If
        '�A���h�D�p�Ɍ��̏���ۑ�����
        Call SaveUndoInfo(E_FillShape, objShapeRange)
        Call objShapeRange.Select
        '�h��Ԃ�����ƂȂ������݂��鎞�́A���̂܂܂ł͓h��Ԃ��Ȃ��̐}�`�ɐF�����Ȃ��̂�
        '�������񂷂ׂĂ̐}�`���N���A����
        Selection.Interior.ColorIndex = xlNone
    End Select
    
    Selection.Interior.Color = FFillColor
    Call SetOnUndo
ErrHandle:
    Call GetRibbonUI.InvalidateControl("B631")
End Sub

'*****************************************************************************
'[�T�v] ShapeRange�̂���Interior���L����ShapeRange�̂ݕԂ�
'[����] ShapeRange
'[�ߒl] Interior������ShapeRange
'*****************************************************************************
Private Function HasInteriorShapes(ByRef objShapeRange As ShapeRange) As ShapeRange
    Dim i As Long, j As Long
    Dim Dummy
    ReDim lngIDArray(1 To objShapeRange.Count) As Variant
    
    '�}�`�̐��������[�v
    For j = 1 To objShapeRange.Count 'For each�\������Excel2007�Ō^�Ⴂ�ƂȂ�(���Ԃ�o�O)
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
'[�T�v] �I������Z����Bookmark��ݒ�/��������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SetBookmark()
On Error GoTo ErrHandle
    Dim objRange As Range
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If TypeOf Selection Is Range Then
        Set objRange = Selection
    Else
        Exit Sub
    End If
        
    With objRange.Cells(1).Interior
        If .Pattern = xlSolid And _
           .PatternColor = C_PatternColor Then
            '�����̃N���A
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
'[�T�v] ����Bookmark�Ɉړ�
'[����] ��������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub NextOrPrevBookmark(ByVal xlDirection As XlSearchDirection)
    '[Shift]�܂���[Ctrl]Key����������Ă���΁A�t�����Ɍ���
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
'[�T�v] ����Bookmark�Ɉړ�
'[����] ��������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub JumpBookmark(ByVal xlDirection As XlSearchDirection)
On Error GoTo ErrHandle
    Dim objCell      As Range
    Dim objNextCell  As Range
    Dim objSheetCell As Range
    
    Call SetFindFormat
    
    '****************************************
    '�A�N�e�B�u�V�[�g���̌���
    '****************************************
    Dim blnFind  As Boolean
    Set objCell = ActiveCell
    Set objNextCell = FindNextFormat(objCell, xlDirection)
    If (objNextCell Is Nothing) Then
        '���̃V�[�g��ΏۂƂ��邩�ǂ���
        If GetTmpControl("C2").State = False Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    Else
        Set objSheetCell = objNextCell
        If TypeOf Selection Is Range Then
            '���̃V�[�g��ΏۂƂ��邩�ǂ���
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
    '�ׂ̃V�[�g�̌���
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
'[�T�v] Bookmark�����p�̃Z��������ݒ肷��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SetFindFormat()
    Application.FindFormat.Clear
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColor = C_PatternColor
    End With
    
    If TypeOf Selection Is Range Then
        '�I������Ă���Z����1����������
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
'[�T�v] ���ׂĂ�Bookmark��I��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
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
'[�T�v] ���ׂĂ�Bookmark���N���A
'[����] �Ȃ�
'[�ߒl] �Ȃ�
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
    
    '�A�N�e�B�u�V�[�g���̃u�b�N�}�[�N���擾
    Set objActiveSheetRange = GetBookmarks(ActiveWorkbook.ActiveSheet)
    If (objActiveSheetRange Is Nothing) Then
        '���̃V�[�g��ΏۂƂ��Ȃ���
        If GetTmpControl("C2").State = False Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    End If
    
    '�A�N�e�B�u�V�[�g�̂ݑΏۂ̎�
    If GetTmpControl("C2").State = False Then
        '�����Z����I�����Ă��鎞
        If Not IsOnlyCell(Selection) Then
            Set objRange = IntersectRange(Selection, objActiveSheetRange)
            If Not (objRange Is Nothing) Then
                If MsgBox("�I��͈͒��� " & objRange.Count & " �̃u�b�N�}�[�N���폜���܂�" & vbLf & "��낵���ł����H", vbOKCancel + vbQuestion) <> vbCancel Then
                    ArrangeRange(objRange).Interior.ColorIndex = xlNone
                End If
                Application.FindFormat.Clear
                Exit Sub
            End If
        End If
        j1 = objActiveSheetRange.Count
    Else
        Set colRange = New Collection
        '���ׂẴu�b�N�}�[�N�̐����v�Z
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
    
    '�I���Z�����P��̃u�b�N�}�[�N�̃Z���̎�
    If IsOnlyCell(Selection) And _
       Not (IntersectRange(Selection, objActiveSheetRange) Is Nothing) Then
        '�I���Z���Ɠ��F�̃u�b�N�}�[�N�̐����v�Z
        Application.FindFormat.Interior.Color = ActiveCell.Interior.Color
        If GetTmpControl("C2").State = False Then
            '���̃V�[�g��ΏۂƂ��Ȃ���
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
    '���s�m�F
    '****************************************
    If j1 = j2 Then
        If MsgBox(j1 & " �̃u�b�N�}�[�N���폜���܂�" & vbLf & "��낵���ł����H", vbOKCancel + vbQuestion) = vbCancel Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    Else
        If MsgBox(j1 & " �̃u�b�N�}�[�N�̂���" & vbLf & "�I�����ꂽ�Z���Ɠ����F�� " & j2 & " �̃u�b�N�}�[�N���폜���܂�" & vbLf & "��낵���ł����H", vbOKCancel + vbQuestion) = vbCancel Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    End If
    
    '****************************************
    '���ׂẴu�b�N�}�[�N���폜
    '****************************************
    Application.ScreenUpdating = False
    
    '�A�N�e�B�u�V�[�g�̂ݑΏۂ̎�
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
'[�T�v] �ΏۃV�[�g�̂��ׂĂ�Bookmark���擾
'[����] �ΏۃV�[�g
'[�ߒl] Bookmark���ݒ肳�ꂽ�Z�����ׂ�
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
'[�T�v] ���̏����̃Z���Ɉړ�
'[����] �����J�n�Z���A��������
'[�ߒl] ���̏����̃Z��
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
'[�T�v] ��������
'[����] ��������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub FindNext()
    FPressKey = 0
    Call FindNextOrPrev(xlNext)
End Sub

'*****************************************************************************
'[�T�v] ��������
'[����] ��������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub FindPrev()
    FPressKey = 0
    Call FindNextOrPrev(xlPrevious)
End Sub

'*****************************************************************************
'[�T�v] ���܂��͑O������
'[����] ��������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub FindNextOrPrev(ByVal xlDirection As XlSearchDirection)
On Error GoTo ErrHandle
    Dim objCell As Range
    
    '[Shift]�܂���[Ctrl]Key����������Ă���΁A�t�����Ɍ���
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
'[�T�v] ��������
'[����] �����J�n�Z���A��������
'[�ߒl] ���̃Z��
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
'        '�������g�̃Z����I������Ӗ��s���ȃo�O�Ή�
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
'[�T�v] �g�p����Ă���Ō�̃Z�����擾����
'[����] �Ώۂ̃V�[�g
'[�ߒl] �Ō�̃Z��
'*****************************************************************************
Private Function GetLastCell(ByRef objSheet As Worksheet) As Range
    Set GetLastCell = objSheet.Cells.SpecialCells(xlLastCell)
End Function

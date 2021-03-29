Attribute VB_Name = "Organize"
Option Explicit
Option Private Module

Private FCount As Long

'*****************************************************************************
'[�T�v] �W���t�H���g�̕ύX
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ChangeNormalFont()
    ActiveWorkbook.Styles("Normal").Font.Name = GetSetting(REGKEY, "KEY", "FontName", DEFAULTFONT)
End Sub

'*****************************************************************************
'[�T�v] �Z���ɕW���t�H���g��K�p
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ApplyNormalFont()
    If CheckSelection() = E_Range Then
        Selection.Font.Name = ActiveWorkbook.Styles("Normal").Font.Name
    End If
End Sub

'*****************************************************************************
'[�T�v] �}�`�ɕW���t�H���g��K�p
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ApplyNormalFontToShape()
    ReDim lngArray(1 To ActiveSheet.Shapes.Count)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim blnShapeSelect As Boolean
    blnShapeSelect = (CheckSelection() = E_Shape)
    
    For i = 1 To ActiveSheet.Shapes.Count
        Select Case ActiveSheet.Shapes(i).Type
        Case msoAutoShape, msoTextBox
            If blnShapeSelect Then
                For k = 1 To Selection.ShapeRange.Count
                    If Selection.ShapeRange(k).Name = ActiveSheet.Shapes(i).Name Then
                        j = j + 1
                        lngArray(j) = i
                        Exit For
                    End If
                Next
            Else
                j = j + 1
                lngArray(j) = i
            End If
        End Select
    Next
    If j = 0 Then
        Exit Sub
    End If
    
    ReDim Preserve lngArray(1 To j)
    With ActiveSheet.Shapes.Range(lngArray).TextFrame2.TextRange.Font
        .NameComplexScript = ActiveWorkbook.Styles("Normal").Font.Name
        .NameFarEast = ActiveWorkbook.Styles("Normal").Font.Name
        .Name = ActiveWorkbook.Styles("Normal").Font.Name
    End With
End Sub

'*****************************************************************************
'[�T�v] �S�ẴV�[�g�̔{����100%�ɂ��āAA1�Z����I��
'       & ���y�[�W�v���r���[������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SelectHomePosition()
    Dim objActiveSheet As Worksheet
    Set objActiveSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    Dim objSheet As Worksheet
    For Each objSheet In Worksheets
        Call objSheet.Activate
        Call ActiveWindow.SmallScroll(, Rows.Count, , Columns.Count)
        Call objSheet.Range("A1").Select
        
        '���y�[�W�v���r���[����
        ActiveWindow.View = xlNormalView
        
        '�{��100%
        ActiveWindow.Zoom = 100
    Next
    
    Call objActiveSheet.Select
End Sub

'*****************************************************************************
'[�T�v] �������G���[�̃Z����I��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SelectErrFormula()
    If CheckSelection <> E_Range Then
        Call Range("A1").Select
    End If
    On Error Resume Next
    Selection.SpecialCells(xlCellTypeFormulas, xlErrors).Select
    If Err.Number <> 0 Then
        Call MsgBox("�������G���[�̃Z���͂���܂���")
        Exit Sub
    End If
End Sub

'*****************************************************************************
'[�T�v] �������G���[�ɂȂ��Ă�������t���������폜
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub DeleteErrFormatConditions()
    Call MsgBox("DeleteErrFormatConditions �H����")
End Sub

'*****************************************************************************
'[�T�v] ���[�U��`�X�^�C�������ׂč폜
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub DeleteUserStyles()
    '�����̃J�E���g
    Dim objStyle  As Style
    FCount = DeleteStyles(ActiveWorkbook, True)
    If FCount = 0 Then
        Call MsgBox("���[�U��`�̃X�^�C���͂���܂���")
        Exit Sub
    End If

    '�z�[���^�u(�X�^�C��)��\��������
    Call GetRibbonUI.ActivateTabMso("TabHome")
    
    '�^�u��؂�ւ��邽�߁A�^�C�}�[���g�p
    Call Application.OnTime(Now(), "DeleteUserStyles2")
End Sub
Private Sub DeleteUserStyles2()
    DoEvents
    
    Dim strMsg As String
    strMsg = FCount & " �� �̃��[�U��`�X�^�C����������܂���" & vbLf
    strMsg = strMsg & "�폜���܂����H"
    Select Case MsgBox(strMsg, vbYesNo)
    Case vbYes
        Call DeleteStyles(ActiveWorkbook)
    End Select
End Sub

'*****************************************************************************
'[�T�v] ���[�U��`�̖��O�����ׂč폜
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub DeleteNameObjects()
    '�����̃J�E���g
    Dim lngCnt As Long
    Dim objName As Name
    lngCnt = DeleteNames(ActiveWorkbook, True)
    If lngCnt = 0 Then
        Call MsgBox("���[�U��`�̖��O�͂���܂���")
        Exit Sub
    End If
    
    Dim strMsg As String
    strMsg = lngCnt & " �� �̖��O��������܂���" & vbLf & vbLf
    strMsg = strMsg & "�폜���܂����H" & vbCrLf
    strMsg = strMsg & "�@�u �͂� �v���� �폜�����s����" & vbCrLf
    strMsg = strMsg & "�@�u�������v���� ���O�̊Ǘ���ʂ�\������"
    Select Case MsgBox(strMsg, vbYesNoCancel + vbQuestion + vbDefaultButton1)
    Case vbYes
        Call DeleteNames(ActiveWorkbook)
    Case vbNo
        Call ActiveWorkbook.Activate
        Call ActiveWorkbook.ActiveSheet.Activate
        Call CommandBars.ExecuteMso("NameManager")
    End Select
End Sub

'*****************************************************************************
'[�T�v] ���[�U�ݒ�̃r���[�����ׂč폜
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub DeleteUserViews()
    '�����̃J�E���g
    Dim lngCnt As Long
    lngCnt = ActiveWorkbook.CustomViews.Count
    If lngCnt = 0 Then
        Call MsgBox("���[�U�ݒ�̃r���[�͂���܂���")
        Exit Sub
    End If
    
    Dim strMsg As String
    strMsg = lngCnt & " �� �̃��[�U�ݒ�̃r���[��������܂���" & vbLf & vbLf
    strMsg = strMsg & "�폜���܂����H" & vbCrLf
    strMsg = strMsg & "�@�u �͂� �v���� �폜�����s����" & vbCrLf
    strMsg = strMsg & "�@�u�������v���� ���[�U�ݒ�̃r���[��\������"
    Select Case MsgBox(strMsg, vbYesNoCancel + vbQuestion + vbDefaultButton1)
    Case vbYes
        Call DeleteViews(ActiveWorkbook)
    Case vbNo
        Call ActiveWorkbook.Activate
        Call ActiveWorkbook.ActiveSheet.Activate
        Call CommandBars.ExecuteMso("ViewCustomViews")
    End Select
End Sub

'*****************************************************************************
'[�T�v] �ʐσ[���̐}�`�̑I��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SelectFlatShapes()
On Error GoTo ErrHandle
    If ActiveSheet.Shapes.Count = 0 Then
        Call MsgBox("�Ώۂ̐}�`�͂���܂���")
        Exit Sub
    End If
    If ActiveWorkbook.DisplayDrawingObjects = xlHide Then
        ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
        Call GetRibbonUI.InvalidateControl("C1")
    End If
    
    ReDim lngArray(1 To ActiveSheet.Shapes.Count)
    Dim objShape As Shape
    Dim i As Long
    Dim j As Long
    For i = 1 To ActiveSheet.Shapes.Count
        Set objShape = ActiveSheet.Shapes(i)
        '�R�����g�̐}�`�͑ΏۊO�Ƃ���
        If ActiveSheet.Shapes(i).Type <> msoComment Then
            '�������ǂ���
            If TypeName(objShape.DrawingObject) = "Line" Then
                If objShape.Width = 0 And objShape.Height = 0 Then
                    j = j + 1
                    lngArray(j) = i
                End If
            Else
                If objShape.Width = 0 Or objShape.Height = 0 Then
                    j = j + 1
                    lngArray(j) = i
                End If
            End If
        End If
    Next
    
    If j = 0 Then
        Call MsgBox("�Ώۂ̐}�`�͂���܂���")
        Call ShowSelectionPane
        Exit Sub
    End If

    On Error Resume Next
    ReDim Preserve lngArray(1 To j)
    Call ActiveSheet.Shapes.Range(lngArray).Select
    
    Dim strMsg As String
    strMsg = j & " �� �̑Ώۂ̐}�`��I�����܂���" & vbLf & vbLf
    strMsg = strMsg & "�s�v�ł����Delete�L�[�ō폜���Ă�������"
    Call MsgBox(strMsg)
    Call ShowSelectionPane
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �r���Ɠ������������̑I��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SelectFlatArrows()
On Error GoTo ErrHandle
    If ActiveSheet.Shapes.Count = 0 Then
        Call MsgBox("�Ώۂ̒����͂���܂���")
        Exit Sub
    End If
    If ActiveWorkbook.DisplayDrawingObjects = xlHide Then
        ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
        Call GetRibbonUI.InvalidateControl("C1")
    End If
    
    ReDim lngArray(1 To ActiveSheet.Shapes.Count)
    Dim objShape As Shape
    Dim i As Long
    Dim j As Long
    For i = 1 To ActiveSheet.Shapes.Count
        Set objShape = ActiveSheet.Shapes(i)
        '����
        If TypeName(objShape.DrawingObject) = "Line" Then
            If objShape.Width = 0 And objShape.Height = 0 Then
                j = j + 1
                lngArray(j) = i
            ElseIf objShape.Width = 0 Then
                If objShape.Left = objShape.TopLeftCell.Left Then
                    j = j + 1
                    lngArray(j) = i
                End If
            ElseIf objShape.Height = 0 Then
                If objShape.Top = objShape.TopLeftCell.Top Then
                    j = j + 1
                    lngArray(j) = i
                End If
            End If
        End If
    Next
    
    If j = 0 Then
        Call MsgBox("�Ώۂ̒����͂���܂���")
        Call ShowSelectionPane
        Exit Sub
    End If

    On Error Resume Next
    ReDim Preserve lngArray(1 To j)
    Call ActiveSheet.Shapes.Range(lngArray).Select

    Dim strMsg As String
    strMsg = j & " �� �̑Ώۂ̒�����I�����܂���" & vbLf & vbLf
    strMsg = strMsg & "�s�v�ł����Delete�L�[�ō폜���Ă�������"
    Call MsgBox(strMsg)
    Call ShowSelectionPane
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] A1�`����R1C1�Q�ƌ`���𑊌݂ɐؑւ���
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ToggleA1R1C1()
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub

'*****************************************************************************
'[�T�v] Workbook���N���A����
'[����] Workbook
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ClearBook(ByRef objWorkbook As Workbook)
    '���O�I�u�W�F�N�g�����ׂč폜����
    Call DeleteNames(ThisWorkbook)
    '�X�^�C�������ׂč폜����
    Call DeleteStyles(ThisWorkbook)
End Sub

'*****************************************************************************
'[�T�v] ���O�I�u�W�F�N�g���폜����
'[����] Workbook, blnCountOnly:�����̃J�E���g�݂̂̎�True
'[�ߒl] �폜�Ώۂ̃I�u�W�F�N�g�̌���
'*****************************************************************************
Private Function DeleteNames(ByRef objWorkbook As Workbook, Optional ByVal blnCountOnly As Boolean = False) As Long
    Dim objName     As Name
    For Each objName In objWorkbook.Names
        Select Case objName.MacroType
        'EXCEL2019�̓�̎��ۑΉ�(TEXTJOIN�֐������g���Ώ���ɖ��O����`����邪�폜����Ɨ�O�ɂȂ�̂ŉ��)
        Case xlFunction, xlCommand, xlNotXLM
        Case Else
            If Right(objName.RefersTo, 5) = "#REF!" Then
                DeleteNames = DeleteNames + 1
                If Not blnCountOnly Then
                    Call objName.Delete
                    DoEvents
                End If
            ElseIf (Right(objName.Name, Len("Print_Area")) <> "Print_Area") And _
               (Right(objName.Name, Len("Print_Titles")) <> "Print_Titles") And _
               objName.Visible Then
                DeleteNames = DeleteNames + 1
                If Not blnCountOnly Then
                    Call objName.Delete
                    DoEvents
                End If
            End If
        End Select
    Next
End Function

'*****************************************************************************
'[�T�v] �X�^�C�����폜����
'[����] Workbook, blnCountOnly:�����̃J�E���g�݂̂̎�True
'[�ߒl] �폜�Ώۂ̃I�u�W�F�N�g�̌���
'*****************************************************************************
Private Function DeleteStyles(ByRef objWorkbook As Workbook, Optional ByVal blnCountOnly As Boolean = False) As Long
    Dim objStyle  As Style
    For Each objStyle In objWorkbook.Styles
        If objStyle.BuiltIn = False Then
            DeleteStyles = DeleteStyles + 1
            If Not blnCountOnly Then
                Call objStyle.Delete
                DoEvents
            End If
        End If
    Next
End Function

'*****************************************************************************
'[�T�v] ���[�U�ݒ�̃r���[���폜����
'[����] Workbook
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub DeleteViews(ByRef objWorkbook As Workbook)
    Dim objView  As CustomView
    For Each objView In objWorkbook.CustomViews
        Call objView.Delete
        DoEvents
    Next
End Sub

'*****************************************************************************
'[�T�v] �g�p���ꂽ�Z���͈̔͂��œK������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub UsedRange()
    ActiveSheet.UsedRange.Select
End Sub


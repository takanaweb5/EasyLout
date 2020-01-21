Attribute VB_Name = "ElseTools"
Option Explicit
Option Private Module

'�r�����
Public Type TBorder
    ColorIndex As Long
    LineStyle  As Long
    Weight     As Long
    Color      As Long
End Type

'�ړ��\���ǂ����̏��
Private Type TPlacement
    Placement As Byte
    Top       As Double
    Height    As Double
    Left      As Double
    Width     As Double
End Type

'���͎x���t�H�[���̏��
Private Type TEditInfo
    Top      As Long
    Left     As Long
    Width    As Long
    Height   As Long
    FontSize As Long
    Zoomed   As Boolean
    WordWarp As Boolean
End Type

Public Const C_CheckErrMsg = 514

'�Ăяo����̃t�H�[�������[�h����Ă��邩�ǂ���
Public blnFormLoad As Boolean

Private clsUndoObject  As CUndoObject  'Undo���
Private lngProcessId As Long   '�w���v�̃v���Z�XID
Private hHelp        As LongPtr   '�w���v�̃n���h��
Private udtPlacement() As TPlacement

Private objRepeatCmd As CommandBarControl

'*****************************************************************************
'[ �֐��� ]�@ColRowChange
'[ �T  �v ]�@�s��ؑւ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ColRowChange()
    Dim i As Long
    Dim strClipText As String
    
    Call SetKeys
    
    'Ctlr + Shift ����������Ă����
    If GetKeyState(vbKeyControl) < 0 Or GetKeyState(vbKeyShift) < 0 Then
        '�I�v�V�����ݒ��ʂ�\��
        Call SetOption
        Exit Sub
    End If
    
    strClipText = GetClipbordText()
    
    With CommandBars("���񂽂񃌃C�A�E�g").Controls(1)
        If .Caption = "�s" Then
            .Caption = "��"
        Else
            .Caption = "�s"
        End If
    End With
    
    Application.ScreenUpdating = False
    On Error Resume Next
    With CommandBars("���񂽂񃌃C�A�E�g").Controls
        For i = 2 To .Count
            Call SetCommand(.Item(1).Caption, .Item(i))
        Next i
    End With
    
    '�N���b�v�{�[�h�̕���
    If strClipText = "" Then
        Call ClearClipbord
    Else
        Call SetClipbordText(strClipText)
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetCommand
'[ �T  �v ]�@�R�}���h�o�[�{�^���ɃA�C�R����R�}���h�TooltipText��ݒ�
'[ ��  �� ]�@strGroup : "�s" or "��" or "���̑�"
'            objCmdBarBtn : �R�}���h��ؑւ���{�^��
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetCommand(ByVal strGroup As String, ByRef objCmdBarBtn As CommandBarButton)
    Dim i As Long
    Dim strCommand As String
    Dim objRange   As Range
'    Dim objMask    As IPictureDisp
    
    strCommand = objCmdBarBtn.Caption
    
    '���[�N�V�|�g�̃R�}���h��ݒ肷��
    Set objRange = ThisWorkbook.Worksheets("Commands").Cells(1, "A").CurrentRegion

    For i = 2 To objRange.Rows.Count
        If objRange(i, "A") = strGroup And objRange(i, "B") = strCommand Then
            With objCmdBarBtn
                '�A�C�R���̐ݒ�
                If objRange(i, "C") <> "" Then
                    .FaceId = objRange(i, "C")
                Else
                    If CopyIconFromHidden(strGroup & "_" & strCommand) = True Then
                        Call .PasteFace
                    End If
'                    Set objMask = Nothing
'                    If CopyIconFromCell(objRange.Cells(i, "E")) = True Then
'                        Call .PasteFace
'                        Set objMask = .Picture
'                    End If
'                    If CopyIconFromCell(objRange.Cells(i, "D")) = True Then
'                        Call .PasteFace
'                        If Not (objMask Is Nothing) Then
'                            .Mask = objMask
'                        End If
'                    End If
'                    .Picture = objRange.Parent.OLEObjects(strPrefix & .Caption).Object.Picture
                End If
                
                '�R�}���h�̐ݒ�
                If objRange(i, "F") = .ID Then
                    .OnAction = ""
                Else
                    .OnAction = objRange(i, "G")
                    .Parameter = objRange(i, "H")
                End If
                
                '�w���v�̐ݒ�
                If Val(Application.Version) >= 12 Then
                    .TooltipText = Replace$(objRange(i, "I"), vbLf, "  ")
                Else
                    .TooltipText = objRange(i, "I")
                End If

'                .Tag = .Caption '�T�C�Y�ꗗ�̏��̕ۑ���Tag���g�p����̂Œ��ӂ��邱��
            End With
                    
            Exit Sub
        End If
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ]�@CopyIconFromCell
'[ �T  �v ]�@�����̃Z���Ɋ܂܂��A�C�R�����N���b�v�{�[�h�ɃR�s�[����
'[ ��  �� ]�@�A�C�R�����܂ރZ��
'[ �߂�l ]�@True:�����AFalse:���s
'*****************************************************************************
'Private Function CopyIconFromCell(ByRef objCell As Range) As Boolean
'    Dim objShape As Shape
'    For Each objShape In objCell.Worksheet.Shapes
'        If objCell.Top = objShape.Top And _
'           objCell.Left = objShape.Left Then
'            Call objShape.CopyPicture(xlScreen, xlBitmap)
'            CopyIconFromCell = True
'            Exit Function
'        End If
'    Next objShape
'End Function

'*****************************************************************************
'[ �֐��� ]�@CopyIconFromHidden
'[ �T  �v ]�@���񂽂񃌃C�A�E�g�A�C�R���̃A�C�R�����N���b�v�{�[�h�ɃR�s�[����
'[ ��  �� ]�@�R�}���h�̖��O�@��F��_�k��
'[ �߂�l ]�@True:�����AFalse:���s
'*****************************************************************************
Private Function CopyIconFromHidden(ByVal strCommand As String) As Boolean
On Error GoTo ErrHandle
    Dim objBtn As CommandBarButton

    For Each objBtn In CommandBars("���񂽂񃌃C�A�E�g�A�C�R��").Controls
        If objBtn.Caption = strCommand Then
            Call objBtn.CopyFace
            CopyIconFromHidden = True
            Exit Function
        End If
    Next objBtn
ErrHandle:
End Function

'*****************************************************************************
'[ �֐��� ]�@OpenHelp
'[ �T  �v ]�@�w���v�t�@�C�����J��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub OpenHelp()
    Call OpenHelpPage("Introduction.htm")
End Sub

'*****************************************************************************
'[ �֐��� ]�@OpenHelpPage
'[ �T  �v ]�@�w���v�t�@�C���̓���̃y�[�W���J��
'[ ��  �� ]�@�y�[�W���
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub OpenHelpPage(ByVal Bookmark As String)
On Error GoTo ErrHandle
    Dim strHelpPath As String
    Dim strMsg As String
    
    strHelpPath = ThisWorkbook.Path & "\" & "EasyLout.chm"
    If Dir(strHelpPath) = "" Then
        strMsg = "�w���v�t�@�C����������܂���B" & vbLf
        strMsg = strMsg & "EasyLout.chm�t�@�C����EasyLout.xla�t�@�C���Ɠ����t�H���_�ɃR�s�[���ĉ������B"
        Call MsgBox(strMsg, vbExclamation)
        Exit Sub
    End If
    
    'Help�����ׂɋN�����Ă��邩�ǂ�������
    If hHelp <> 0 Then
        Dim lngExitCode As Long
        Call GetExitCodeProcess(hHelp, lngExitCode)
        If lngExitCode = STILL_ACTIVE Then
            Call AppActivate(lngProcessId)
            Exit Sub
        End If
    End If
    
    '�w���v�t�@�C�����I�[�v������
    lngProcessId = Shell("hh.exe " & strHelpPath & "::/_RESOURCE/" & Bookmark, vbNormalFocus)
    
    '�v���Z�X�̃n���h�����擾����
    hHelp = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE Or PROCESS_QUERY_INFORMATION, 0&, lngProcessId)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@MergeCell
'[ �T  �v ]�@�Z�����������A�l���󔒂Ɖ��s�łȂ��œ��͂���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MergeCell()
On Error GoTo ErrHandle
    Dim objRange     As Range
    Dim strSelection As String
    Dim lngCalculation As Long
    
    '**************************************
    '��������
    '**************************************
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    strSelection = Selection.Address
    lngCalculation = Application.Calculation
    
    '***********************************************
    '�I���G���A���P�Ō����Z���Ȃ�ΏۊO
    '***********************************************
    With Range(strSelection)
        If .Areas.Count = 1 And IsOnlyCell(.Cells) Then
            Exit Sub
        End If
    End With
    
    '***********************************************
    '�d���̈�̃`�F�b�N
    '***********************************************
    If CheckDupRange(Range(strSelection)) = True Then
        Call MsgBox("�I������Ă���̈�ɏd��������܂�", vbExclamation)
        Exit Sub
    End If
    
    '**************************************
    '�����̃Z���̑��݃`�F�b�N
    '**************************************
    Dim objFormulaCell  As Range
    Dim objFormulaCells As Range
    
    '�G���A�̐��������[�v
    For Each objRange In Range(strSelection).Areas
        If WorksheetFunction.CountA(objRange) > 1 Then
            On Error Resume Next
            Set objFormulaCell = objRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            Set objFormulaCells = UnionRange(objFormulaCells, objFormulaCell)
        End If
    Next objRange
    If Not (objFormulaCells Is Nothing) Then
        Call objFormulaCells.Select
        If MsgBox("�����͒l�ɕϊ�����܂��B��낵���ł����H", vbOKCancel + vbQuestion) = vbCancel Then
            Exit Sub
        Else
            Call Range(strSelection).Select
        End If
    End If
    
    '**************************************
    '���s
    '**************************************
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False '�R�����g�������Z���ɂ��鎞�̑Ή�
    Application.Calculation = xlManual
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    Call SaveUndoInfo(E_MergeCell, Range(strSelection))
    
    '�G���A�̐��������[�v
    For Each objRange In Range(strSelection).Areas
        Call MergeArea(objRange)
    Next objRange
    
    Call SetOnUndo
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
    Call SetOnRepeat
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@MergeArea
'[ �T  �v ]�@�Z�����������A�l���󔒂Ɖ��s�łȂ��œ��͂���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MergeArea(ByRef objRange As Range)
    Dim i As Long
    Dim j As Long
    Dim strValues   As String
    
    '**************************************
    '�l�����͂��ꂽ�Z�����P�ȉ��̎��̏���
    '**************************************
    If WorksheetFunction.CountA(objRange) <= 1 Then
        Call objRange.Merge
        Exit Sub
    End If
    
    '**************************************
    '�e�s�̕���������s�ŁA�e��̕�������󔒂ŋ�؂��ĘA������
    '**************************************
    strValues = Replace$(GetRangeText(objRange), vbTab, " ")
    
    '**************************************
    '�Z�����������Ēl��ݒ�
    '**************************************
    '�����Z�������ׂĉ�������
    Call objRange.UnMerge
    '��U�l������
    Call objRange.ClearContents
    '�Z��������
    Call objRange.Merge
    '������W���ɕύX(�������������255�����𒴂�����\���o���Ȃ�����)
    objRange.NumberFormat = "General"
    objRange.Value = strValues
End Sub

'*****************************************************************************
'[ �֐��� ]�@ParseCell
'[ �T  �v ]�@�Z���̌������������A�l�̂P�s���Z���̂P�s�ɂ���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ParseCell()
On Error GoTo ErrHandle
    Dim lngCalculation As Long
    Dim objRange     As Range

    lngCalculation = Application.Calculation
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    '�s�����̌����Z�����Ȃ���ΑΏۊO
    Set objRange = GetMergeRange(Selection, E_MTROW)
    If objRange Is Nothing Then
        Exit Sub
    End If

    '�s�����̌����Z���������P�̌����Z���̎�
    If objRange.Count = 1 Then
        Call objRange.Select
    End If
    
    '****************************************
    '�㑵���A���������A��������I��������
    '****************************************
    Dim lngAlignment  As Long
    '�f�[�^�̓��͂��ꂽ�Z�������݂��Ȃ���
    If WorksheetFunction.CountA(objRange) = 0 Then
        lngAlignment = E_Top
    Else
        '�ΏۃZ�����P�����̎�
        If objRange.Count = 1 Then
            '���ꂪ�����̎�
            If objRange.HasFormula = True Then
                lngAlignment = E_Top
            Else
                If GetStrArray(objRange.Value) < objRange.MergeArea.Rows.Count Then
                    lngAlignment = InputAlignment()
                Else
                    lngAlignment = E_Top
                End If
            End If
        Else
            lngAlignment = InputAlignment()
        End If
    End If
    '�L�����Z�����ꂽ��
    If lngAlignment = E_Cancel Then
        Exit Sub
    End If
    
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    Call SaveUndoInfo(E_MergeCell, Selection)
    
    '****************************************
    '�l�̂���Z�����P�s���Ƃɕ���
    '****************************************
    Dim objCell As Range
    For Each objCell In objRange
        Call ParseOneValueRange(objCell.MergeArea, lngAlignment)
    Next
    
    Call SetOnUndo
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
    Call SetOnRepeat
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Application.Calculation = lngCalculation
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@InputAlignment
'[ �T  �v ]�@������̏c�����̑�������I��������
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@xlCancel:�L�����Z���AE_Top:�㑵���AE_Center:���������AE_Bottom:������
'*****************************************************************************
Private Function InputAlignment() As Long
    With frmAlignment
        '�t�H�[����\��
        Call .Show

        '�L�����Z����
        If blnFormLoad = False Then
            InputAlignment = E_Cancel
            Exit Function
        End If

        InputAlignment = .SelectType
        Call Unload(frmAlignment)
    End With
End Function

'*****************************************************************************
'[ �֐��� ]�@ParseOneValueRange
'[ �T  �v ]  �l�̂���Z�����P�s���Ƃɕ���
'[ ��  �� ]�@�l�̂��錋������������Z���A�s�����̑���
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ParseOneValueRange(ByRef objRange As Range, ByVal lngAlignment As Long)
    Dim strArray() As String
    Dim lngLine    As Long '�s��
    
    '��������P�s�Âz��Ɋi�[
    lngLine = GetStrArray(objRange(1, 1).Value, strArray())
    
    '�����Z�� �܂���
    '�l���P�s�����ŏ㑵���̎�
    If (objRange(1, 1).HasFormula = True) Or _
       (lngLine = 1 And lngAlignment = E_Top) Then
    Else
        '��U�l������
        Call objRange.ClearContents
    End If
    
    '********************
    '����������
    '********************
    Call objRange.Merge(True)
    '�Z�����E�̌r��������
    objRange.Borders(xlInsideHorizontal).LineStyle = xlNone
    '�u�܂�Ԃ��đS�̂�\������v�̃`�F�b�N���͂���
    objRange.WrapText = False
    
    '�����Z�� �܂��́@�󔒃Z���@�܂���
    '�l���P�s�����ŏ㑵���̎�
    If (objRange(1, 1).HasFormula = True) Or (lngLine = 0) Or _
       (lngLine = 1 And lngAlignment = E_Top) Then
        Exit Sub
    End If
    
    '**************************************
    '�P�s�P�ʂɒl�ݒ�
    '**************************************
    Dim lngStart   As Long
    Dim strLastCell As String
    Dim i  As Long
    Dim j  As Long
    If lngLine < objRange.Rows.Count Then
        Select Case lngAlignment
        Case E_Top
            lngStart = 1
        Case E_Center
            lngStart = Int((objRange.Rows.Count - lngLine) / 2) + 1
        Case E_Bottom
            lngStart = objRange.Rows.Count - lngLine + 1
        End Select
    Else
        lngStart = 1
    End If
    
    '�s�̐��������[�v
    For i = lngStart To objRange.Rows.Count
        With objRange(i, 1)
            If .NumberFormat <> "@" And StrConv(Left$(strArray(j), 1), vbNarrow) = "=" Then
                .Value = "'" & strArray(j)
            Else
                .Value = strArray(j)
            End If
        End With
        j = j + 1
        If j > UBound(strArray) Then
            Exit For
        End If
    Next i
    
    If lngLine > objRange.Rows.Count Then
        strLastCell = strArray(j - 1)
        For i = j To UBound(strArray)
            strLastCell = strLastCell & vbLf & strArray(i)
        Next
        
        With objRange(objRange.Rows.Count, 1)
            If .NumberFormat <> "@" And StrConv(Left$(strLastCell, 1), vbNarrow) = "=" Then
                .Value = "'" & strLastCell
            Else
                .Value = strLastCell
            End If
        End With
    End If
End Sub
    
'*****************************************************************************
'[ �֐��� ]�@PasteValue
'[ �T  �v ]�@�l��\��t����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub PasteValue()
    Call SetKeys
    
    If GetKeyState(vbKeyControl) < 0 Then
        Call CopyText  '�Z���̒l���N���b�v�{�[�h�ɃR�s�[����
    Else
        Call PasteText '�l���Z���ɓ\��t����
    End If
End Sub
    
'*****************************************************************************
'[ �֐��� ]�@PasteText
'[ �T  �v ]�@�l���Z���ɓ\��t����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub PasteText()
On Error GoTo ErrHandle
    Dim strCopyRange  As String
    Dim objSelection  As Range
    Dim strCopyText   As String
    Dim blnOnlyCell   As Boolean
    Dim blnAllCell    As Boolean

    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If

    Set objSelection = Selection

    Select Case Application.CutCopyMode
    Case xlCut
        Call MsgBox("�؂��莞�́A���s�o���܂���B", vbExclamation)
        Exit Sub
    Case xlCopy
        On Error Resume Next
        strCopyRange = GetCopyRangeAddress()
        On Error GoTo 0
        If strCopyRange <> "" Then
            blnOnlyCell = IsOnlyCell(Range(strCopyRange))
            If blnOnlyCell Then
                strCopyText = GetCellText(Range(strCopyRange).Cells(1, 1))
            Else
                strCopyText = MakeCopyText(strCopyRange)
            End If
        Else
            strCopyText = MakeCopyText()
        End If
    Case Else
        strCopyText = GetClipbordText()
        blnAllCell = CheckPasteMode(strCopyText, objSelection)
    End Select
    
    If strCopyText = "" Then
        Exit Sub
    Else
        '���s��CRLF��LF
        strCopyText = Replace$(strCopyText, vbCr, "")
    End If
    
    '�I��͈͂�\��t����ɕύX
    If blnOnlyCell = False And blnAllCell = False Then
        Set objSelection = GetPasteRange(strCopyText, objSelection)
    End If
    Call objSelection.Parent.Activate
    Call objSelection.Select
    
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    Call SaveUndoInfo(E_PasteValue, objSelection)
    If blnOnlyCell Or blnAllCell Then
        objSelection.Value = strCopyText
    Else
        '�^�u���܂܂�邩�ǂ����H
        If InStr(1, strCopyText, vbTab, vbBinaryCompare) = 0 Then
            Call PasteRows(strCopyText, objSelection)
        Else
            Call PasteTabText(strCopyText, objSelection)
        End If
    End If
    Call SetOnUndo
    Call SetOnRepeat
    
    '�\��t������������N���b�v�{�[�h�ɃR�s�[
    Call SetClipbordText(Replace$(strCopyText, vbLf, vbCrLf))
ErrHandle:
End Sub
   
'*****************************************************************************
'[ �֐��� ]�@MakeCopyText
'[ �T  �v ]�@�R�s�[�Ώۂ̕�������쐬����
'[ ��  �� ]�@Copy���̗̈�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Function MakeCopyText(Optional ByVal strCopyRange As String = "") As String
On Error GoTo ErrHandle
    If strCopyRange <> "" Then
        MakeCopyText = GetRangeText(Range(strCopyRange))
        Exit Function
    End If
    
    '���[�N�V�[�g�𗘗p���ăR�s�[����
    Dim objSheet As Worksheet
    Set objSheet = ThisWorkbook.Worksheets("Workarea1")
    
    With objSheet
        Call .Range("A1").PasteSpecial(xlPasteValues)
        Call ThisWorkbook.Activate
        Call .Activate
    End With

    MakeCopyText = GetRangeText(Selection)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
Exit Function
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
End Function

'*****************************************************************************
'[ �֐��� ]�@PasteRows
'[ �T  �v ]�@�N���b�v�{�[�h�̃e�L�X�g��1�s���Ƃɕ������ē\��t����
'[ ��  �� ]�@�R�s�[������A�R�s�[��̃Z��
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub PasteRows(ByVal strCopyText As String, ByRef objSelection As Range)
    Dim i           As Long
    Dim strArray()  As String
    
    Call GetStrArray(strCopyText, strArray())
        
    '1�s���ɕ������ē\��t��
    For i = 0 To UBound(strArray)
        objSelection.Rows(i + 1).Value = strArray(i)
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ]�@PasteTabText
'[ �T  �v ]�@�N���b�v�{�[�h�̃e�L�X�g���^�u���Ƃɕ�����ɓn���ē\��t����
'[ ��  �� ]�@�N���b�v�{�[�h�̕�����A�I��͈�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub PasteTabText(ByVal strText As String, ByRef objSelection As Range)
    Dim i           As Long
    Dim j           As Long
    Dim strArray()  As String
    Dim strCols     As Variant
    
    Call GetStrArray(strText, strArray())
    
    Application.ScreenUpdating = False
    
    '�s�̐��������[�v
    For i = 0 To UBound(strArray)
        strCols = Split(strArray(i), vbTab)
        '��̐��������[�v
        For j = 0 To UBound(strCols)
            objSelection(i + 1, j + 1).Value = strCols(j)
        Next j
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetPasteRange
'[ �T  �v ]�@�N���b�v�{�[�h�̃e�L�X�g��\��t����̈���Ď擾
'[ ��  �� ]�@�\��t����Text�A�I��͈�
'[ �߂�l ]�@�\��t���̈�
'*****************************************************************************
Public Function GetPasteRange(ByVal strCopyText As String, ByRef objSelection As Range) As Range
    Dim i           As Long
    Dim lngRowCount As Long
    Dim lngMaxCol   As Long
    Dim strArray()  As String
    Dim strCols     As Variant
    
    lngRowCount = GetStrArray(strCopyText, strArray())
    
    If InStr(1, strCopyText, vbTab, vbBinaryCompare) = 0 Then
        lngMaxCol = objSelection.Columns.Count
    Else
        '�ő��̎擾
        For i = 0 To UBound(strArray)
            strCols = Split(strArray(i), vbTab)
            lngMaxCol = WorksheetFunction.Max(lngMaxCol, UBound(strCols) + 1)
        Next i
    End If
    
    '�Ώۗ̈��I�����Ȃ���
    Set GetPasteRange = objSelection.Resize(lngRowCount, lngMaxCol)
End Function

'*****************************************************************************
'[ �֐��� ]�@CheckPasteMode
'[ �T  �v ]�@�N���b�v�{�[�h�̃e�L�X�g��\��t���郂�[�h�𔻒�
'[ ��  �� ]�@�\��t����Text�A�I��͈�
'[ �߂�l ]�@True:���ׂẴZ���ɓ\��t���AFalse:�Z�����ɕ������ē\��t��
'*****************************************************************************
Private Function CheckPasteMode(ByVal strCopyText As String, ByRef objSelection As Range) As Boolean
    '�I��͈͂������̗̈�̎�
    If objSelection.Areas.Count > 1 Then
        CheckPasteMode = True
        Exit Function
    End If
    
    '�e�L�X�g��1�s ���� 1��̎�
    If InStr(1, strCopyText, vbLf, vbBinaryCompare) = 0 And _
       InStr(1, strCopyText, vbTab, vbBinaryCompare) = 0 Then
        CheckPasteMode = True
        Exit Function
    End If
        
    '�s�����Ɍ����̂Ȃ��P��Z��
    If objSelection.Rows.Count = 1 Then
        '�s�̍�������������2�s�ȏ�̎�
        If objSelection.RowHeight > (objSelection.Font.Size + 2) * 2 Then
            CheckPasteMode = True
            Exit Function
        End If
    End If
    
    '�擪�s�ɍs�����̌��������鎞
    If objSelection.Rows.Count > 1 Then
        If Not (IntersectRange(ArrangeRange(objSelection.Rows(1)), ArrangeRange(objSelection.Rows(2))) Is Nothing) Then
            CheckPasteMode = True
            Exit Function
        End If
    End If

    CheckPasteMode = False
End Function

'*****************************************************************************
'[ �֐��� ]�@CopyText
'[ �T  �v ]�@�e�L�X�g���N���b�v�{�[�h�ɃR�s�[����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub CopyText()
On Error GoTo ErrHandle
    Dim strText      As String
    Dim objSelection As Range
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Set objSelection = Selection
    
    '�I��̈悪�����̎�
    If objSelection.Areas.Count > 1 Then
        Call objSelection.Copy
        strText = MakeCopyText()
    Else
        strText = GetRangeText(objSelection)
    End If
    
    If strText <> "" Then
        Call SetClipbordText(Replace$(strText, vbLf, vbCrLf))
    End If
    Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@MoveObject
'[ �T  �v ]�@Range�I�����F�����Z�����܂ޗ̈���ړ�����
'�@�@�@�@�@�@Shape�I�����F�}�`���ړ��܂��̓T�C�Y�ύX����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveObject()
    If GetKeyState(vbKeyControl) < 0 Then
        Call UnSelect
        Exit Sub
    End If

    'IME���I�t�ɂ���
    Call SetIMEOff

    '�I������Ă���I�u�W�F�N�g�𔻒�
    Select Case CheckSelection()
    Case E_Range
        Call MoveCell
    Case E_Shape
        Call MoveShape
    End Select
    
    Call Application.OnRepeat("", "")
End Sub
    
'*****************************************************************************
'[ �֐��� ]�@MoveCell
'[ �T  �v ]�@�����Z�����܂ޗ̈���ړ�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveCell()
On Error GoTo ErrHandle
    Dim enmModeType  As EModeType
    Dim objToRange   As Range
    Dim objFromRange As Range
    Dim lngCutCopyMode As Long
    
    '���[�h��ݒ肷��
    lngCutCopyMode = Application.CutCopyMode
    Select Case lngCutCopyMode
    Case xlCopy
        enmModeType = E_Copy
    Case xlCut
        enmModeType = E_CutInsert
    Case Else
        enmModeType = E_Move
    End Select
    
    '�R�s�[���Range��ݒ肷��
    Set objToRange = Selection
    
    '�R�s�[����Range��ݒ肷��
    Select Case lngCutCopyMode
    Case xlCopy, xlCut
        Dim strFromRange As String
        strFromRange = GetCopyRangeAddress()
        If strFromRange <> "" Then
            Set objFromRange = Range(strFromRange)
        Else
            Exit Sub
        End If
    Case Else
        Set objFromRange = Selection
    End Select
    
    '�`�F�b�N���s��
    Dim strErrMsg    As String
    If lngCutCopyMode = xlCut Then
        '�R�s�[���ƃR�s�[�悪�����V�[�g���ǂ���
        If CheckSameSheet(objFromRange.Parent, objToRange.Parent) = False Then
            Call MsgBox("�؂��莞�́A�����V�[�g�łȂ��Əo���܂���B", vbExclamation)
            Exit Sub
        End If
    End If
    strErrMsg = CheckMoveCell(objFromRange)
    If strErrMsg <> "" Then
        Call MsgBox(strErrMsg, vbExclamation)
        Exit Sub
    End If
    strErrMsg = CheckMoveCell(objToRange)
    If strErrMsg <> "" Then
        Call MsgBox(strErrMsg, vbExclamation)
        Exit Sub
    End If

    Call ShowMoveCellForm(enmModeType, objFromRange, objToRange)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@ShowMoveCellForm
'[ �T  �v ]�@�����Z�����܂ޗ̈���ړ�����
'[ ��  �� ]�@enmType:��ƃ^�C�v
'            objFromRange:�ړ�(�R�s�[��)�̗̈�
'            objToRange:�I�𒆂̗̈�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ShowMoveCellForm(ByVal enmType As EModeType, ByRef objFromRange As Range, ByRef objToRange As Range)
    Dim blnCopyObjectsWithCells  As Boolean
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells

On Error GoTo ErrHandle
    '�t�H�[����\��
    With frmMoveCell
        Call .Initialize(enmType, objFromRange, objToRange)
        Call .Show
    End With
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
Exit Sub
ErrHandle:
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
    If blnFormLoad = True Then
        Call Unload(frmMoveCell)
    End If
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@CheckSameSheet
'[ �T  �v ]  objSheet1��objSheet2�������V�[�g���ǂ�������
'[ ��  �� ]  ���肷��WorkSheet
'[ �߂�l ]  True:�����V�[�g
'*****************************************************************************
Public Function CheckSameSheet(ByRef objSheet1 As Worksheet, ByRef objSheet2 As Worksheet) As Boolean
    If objSheet1.Name = objSheet2.Name And _
       objSheet1.Parent.Name = objSheet2.Parent.Name Then
        CheckSameSheet = True
    Else
        CheckSameSheet = False
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@UnSelect
'[ �T  �v ]�@�I�����ꂽ�Z���̈悩��A�ꕔ�̗̈���I���ɂ���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub UnSelect()
    Static strLastSheet   As String '�O��̗̈�̕����p
    Static strLastAddress As String '�O��̗̈�̕����p
On Error GoTo ErrHandle
    Dim objSelection As Range
    Dim objUnSelect  As Range
    Dim objRange As Range
    Dim enmUnselectMode As EUnselectMode
    
    '�}�`���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    '����̈��I��������
    With frmUnSelect
        '�O��̕����p
        Call .SetLastSelect(strLastSheet, strLastAddress)
        
        '�t�H�[����\��
        Call .Show
        '�L�����Z����
        If blnFormLoad = False Then
            If objSelection.Areas.Count > 1 And strLastAddress = "" Then
                strLastSheet = ActiveSheet.Name
                strLastAddress = Selection.Address(False, False)
            End If
            Exit Sub
        End If
        
        enmUnselectMode = .Mode
        Set objSelection = Selection
        Select Case (enmUnselectMode)
        Case E_Unselect, E_Reverse, E_Intersect, E_Union
            Set objUnSelect = .SelectRange
        End Select
        Call Unload(frmUnSelect)
    End With

    Select Case (enmUnselectMode)
    Case E_Unselect  '�����
        Set objRange = MinusRange(objSelection, objUnSelect)
        Set objRange = ReSelectRange(objSelection, objRange)
    Case E_Reverse   '���]
        Set objRange = UnionRange(MinusRange(objSelection, objUnSelect), MinusRange(objUnSelect, objSelection))
    Case E_Intersect '�i�荞��
        Set objRange = IntersectRange(objSelection, objUnSelect)
        Set objRange = ReSelectRange(objSelection, objRange)
    Case E_Merge    '�����Z���̂ݑI��
        Set objRange = ArrangeRange(GetMergeRange(objSelection))
        Set objRange = ReSelectRange(objSelection, objRange)
    Case E_Union     '�ǉ�
        Dim strAddress As String
        strAddress = objSelection.Address(False, False) & "," & objUnSelect.Address(False, False)
        If Len(strAddress) < 256 Then
            Set objRange = Range(strAddress)
        Else
            Set objRange = UnionRange(objSelection, objUnSelect)
        End If
    End Select
    
    Call objRange.Select
    If Not (MinusRange(Selection, objRange) Is Nothing) Then
        '�d���̍폜 (�����Z��������ƑI��̈悪���������Ȃ邱�Ƃ�����)
        Call ArrangeRange(objRange).Select
    End If
    
ErrHandle:
    strLastSheet = ActiveSheet.Name
    strLastAddress = Selection.Address(False, False)
    If blnFormLoad = True Then
        Call Unload(frmUnSelect)
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@CheckMoveCell
'[ �T  �v ]�@MoveCell���\���ǂ����̃`�F�b�N
'[ ��  �� ]�@�R�s�[����Range
'[ �߂�l ]�@�G���[�̎��A�G���[���b�Z�[�W
'*****************************************************************************
Private Function CheckMoveCell(objSelection As Range) As String
    Dim objWorksheet As Worksheet
    Dim strSelection As String
    
    Application.ScreenUpdating = False
    Set objWorksheet = ActiveSheet
On Error GoTo ErrHandle
    
    '�I���G���A�������Ȃ�ΏۊO
    If objSelection.Areas.Count <> 1 Then
        CheckMoveCell = "���̃R�}���h�͕����̑I��͈͂ɑ΂��Ď��s�ł��܂���B"
        Exit Function
    End If
    
    '���ׂĂ̍s���I���Ȃ�ΏۊO(���삪���ɒx���Ȃ邽��)
    If objSelection.Rows.Count = Rows.Count Then
        CheckMoveCell = "���̃R�}���h�͂��ׂĂ̍s�̑I�����͎��s�ł��܂���B"
        Exit Function
    End If
    
    '���ׂĂ̗񂪑I������āA�s�����̌����Z�����܂ގ�
    If IsBorderMerged(objSelection) Then
        CheckMoveCell = "�������ꂽ�Z���̈ꕔ��ύX���邱�Ƃ͂ł��܂���B"
        Exit Function
    End If
ErrHandle:
    Call objWorksheet.Activate
    Application.ScreenUpdating = True
End Function

'*****************************************************************************
'[ �֐��� ]�@MoveShape
'[ �T  �v ]�@�}�`���ړ��܂��̓T�C�Y�ύX����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub MoveShape()
On Error GoTo ErrHandle
    '�t�H�[����\��
    Call frmMoveShape.Show
Exit Sub
ErrHandle:
    If blnFormLoad = True Then
        Call Unload(frmMoveShape)
    End If
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@FitShapes
'[ �T  �v ]  �I�����ꂽ�}�`��g���ɂ��킹��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub FitShapes()
On Error GoTo ErrHandle
    Dim blnOK      As Boolean
    Dim blnSizeChg As Boolean '�T�C�Y��ύX���邩�ǂ���
    Dim i          As Long
    
    '�}�`���I������Ă��邩����
    Select Case (CheckSelection())
    Case E_Range
        Call MsgBox("�}�`���I������Ă��܂���", vbExclamation)
        Exit Sub
    Case E_Other
        Exit Sub
    End Select
    
    '****************************************
    '�^�C�v��I��������
    '****************************************
    Dim enmSizeType As EFitType  '�I�����ꂽ�^�C�v
    With frmFitShapes
        '�t�H�[����\��
        Call .Show

        '�L�����Z����
        If blnFormLoad = False Then
            Exit Sub
        End If

        enmSizeType = .SelectType
        Call Unload(frmFitShapes)
    End With
    
    Select Case enmSizeType
    Case E_Default
        blnSizeChg = True
    Case E_TopLeft
        blnSizeChg = False
    Case E_Another
        Call MoveShape
        Exit Sub
    End Select
    
    Application.ScreenUpdating = False
    '�A���h�D�p�Ɍ��̃T�C�Y��ۑ�����
    Call SaveUndoInfo(E_ShapeSize, Selection.ShapeRange)
    
    '��]���Ă���}�`���O���[�v������
    Dim objGroups As ShapeRange
    Set objGroups = GroupSelection(Selection.ShapeRange)
    
    Call FitShapesGrid(objGroups, blnSizeChg)
    
    '��]���Ă���}�`�̃O���[�v�������������̐}�`��I������
    Call UnGroupSelection(objGroups).Select
    Call SetOnUndo
    Call SetOnRepeat
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@FitShapesGrid
'[ �T  �v ]  �I�����ꂽ�}�`��g���ɂ��킹��
'[ ��  �� ]�@objShapeRange�F�Ώې}�`�AblnSizeChg�F�T�C�Y��ύX���邩�ǂ���
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub FitShapesGrid(ByRef objShapeRange As ShapeRange, Optional blnSizeChg As Boolean = True)
    Dim objShape   As Shape     '�}�`
    Dim objRange   As Range

    '�}�`�̐��������[�v
    For Each objShape In objShapeRange
        Set objRange = GetNearlyRange(objShape)
        With objShape
            .Top = objRange.Top
            .Left = objRange.Left
            If blnSizeChg = True Then
                If .Height > 0.5 Then
                    .Height = objRange.Height
                Else
                    .Height = 0
                End If
                If .Width > 0.5 Then
                    .Width = objRange.Width
                Else
                    .Width = 0
                End If
            End If
        End With
    Next objShape
End Sub

'*****************************************************************************
'[ �֐��� ]�@GroupSelection
'[ �T  �v ]  �ύX�Ώۂ̐}�`�̒��ŉ�]���Ă�����̂��O���[�v������
'[ ��  �� ]�@�O���[�v���O�̐}�`
'[ �߂�l ]�@�O���[�v����̐}�`
'*****************************************************************************
Public Function GroupSelection(ByRef objShapes As ShapeRange) As ShapeRange
    Dim i            As Long
    Dim objShape     As Shape
    Dim btePlacement As Byte
    ReDim blnRotation(1 To objShapes.Count) As Boolean
    ReDim lngIDArray(1 To objShapes.Count) As Variant
    
    '�}�`�̐��������[�v
    For i = 1 To objShapes.Count
        Set objShape = objShapes(i)
        lngIDArray(i) = objShape.ID
        
        Select Case objShape.Rotation
        Case 90, 270, 180
            blnRotation(i) = True
        End Select
    Next

    '�}�`�̐��������[�v
    For i = 1 To objShapes.Count
        If blnRotation(i) = True Then
            Set objShape = GetShapeFromID(lngIDArray(i))
            btePlacement = objShape.Placement
            '�T�C�Y�ƈʒu������̃N���[�����쐬���O���[�v������
            With objShape.Duplicate
                .Top = objShape.Top
                .Left = objShape.Left
                '�����ɂ���
                .Fill.Visible = msoFalse
                .Line.Visible = msoFalse
                With GetShapeRangeFromID(Array(.ID, objShape.ID)).Group
                    .AlternativeText = "EL_TemporaryGroup" & i
                    .Placement = btePlacement
                    lngIDArray(i) = .ID
                End With
            End With
        End If
    Next
    
    Set GroupSelection = GetShapeRangeFromID(lngIDArray)
End Function

'*****************************************************************************
'[ �֐��� ]�@UnGroupSelection
'[ �T  �v ]  �ύX�Ώۂ̐}�`�̒��ŃO���[�v���������̂����ɖ߂�
'[ ��  �� ]�@�O���[�v�����O�̐}�`
'[ �߂�l ]�@�O���[�v������̐}�`
'*****************************************************************************
Public Function UnGroupSelection(ByRef objGroups As ShapeRange) As ShapeRange
    Dim i            As Long
    Dim btePlacement As Byte
    Dim objShape     As Shape
    ReDim blnRotation(1 To objGroups.Count) As Boolean
    ReDim lngIDArray(1 To objGroups.Count) As Variant
    
    '�}�`�̐��������[�v
    For i = 1 To objGroups.Count
        Set objShape = objGroups(i)
        lngIDArray(i) = objShape.ID
        
        If Left$(objShape.AlternativeText, 17) = "EL_TemporaryGroup" Then
            blnRotation(i) = True
        End If
    Next

    '�}�`�̐��������[�v
    For i = 1 To objGroups.Count
        If blnRotation(i) = True Then
            Set objShape = GetShapeFromID(lngIDArray(i))
            btePlacement = objShape.Placement
            With objShape.Ungroup
                .Item(1).Placement = btePlacement
                Call .Item(2).Delete
                lngIDArray(i) = .Item(1).ID
            End With
        End If
    Next i
    
    Set UnGroupSelection = GetShapeRangeFromID(lngIDArray)
End Function

'*****************************************************************************
'[ �֐��� ]�@ChangeTextboxesToCells
'[ �T  �v ]�@�e�L�X�g�{�b�N�X���Z���ɕϊ�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ChangeTextboxesToCells()
On Error GoTo ErrHandle
    Dim i            As Long
    Dim objTextbox   As Shape
    Dim objSelection As ShapeRange
    
    '�}�`���I������Ă��邩����
    Select Case (CheckSelection())
    Case E_Range
        Call MsgBox("�e�L�X�g�{�b�N�X���I������Ă��܂���", vbExclamation)
        Exit Sub
    Case E_Other
        Exit Sub
    End Select
    
    Set objSelection = Selection.ShapeRange
    ReDim lngIDArray(1 To objSelection.Count) As Variant
    
    '**************************************
    '�I�����ꂽ�}�`�̃`�F�b�N
    '**************************************
    '�}�`�̐��������[�v
    For i = 1 To objSelection.Count 'For each�\������Excel2007�Ō^�Ⴂ�ƂȂ�(���Ԃ�o�O)
        Set objTextbox = objSelection(i)
        lngIDArray(i) = objTextbox.ID
        '�e�L�X�g�{�b�N�X�������I������Ă��邩����
        If CheckTextbox(objTextbox) = False Then
            Call MsgBox("�e�L�X�g�{�b�N�X�ȊO�͑I�����Ȃ��ŉ�����", vbExclamation)
            Exit Sub
        End If
    Next

    '**************************************
    '�ϊ������Z����objRange�ɐݒ�
    '**************************************
    Dim objRange     As Range
    Dim blnNotMatch  As Boolean
    
    '�e�L�X�g�{�b�N�X�̐��������[�v
    For i = 1 To objSelection.Count 'For each�\������Excel2007�Ō^�Ⴂ�ƂȂ�(���Ԃ�o�O)
        Set objTextbox = objSelection(i)
        With GetNearlyRange(objTextbox)
            If IsBorderMerged(.Cells) Then
                Call MsgBox("�������ꂽ�Z���̈ꕔ��ύX���邱�Ƃ͂ł��܂���", vbExclamation)
                Call objTextbox.Select
                Exit Sub
            End If
        
            If IntersectRange(objRange, .Cells) Is Nothing Then
                Set objRange = UnionRange(objRange, .Cells)
            Else
                Call objTextbox.Select
                Call MsgBox("�ϊ������Z���ɏd��������܂�", vbExclamation)
                Exit Sub
            End If
            
            '�e�L�X�g�{�b�N�X���g���ƈ�v���Ă��邩����
            If .Top = objTextbox.Top And .Left = objTextbox.Left And _
               .Width = objTextbox.Width And .Height = objTextbox.Height Then
            Else
                blnNotMatch = True
            End If
        End With
    Next
    
    If blnNotMatch = True Then
        Call objRange.Select
        If MsgBox("�e�L�X�g�{�b�N�X���O���b�h(�g��)�ɂ����Ă��܂���B" & vbLf & _
                  "�ʒu�E�T�C�Y�̍ł��߂��Z���ɕϊ����܂��B" & vbLf & _
                  "��낵���ł����H", vbOKCancel + vbQuestion) = vbCancel Then
            Exit Sub
        End If
    End If
    
    '**************************************
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    '**************************************
    Dim strSelectAddress  As String
    strSelectAddress = objRange.Address
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False '�R�����g������ƌx�����o�鎞������
    
    Call SaveUndoInfo(E_TextToCell, objSelection, objRange)
    
    '**************************************
    '�e�L�X�g�{�b�N�X���Z���ɕϊ�����
    '**************************************
    Dim objShapeRange As ShapeRange
    Set objShapeRange = GetShapeRangeFromID(lngIDArray)
    '�e�L�X�g�{�b�N�X�̐��������[�v
    For i = 1 To objShapeRange.Count 'For each�\������Excel2007�Ō^�Ⴂ�ƂȂ�(���Ԃ�o�O)
        Set objTextbox = objShapeRange(i)
        Set objRange = GetNearlyRange(objTextbox)
        Call ChangeTextboxToCell(objTextbox, objRange)
    Next
    
    '�e�L�X�g�{�b�N�X���폜
    Call objShapeRange.Delete
    
    '�ϊ����ꂽ�Z����I������
    Call Range(strSelectAddress).Select
    Call SetOnUndo
    Application.DisplayAlerts = True
    Call SetOnRepeat
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@CheckTextbox
'[ �T  �v ]�@Shape���e�L�X�g�{�b�N�X���ǂ������肷��
'[ ��  �� ]�@���肷��Shape
'[ �߂�l ]�@True:�e�L�X�g�{�b�N�X
'*****************************************************************************
Private Function CheckTextbox(ByRef objShape As Shape) As Boolean
On Error GoTo ErrHandle
    '��]���Ă��邩
    If objShape.Rotation <> 0 Then
        Exit Function
    End If
    
    If (objShape.Type <> msoTextBox) And (objShape.Type <> msoAutoShape) Then
        Exit Function
    End If
    
'    If (TypeOf objShape.DrawingObject Is TextBox) Or _
'       (TypeOf objShape.DrawingObject Is Rectangle) Then '�l�p�`
'    Else
'       Exit Function
'    End If
    
    If IsNull(objShape.TextFrame.Characters.Count) Then
        Exit Function
    End If
    
    CheckTextbox = True
Exit Function
ErrHandle:
    CheckTextbox = False
End Function

'*****************************************************************************
'[ �֐��� ]�@ChangeTextboxToCell
'[ �T  �v ]�@�e�L�X�g�{�b�N�X���Z���ɕϊ�����
'[ ��  �� ]�@�e�L�X�g�{�b�N�X
'[ �߂�l ]�@�Z��
'*****************************************************************************
Private Sub ChangeTextboxToCell(ByRef objTextbox As Shape, ByRef objRange As Range)
    '�Z������������
    Call objRange.UnMerge
    Call objRange.Merge
    
    '�t�H���g�ƕ�����̐ݒ�
    objRange(1, 1).Value = GetCharactersText(objTextbox.TextFrame)
    On Error Resume Next
    With objTextbox.TextFrame.Font
        objRange.Font.Name = .Name
        objRange.Font.Size = .Size
        objRange.Font.FontStyle = .FontStyle
    End With
    On Error GoTo 0
    
    '�z�u�̐ݒ�
    With objTextbox.TextFrame
        On Error Resume Next
        If .Orientation = msoTextOrientationVertical Then
            objRange.Orientation = xlVertical       '�c����
        End If
        On Error GoTo 0
        
        '���ʒu�ݒ�
        On Error Resume Next
        Select Case .HorizontalAlignment
        Case xlHAlignLeft
            objRange.HorizontalAlignment = xlLeft
        Case xlHAlignCenter
            objRange.HorizontalAlignment = xlCenter
        Case xlHAlignRight
            objRange.HorizontalAlignment = xlRight
        Case xlHAlignDistributed
            objRange.HorizontalAlignment = xlDistributed
        Case xlHAlignJustify
            objRange.HorizontalAlignment = xlJustify
        End Select
        On Error GoTo 0
        
        '�c�ʒu�ݒ�
        On Error Resume Next
        Select Case .VerticalAlignment
        Case xlVAlignTop
            objRange.VerticalAlignment = xlTop
        Case xlVAlignCenter
            objRange.VerticalAlignment = xlCenter
        Case xlVAlignBottom
            objRange.VerticalAlignment = xlBottom
        Case xlVAlignDistributed
            objRange.VerticalAlignment = xlDistributed
        Case xlVAlignJustify
            objRange.VerticalAlignment = xlJustify
        End Select
        On Error GoTo 0
    End With

    '�r���̐ݒ�
    If objTextbox.Line.Visible <> msoFalse Then
        On Error Resume Next
        With objRange '�ΐ�
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
        End With
        With objRange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objRange.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        On Error GoTo 0
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@ChangeCellsToTextboxes
'[ �T  �v ]�@�Z�����e�L�X�g�{�b�N�X�ɕϊ�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ChangeCellsToTextboxes()
On Error GoTo ErrHandle
    Dim objRange        As Range
    Dim objSelection    As Range
    Dim strTextboxes()  As Variant
    Dim blnClear        As Boolean
    Dim i As Long
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    
    Dim strMsg As String
    strMsg = "���̃Z���̓��e���N���A���܂����H" & vbLf
    strMsg = strMsg & "�@�u �͂� �v���� �Z���̒l�������������͉�������" & vbCrLf
    strMsg = strMsg & "�@�u�������v���� �Z�������̂܂܂ɂ��Ă���"
    Select Case MsgBox(strMsg, vbYesNoCancel + vbQuestion + vbDefaultButton2)
    Case vbYes
        blnClear = True
    Case vbNo
        blnClear = False
    Case vbCancel
        Exit Sub
    End Select
    
    '�}�`����\�����ǂ�������
    If ActiveWorkbook.DisplayDrawingObjects = xlHide Then
        ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
    End If
    
    Set objSelection = Selection
    
    Application.ScreenUpdating = False
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    Call SaveUndoInfo(E_CellToText, objSelection)
    
    '�Z���P�ʂŃ��[�v
    For Each objRange In objSelection
        '�����Z���̎��A����̃Z���ȊO�͑ΏۊO
        If objRange(1, 1).Address = objRange.MergeArea(1, 1).Address Then
            i = i + 1
            ReDim Preserve strTextboxes(1 To i)
            strTextboxes(i) = ChangeCellToTextbox(objRange.MergeArea).Name
        End If
    Next objRange
        
    If blnClear Then
        With objSelection
            '���̗̈���N���A
            Call .Clear
            If Cells(Rows.Count - 2, Columns.Count - 2).MergeCells = False Then
                '�V�[�g��̕W���I�ȏ����ɐݒ�
                Call Cells(Rows.Count - 2, Columns.Count - 2).Copy(objSelection)
                Call .ClearContents
            End If
        End With
    End If
    
    '�쐬�����e�L�X�g�{�b�N�X��I��
    Call ActiveSheet.Shapes.Range(strTextboxes).Select
    Call SetOnUndo
    Call SetOnRepeat
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@ChangeCellToTextbox
'[ �T  �v ]�@�Z�����e�L�X�g�{�b�N�X�ɕϊ�����
'[ ��  �� ]�@�Z��
'[ �߂�l ]�@�e�L�X�g�{�b�N�X
'*****************************************************************************
Private Function ChangeCellToTextbox(ByRef objRange As Range) As Shape
    Dim objTextbox As Shape
    Dim objCell    As Range
    
    With objRange
        '�e�L�X�g�{�b�N�X�쐬
        Set objTextbox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
    End With
    
    Set objCell = objRange(1, 1)
    
    '�t�H���g�ƕ�����̐ݒ�
    objTextbox.DrawingObject.Formula = objCell.Address
    objTextbox.DrawingObject.Formula = ""
    If Val(Application.Version) >= 12 Then
        With objTextbox.TextFrame.Characters
            If .Text <> "" Then
                .Font.Name = objCell.Font.Name
                .Font.Size = objCell.Font.Size
                .Font.FontStyle = objCell.Font.FontStyle
            End If
        End With
    End If
    
    '�z�u�̐ݒ�
    With objTextbox.TextFrame
        If objCell.Orientation = xlVertical Then
            .Orientation = msoTextOrientationVertical     '�c����
        End If
        
        '���ʒu�ݒ�
        Select Case objCell.HorizontalAlignment
        Case xlLeft
            .HorizontalAlignment = xlHAlignLeft
        Case xlCenter
            .HorizontalAlignment = xlHAlignCenter
        Case xlRight
            .HorizontalAlignment = xlHAlignRight
        Case xlDistributed
            .HorizontalAlignment = xlHAlignDistributed
        Case xlJustify
            .HorizontalAlignment = xlHAlignJustify
        End Select
        
        '�c�ʒu�ݒ�
        Select Case objCell.VerticalAlignment
        Case xlTop
            .VerticalAlignment = xlVAlignTop
        Case xlCenter
            .VerticalAlignment = xlVAlignCenter
        Case xlBottom
            .VerticalAlignment = xlVAlignBottom
        Case xlDistributed
            .VerticalAlignment = xlVAlignDistributed
        Case xlJustify
            .VerticalAlignment = xlVAlignJustify
        End Select
    End With
    
    '���̐ݒ�
    With objTextbox.Line
        .Weight = DPIRatio
        .DashStyle = msoLineSolid
        .Style = msoLineSingle
        .ForeColor.RGB = 0
        .BackColor.RGB = RGB(255, 255, 255)
    End With
    With objRange
        If .Borders(xlEdgeTop).LineStyle = xlNone Or _
           .Borders(xlEdgeLeft).LineStyle = xlNone Or _
           .Borders(xlEdgeBottom).LineStyle = xlNone Or _
           .Borders(xlEdgeRight).LineStyle = xlNone Then
            objTextbox.Line.Visible = msoFalse
        Else
            objTextbox.Line.Visible = msoTrue
        End If
    End With
    
    Set ChangeCellToTextbox = objTextbox
End Function

'*****************************************************************************
'[ �֐��� ]�@HideShapes
'[ �T  �v ]  �u�b�N���̂��ׂĂ̐}�`���\���ɂ���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub HideShapes()
    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If
    
    With ActiveWorkbook
        If .DisplayDrawingObjects = xlDisplayShapes Then
            If MsgBox("���[�N�u�b�N���̂��ׂĂ̐}�`���\���ɂ��܂�" & vbLf & _
                      "���s����΁A�}�`�̕ҏW���o���Ȃ��Ȃ�܂�" & vbLf & _
                      "��낵���ł����H" & vbLf & _
                      "���ēx�\������ɂ́A������x�N���b�N���ĉ�����" & vbLf & _
                      "���uCtrl+6�v���������Ă�Excel�̕W���@�\�œ��l�̑��삪�s���܂�" _
                      , vbOKCancel + vbQuestion) = vbOK Then
                .DisplayDrawingObjects = xlHide
            End If
        Else
            .DisplayDrawingObjects = xlDisplayShapes
        End If
    End With
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetBorder
'[ �T  �v ]�@Border�I�u�W�F�N�g��TBorder�\���̂ɑ������
'[ ��  �� ]�@Border�I�u�W�F�N�g
'[ �߂�l ]�@TBorder�\����
'*****************************************************************************
Public Function GetBorder(ByRef objBorder As Border) As TBorder
    With objBorder
        GetBorder.LineStyle = .LineStyle
        GetBorder.ColorIndex = .ColorIndex
        GetBorder.Weight = .Weight
        GetBorder.Color = .Color
    End With
End Function

'*****************************************************************************
'[ �֐��� ]�@SetBorder
'[ �T  �v ]�@TBorder�\���̂�Border�I�u�W�F�N�g�ɐݒ肷��
'[ ��  �� ]�@TBorder�\����
'            Border�I�u�W�F�N�g
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetBorder(ByRef udtBorder As TBorder, ByRef objBorder As Border)
    With objBorder
        If .LineStyle <> udtBorder.LineStyle Then
            .LineStyle = udtBorder.LineStyle
        End If
        If .ColorIndex <> udtBorder.ColorIndex Then
            .ColorIndex = udtBorder.ColorIndex
        End If
        If .Weight <> udtBorder.Weight Then
            .Weight = udtBorder.Weight
        End If
        If .Color <> udtBorder.Color Then
            .Color = udtBorder.Color
        End If
    End With
End Sub

'*****************************************************************************
'[ �֐��� ]�@SaveUndoInfo
'[ �T  �v ]�@Undo����ۑ�����
'[ ��  �� ]�@enmType:Undo�^�C�v�AobjObject:�����Ώۂ̑I�����ꂽ�I�u�W�F�N�g
'            varInfo:�t�����
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SaveUndoInfo(ByVal enmType As EUndoType, ByRef objObject As Object, Optional ByVal varInfo As Variant = Nothing)
    '���łɃC���X�^���X�����݂��鎞�́ARollback�ɑΉ����邽��New���Ȃ�
    If clsUndoObject Is Nothing Then
        Set clsUndoObject = New CUndoObject
    End If
    
    '�I�[�g�t�B���^���ݒ肳��Ă��鎞�́AUndo�s�ɂ���
    If (ActiveSheet.AutoFilter Is Nothing) And (ActiveSheet.FilterMode = False) Then
        Call clsUndoObject.SaveUndoInfo(enmType, objObject, varInfo)
    Else
        Call clsUndoObject.SaveUndoInfo(E_FilterERR, objObject, varInfo)
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetOnUndo
'[ �T  �v ]�@Application�I�u�W�F�N�g��OnUndo�C�x���g��ݒ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetOnUndo()
    '�L�[�̍Ē�` Excel�̃o�O?�ŃL�[�������ɂȂ邱�Ƃ����邽��
    Call SetKeys

    Call clsUndoObject.SetOnUndo
    Call Application.OnRepeat("", "")
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetOnRepeat
'[ �T  �v ]�@Application�I�u�W�F�N�g��OnRepeat�C�x���g��ݒ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetOnRepeat()
    Set objRepeatCmd = CommandBars.ActionControl
    If Not (objRepeatCmd Is Nothing) Then
        Call Application.OnRepeat("�J��Ԃ��@" & CommandBars.ActionControl.Caption, "OnRepeat")
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@OnRepeat
'[ �T  �v ]�@�J��Ԃ��N���b�N��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub OnRepeat()
    Call objRepeatCmd.Execute
End Sub

'*****************************************************************************
'[ �֐��� ]�@ExecUndo
'[ �T  �v ]�@Undo�����s����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub ExecUndo()
On Error GoTo ErrHandle
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Call clsUndoObject.ExecUndo
    Set clsUndoObject = Nothing
    Application.DisplayAlerts = True
    Call Application.OnRepeat("", "")
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Set clsUndoObject = Nothing
    Call MsgBox(Err.Description, vbExclamation)
    Call Application.OnRepeat("", "")
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetPlacement
'[ �T  �v ]�@Shape��Placement�v���p�e�B��ύX����
'�@�@�@�@�@�@�Z���ɂ��킹��Shape�̈ʒu�ƃT�C�Y��ύX�����Ȃ��悤�ɂ���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetPlacement()
On Error GoTo ErrHandle
    Dim i          As Long
    Dim lngDisplay As Long
    
    If ActiveSheet.Shapes.Count = 0 Then
        Exit Sub
    End If
    
    lngDisplay = ActiveWorkbook.DisplayDrawingObjects
    ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
    
    ReDim udtPlacement(1 To ActiveSheet.Shapes.Count)
    For i = 1 To ActiveSheet.Shapes.Count
        With ActiveSheet.Shapes(i)
            udtPlacement(i).Placement = .Placement
            .Placement = xlFreeFloating
            If .Type = msoComment Then
                udtPlacement(i).Top = .Top
                udtPlacement(i).Height = .Height
                udtPlacement(i).Left = .Left
                udtPlacement(i).Width = .Width
            End If
        End With
    Next i
ErrHandle:
    ActiveWorkbook.DisplayDrawingObjects = lngDisplay
End Sub

'*****************************************************************************
'[ �֐��� ]�@ResetPlacement
'[ �T  �v ]�@Shape��Placement�v���p�e�B�����ɖ߂�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub ResetPlacement()
On Error GoTo ErrHandle
    Dim i          As Long
    Dim lngDisplay As Long
    
    If ActiveSheet.Shapes.Count = 0 Then
        Exit Sub
    End If
    
    lngDisplay = ActiveWorkbook.DisplayDrawingObjects
    ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes

    For i = 1 To ActiveSheet.Shapes.Count
        With ActiveSheet.Shapes(i)
            .Placement = udtPlacement(i).Placement
            If .Type = msoComment Then
                .Top = udtPlacement(i).Top
                .Height = udtPlacement(i).Height
                .Left = udtPlacement(i).Left
                .Width = udtPlacement(i).Width
            End If
        End With
    Next i
ErrHandle:
    Erase udtPlacement()
    ActiveWorkbook.DisplayDrawingObjects = lngDisplay
End Sub

'*****************************************************************************
'[ �֐��� ]�@GetShapeRangeFromID
'[ �T  �v ]�@Shpes�I�u�W�F�N�g��ID����ShapeRange�I�u�W�F�N�g���擾
'[ ��  �� ]�@ID�̔z��
'[ �߂�l ]�@ShapeRange�I�u�W�F�N�g
'*****************************************************************************
Public Function GetShapeRangeFromID(ByRef lngID As Variant) As ShapeRange
    Dim i As Long
    Dim j As Long
    Dim lngShapeID As Long
    ReDim lngArray(LBound(lngID) To UBound(lngID)) As Variant
    
    For j = 1 To ActiveSheet.Shapes.Count
        lngShapeID = ActiveSheet.Shapes(j).ID
        For i = LBound(lngID) To UBound(lngID)
            If lngShapeID = lngID(i) Then
                lngArray(i) = j
                Exit For
            End If
        Next
    Next
    
    Set GetShapeRangeFromID = ActiveSheet.Shapes.Range(lngArray)
End Function

'*****************************************************************************
'[ �֐��� ]�@GetShapeFromID
'[ �T  �v ]�@Shape�I�u�W�F�N�g��ID����Shape�I�u�W�F�N�g���擾
'[ ��  �� ]�@ID
'[ �߂�l ]�@Shape�I�u�W�F�N�g
'*****************************************************************************
Public Function GetShapeFromID(ByVal lngID As Long) As Shape
    Dim j As Long
    Dim lngIndex As Long
        
    For j = 1 To ActiveSheet.Shapes.Count
        If ActiveSheet.Shapes(j).ID = lngID Then
            lngIndex = j
            Exit For
        End If
    Next j
    
    Set GetShapeFromID = ActiveSheet.Shapes.Range(j).Item(1)
End Function

'*****************************************************************************
'[ �֐��� ]�@OnPopupClick
'[ �T  �v ]�@MoveShape��ʂ̃|�b�v�A�b�v���j���[���N���b�N���������s�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub OnPopupClick()
    Call frmMoveShape.OnPopupClick
End Sub

'*****************************************************************************
'[ �֐��� ]�@OnPopupClick2
'[ �T  �v ]�@���͕⏕��ʂ̃|�b�v�A�b�v���j���[���N���b�N���������s�����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub OnPopupClick2()
    Call frmEdit.OnPopupClick
End Sub

'*****************************************************************************
'[ �֐��� ]�@ConvertStr
'[ �T  �v ]�@������̕ϊ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ConvertStr()
On Error GoTo ErrHandle
    Dim objSelection As Range
    Dim objWkRange   As Range
    Dim objCell      As Range
    Dim strText      As String
    Dim strConvText  As String
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    Set objSelection = Selection
    
    '�����̓��͂��ꂽ�Z���̂ݑΏۂɂ���
    On Error Resume Next
    Set objWkRange = IntersectRange(objSelection, Cells.SpecialCells(xlCellTypeConstants))
    On Error GoTo 0
    If objWkRange Is Nothing Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim i As Long
    Dim j As Long
'    Dim t
'    t = Timer
    
    Call SaveUndoInfo(E_CellValue, objSelection)
    On Error Resume Next
    For Each objCell In objWkRange
        i = i + 1
        strText = objCell
        If strText <> "" Then '�����Z���̍���ȊO�𖳎����邽��
            strConvText = StrConvert(strText, CommandBars.ActionControl.Parameter)
            If strConvText <> strText Then
'                objCell = strConvText
                Call SetTextToCell(objCell, strConvText)
                
                '�X�e�[�^�X�o�[�ɐi���󋵂�\��
                If i / objWkRange.Count * 12 <> j Then
                    j = i / objWkRange.Count * 12
                    Application.StatusBar = String(j, "��") & String(12 - j, "��")
                End If
            End If
        End If
    Next
    
'    MsgBox Timer - t
    
    Call SetOnUndo
    
    Set objRepeatCmd = CommandBars.ActionControl
    Call Application.OnRepeat("�J��Ԃ� " & objRepeatCmd.Caption, "OnRepeat")
    Application.StatusBar = False
Exit Sub
ErrHandle:
    Application.StatusBar = False
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetTextToCell
'[ �T  �v ]�@�ݒ肷��Cell,�ݒ肷�镶����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub SetTextToCell(ByRef objCell As Range, ByVal strConvText As String)
On Error GoTo ErrHandle
    objCell = objCell.PrefixCharacter & strConvText
    If objCell.HasFormula Then
        objCell = "'" & strConvText
    End If
Exit Sub
ErrHandle:
    objCell = "'" & strConvText
End Sub

'*****************************************************************************
'[ �֐��� ]�@OpenEdit
'[ �T  �v ]�@���͎x���G�f�B�^���J��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub OpenEdit()
On Error GoTo ErrHandle
    Static udtEditInfo As TEditInfo
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If

    '�f�t�H���g�l�̐ݒ�
    If udtEditInfo.Width = 0 Then
        frmEdit.StartUpPosition = 2 '��ʂ̒���
        udtEditInfo.Height = 405
        udtEditInfo.Width = 660
        udtEditInfo.FontSize = 10
        udtEditInfo.WordWarp = False
    Else
        frmEdit.StartUpPosition = 0 '�蓮
    End If
    
    With frmEdit
        '��ʈʒu�̐ݒ�
        If .StartUpPosition = 0 Then
            .Top = udtEditInfo.Top
            .Left = udtEditInfo.Left
        End If
        
        '�f�t�H���g�l�̐ݒ�
        .Height = udtEditInfo.Height
        .Width = udtEditInfo.Width
        .SpbSize = udtEditInfo.FontSize
        .Zoomed = udtEditInfo.Zoomed
        .chkWordWrap = udtEditInfo.WordWarp
        
        '�t�H�[����\��
        Call .Show
        
        '�t�H�[���̃T�C�Y����ۑ�
        If ActiveCell.HasFormula = True Then
            udtEditInfo.Height = WorksheetFunction.Max(.Height, udtEditInfo.Height)
        Else
            udtEditInfo.Height = .Height
        End If
        udtEditInfo.Top = .Top
        udtEditInfo.Left = .Left
        udtEditInfo.Width = .Width
        udtEditInfo.FontSize = .SpbSize
        udtEditInfo.Zoomed = .Zoomed
        udtEditInfo.WordWarp = .chkWordWrap
        
        Call Unload(frmEdit)
    End With
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[ �֐��� ]�@OnElseMenuClick
'[ �T  �v ]�@����̑�����j���[���J�����ɁA�e�R�}���h��Enabled�̏����ݒ���s��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub OnElseMenuClick()
    Dim objMenu          As CommandBarPopup
    Dim enmSelectionType As ESelectionType
    Dim objCommand       As Object
        
    Call SetKeys
    
    enmSelectionType = CheckSelection()
    Set objMenu = CommandBars.ActionControl
    
    'Menu���J�X�^�}�C�Y����Ă����ꍇ�̍l��
    For Each objCommand In objMenu.Controls
        objCommand.Enabled = True
    Next
    
    On Error Resume Next
    objMenu.Controls("�e�L�X�g�{�b�N�X���Z���ɕϊ�").Enabled = (enmSelectionType = E_Shape)
    objMenu.Controls("�Z�����e�L�X�g�{�b�N�X�ɕϊ�").Enabled = (enmSelectionType = E_Range)
    objMenu.Controls("�}�`���O���b�h�ɍ�����").Enabled = (enmSelectionType = E_Shape)
    objMenu.Controls("������̕ϊ�").Enabled = (enmSelectionType = E_Range)
    objMenu.Controls("�I��̈�̈ꕔ�����").Enabled = (enmSelectionType = E_Range)
    objMenu.Controls("���͕⏕���").Enabled = (enmSelectionType = E_Range)

    '�}�`��\�����}�`�ĕ\���ɂȂ��Ă����猳�ɖ߂�
    objMenu.Controls("�}�`�ĕ\��").Caption = "�}�`��\��"
    If ActiveWorkbook Is Nothing Then
        objMenu.Controls("�}�`��\��").Enabled = False
    Else
        If ActiveWorkbook.DisplayDrawingObjects <> xlDisplayShapes Then
            objMenu.Controls("�}�`��\��").Caption = "�}�`�ĕ\��"
        End If
    End If
    On Error GoTo 0
End Sub

'*****************************************************************************
'[ �֐��� ]�@PressBackSpace
'[ �T  �v ]�@�o�b�N�X�y�[�X�L�[���������Ƃ��̓����ύX����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub PressBackSpace()
On Error GoTo ErrHandle
    Call Application.OnKey("{BS}")
    If CheckSelection() = E_Range Then
        Call SendKeys("{F2}")
    Else
        Call SendKeys("{BS}")
    End If
    Call Application.OnTime(Now(), "SetBackSpace")
ErrHandle:
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetBackSpace
'[ �T  �v ]�@�o�b�N�X�y�[�X�L�[���������Ƃ��̓����ύX����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetBackSpace()
On Error GoTo ErrHandle
    Call Application.OnKey("{BS}", "PressBackSpace")
ErrHandle:
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetOption
'[ �T  �v ]�@�I�v�V�����̐ݒ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub SetOption()
    '�t�H�[����\��
    Call frmOption.Show
    
    '�V���[�g�J�b�g�L�[�̐ݒ�
    Call SetKeys
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetKeys
'[ �T  �v ]�@�V���[�g�J�b�g�L�[�̐ݒ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetKeys()
On Error GoTo ErrHandle
    Dim strOption As String
    Dim blnKeys(1 To 4) As Boolean
    
    strOption = CommandBars("���񂽂񃌃C�A�E�g").Controls(1).Tag
    blnKeys(1) = Not (InStr(1, strOption, "{S+F2}") = 0)
    blnKeys(2) = Not (InStr(1, strOption, "{C+S+C}") = 0)
    blnKeys(3) = Not (InStr(1, strOption, "{C+S+V}") = 0)
    blnKeys(4) = Not (InStr(1, strOption, "{BS}") = 0)
    
    If blnKeys(1) = True Then
        Call Application.OnKey("+{F2}", "OpenEdit")
    Else
        Call Application.OnKey("+{F2}")
    End If
    
    If blnKeys(2) = True Then
        Call Application.OnKey("+^{c}", "CopyText")
    Else
        Call Application.OnKey("+^{c}")
    End If
    
    If blnKeys(3) = True Then
        Call Application.OnKey("+^{v}", "PasteText")
    Else
        Call Application.OnKey("+^{v}")
    End If
    
    If blnKeys(4) = True Then
        Call Application.OnKey("{BS}", "PressBackSpace")
    Else
        Call Application.OnKey("{BS}")
    End If
ErrHandle:
End Sub

''*****************************************************************************
''[ �֐��� ]�@SubClassProc
''[ �T  �v ]�@���͕⏕��ʂ��}�E�X�z�C�[���ŃX�N���[��������
''[ ��  �� ]�@CallBack�֐��̂���
''[ �߂�l ]�@CallBack�֐��̂���
''*****************************************************************************
'Public Function SubClassProc(ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'On Error Resume Next
'    If MSG = WM_MOUSEWHEEL Then
'        If 0 < wParam Then
'            Call SendKeys("{UP}")
'            Call SendKeys("{UP}")
''            frmEdit.txtEdit.CurLine = frmEdit.txtEdit.CurLine - 2
'        Else
'            Call SendKeys("{DOWN}")
'            Call SendKeys("{DOWN}")
''            frmEdit.txtEdit.CurLine = frmEdit.txtEdit.CurLine + 2
'        End If
'    End If
'
'    '�f�t�H���g�E�B���h�E�v���V�[�W�����Ăяo��
'    SubClassProc = CallWindowProc(frmEdit.WndProc, hWnd, MSG, wParam, lParam)
'End Function

''*****************************************************************************
''[ �֐��� ]�@LoadThisWorkbook
''[ �T  �v ]�@Open���ɍ������̂��߁AThisWorkbook�����[�h������
''[ ��  �� ]�@�Ȃ�
''[ �߂�l ]�@�Ȃ�
''*****************************************************************************
'Public Sub LoadThisWorkbook()
'On Error Resume Next
'    Application.StatusBar = "�J���Ă��܂�  ���񂽂񃌃C�A�E�g"
'    With ThisWorkbook.Styles("Normal").Font
''        .Size = 11
'    End With
'    Application.StatusBar = False
'End Sub

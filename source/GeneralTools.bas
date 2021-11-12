Attribute VB_Name = "GeneralTools"
Option Explicit
Option Private Module

Public Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr
Public Declare PtrSafe Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As LongPtr, ByVal uIDEnableItem As Long, ByVal uEnable As Long) As Long
Public Declare PtrSafe Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
'Public Declare PtrSafe Function CloseHandle Lib "KERNEL32.DLL" (ByVal hObject As Longptr) As Long
'Public Declare PtrSafe Function TerminateProcess Lib "KERNEL32.DLL" (ByVal hProcess As Longptr, ByVal uExitCode As Long) As Long
Public Declare PtrSafe Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
Public Declare PtrSafe Function SetCursor Lib "user32.dll" (ByVal hCursor As LongPtr) As LongPtr
Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
'Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Longptr, ByVal hwnd As Longptr, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare PtrSafe Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ImmSetOpenStatus Lib "imm32.dll" (ByVal himc As LongPtr, ByVal b As Long) As Long
Public Declare PtrSafe Function ImmReleaseContext Lib "imm32.dll" (ByVal hwnd As LongPtr, ByVal himc As LongPtr) As Long
Public Declare PtrSafe Function ReleaseCapture Lib "user32.dll" () As Long
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Public Declare PtrSafe Function GlobalAlloc Lib "Kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As LongPtr
Public Declare PtrSafe Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Public Declare PtrSafe Function GlobalSize Lib "Kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GlobalLock Lib "Kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "Kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

' �萔�̒�`
Public Const IDC_HAND = 32649
Public Const IDC_SIZENWSE = 32642
Public Const SC_CLOSE = 61536
Public Const SC_SIZE = &HF000&
Public Const GWL_WNDPROC = (-4)
Public Const WM_SYSCOMMAND = &H112
Public Const WM_RBUTTONDOWN = &H204 '�E�}�E�X�{�^����������
Public Const WM_MOUSEWHEEL = &H20A  '�z�C�[�����񂳂ꂽ�iWin98,NT4.0�ȍ~�j
Public Const MF_BYCOMMAND = 0
Public Const MF_GRAYED = 1
Public Const GWL_STYLE = (-16)
Public Const WS_THICKFRAME = &H40000 '�E�B���h�E�̃T�C�Y�ύX
Public Const WS_MINIMIZEBOX = &H20000 '�ŏ����{�^��
Public Const WS_MAXIMIZEBOX = &H10000 '�ő剻�{�^��
Public Const SW_SHOWNORMAL = 1
Public Const SW_MAXIMIZE = 3
Public Const SYNCHRONIZE       As Long = &H100000
Public Const PROCESS_TERMINATE As Long = &H1
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103
Public Const MAXROWCOLCNT = 1000
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90

Public Const REGKEY = "EasyLout"
Public Const DEFAULTFONT = "�l�r �S�V�b�N"

'Public DPIRatio As Double

'�I���^�C�v
Public Enum ESelectionType
    E_Range
    E_Shape
    E_Non
    E_Other
End Enum

'�����^�C�v
Public Enum EMergeType
    E_MTROW
    E_MTCOL
    E_MTBOTH
End Enum

'�\�[�g�p�\����
Public Type TSortArray
    Key1  As Long
    Key2  As Long
    Key3  As Long
End Type

'*****************************************************************************
'[ �֐��� ]�@CheckSelection
'[ �T  �v ]�@�I������Ă��邩�I�u�W�F�N�g�̎�ނ𔻒肷��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@Range�AShape�A���̑��@�̂����ꂩ
'*****************************************************************************
Public Function CheckSelection() As ESelectionType
On Error GoTo ErrHandle
    If ActiveWorkbook Is Nothing Then
        CheckSelection = E_Non
        Exit Function
    End If
    
    If Selection Is Nothing Then
        CheckSelection = E_Other
        Exit Function
    End If
    
    If TypeOf Selection Is Range Then
        CheckSelection = E_Range
    ElseIf TypeOf Selection.ShapeRange Is ShapeRange Then
        CheckSelection = E_Shape
    Else
        CheckSelection = E_Other
    End If
Exit Function
ErrHandle:
    CheckSelection = E_Other
End Function

'*****************************************************************************
'[ �֐��� ]�@GetRangeText
'[ �T  �v ]�@�e�s�̕���������s�ŁA�e��̕�������󔒂ŋ�؂��ĘA������
'[ ��  �� ]�@�Ώۂ̗̈�
'[ �߂�l ]�@�A�����ꂽ������
'*****************************************************************************
Public Function GetRangeText(ByRef objRange As Range) As String
    Dim i   As Long
    Dim lngLast    As Long
    
    '���ׂĂ̍s�̑I����
    If objRange.Rows.Count = Rows.Count Then
        '�g�p����Ă���Ō�̍s
        lngLast = Cells.SpecialCells(xlCellTypeLastCell).Row
    Else
        lngLast = objRange.Rows.Count
    End If
    
    '�s�̐��������[�v
    For i = 1 To lngLast
        GetRangeText = GetRangeText & GetRowText(objRange.Rows(i)) & vbLf
    Next i
    
    '�擪�Ɩ����̋󔒍s���폜
    GetRangeText = TrimChr(GetRangeText)
End Function

'*****************************************************************************
'[ �֐��� ]�@GetRowText
'[ �T  �v ]�@�e��̕�������󔒂ŋ�؂��ĘA������
'[ ��  �� ]�@�Ώۂ̂P�s
'[ �߂�l ]�@�A�����ꂽ������
'*****************************************************************************
Private Function GetRowText(ByRef objRange As Range) As String
    Dim i       As Long
    Dim strText As String
    
    '��̐��������[�v
    For i = 1 To objRange.Columns.Count
        strText = GetCellText(objRange.Cells(1, i))
        GetRowText = GetRowText & strText & vbTab
    Next i

    '������Tab���폜
    GetRowText = RTrimChr(GetRowText, vbTab)
End Function

'*****************************************************************************
'[ �֐��� ]�@GetCellText
'[ �T  �v ]�@Cell�̕������\�����ꂽ�����Ŏ擾����
'[ ��  �� ]�@�Ώۂ̃Z��
'[ �߂�l ]�@������
'*****************************************************************************
Public Function GetCellText(ByRef objCell As Range) As String
On Error GoTo ErrHandle
    Select Case objCell.NumberFormat
    Case "General", "@"
        GetCellText = Rtrim$(objCell.Value)
        Exit Function
    End Select
                
    If objCell.Text <> WorksheetFunction.Rept("#", Len(objCell.Text)) Then
        GetCellText = Rtrim$(objCell.Text)
        Exit Function
    End If

    If IsDate(objCell.Value) Then
        GetCellText = WorksheetFunction.Text(objCell.Value, objCell.NumberFormatLocal)
        Exit Function
    End If
    
    If IsNumeric(objCell.Value) Then
        GetCellText = objCell.Value
        Exit Function
    End If
ErrHandle:
    GetCellText = Rtrim$(objCell.Value)
End Function

'*****************************************************************************
'[ �֐��� ]�@TrimChr
'[ �T  �v ]�@������̐擪�Ɩ����̉��s��^�u�������폜����
'[ ��  �� ]�@�폜���镶��
'[ �߂�l ]�@������
'*****************************************************************************
Public Function TrimChr(ByVal strText As String, Optional ByVal strChr As String = vbLf) As String
    TrimChr = LTrimChr(strText, strChr)
    TrimChr = RTrimChr(TrimChr, strChr)
End Function

'*****************************************************************************
'[ �֐��� ]�@LTrimChr
'[ �T  �v ]�@������̐擪�̉��s��^�u�������폜����
'[ ��  �� ]�@�폜���镶��
'[ �߂�l ]�@������
'*****************************************************************************
Public Function LTrimChr(ByVal strText As String, Optional ByVal strChr As String = " ") As String
    Dim i        As Long
    Dim lngStart As Long
    
    '�O����胋�[�v
    For i = 1 To Len(strText)
        If Mid$(strText, i, 1) <> strChr Then
            lngStart = i
            Exit For
        End If
    Next
    
    If lngStart > 0 Then
        LTrimChr = Mid$(strText, lngStart)
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@RTrimChr
'[ �T  �v ]�@������̖����̉��s��^�u�������폜����
'[ ��  �� ]�@�폜���镶��
'[ �߂�l ]�@������
'*****************************************************************************
Public Function RTrimChr(ByVal strText As String, Optional ByVal strChr As String = " ") As String
    Dim i        As Long
    Dim lngEnd   As Long
    
    '�����胋�[�v
    For i = Len(strText) To 1 Step -1
        If Mid$(strText, i, 1) <> strChr Then
            lngEnd = i
            Exit For
        End If
    Next
    
    If lngEnd > 0 Then
        RTrimChr = Left$(strText, lngEnd)
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@GetStrArray
'[ �T  �v ]�@����������s�ł΂炵�ĂP�s���Ƃ̔z��ŕԂ�
'[ ��  �� ]�@strText:���̕�����AStrArray:1�s���Ƃ̔z��
'[ �߂�l ]�@�s��
'*****************************************************************************
Public Function GetStrArray(ByVal strText As String, Optional ByRef strArray As Variant) As Long
    '���s�܂��͋󔒂���
    If Trim$(Replace$(strText, vbLf, "")) = "" Then
        GetStrArray = 0
        Exit Function
    End If
    
    '�P�s���Ƃɔz��Ɋi�[
    strArray = Split(TrimChr(strText), vbLf)
    
    GetStrArray = UBound(strArray) + 1
End Function

'*****************************************************************************
'[ �֐��� ]�@IntersectRange
'[ �T�@�v ]�@�̈�Ɨ̈�̏d�Ȃ�̈���擾����
'�@�@�@�@�@�@�`���a
'[ ���@�� ]�@�Ώۗ̈�(Nothing����)
'[ �߂�l ]�@objRange1 �� objRange2
'*****************************************************************************
Public Function IntersectRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Select Case True
    Case (objRange1 Is Nothing) Or (objRange2 Is Nothing)
        Set IntersectRange = Nothing
    Case Else
        Set IntersectRange = Intersect(objRange1, objRange2)
    End Select
End Function

'*****************************************************************************
'[ �֐��� ]�@UnionRange
'[ �T�@�v ]�@�̈�ɗ̈��������
'�@�@�@�@�@�@�`���a
'[ ���@�� ]�@�Ώۗ̈�(Nothing����)
'[ �߂�l ]�@objRange1 �� objRange2
'*****************************************************************************
Public Function UnionRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Select Case True
    Case (objRange1 Is Nothing) And (objRange2 Is Nothing)
        Set UnionRange = Nothing
    Case (objRange1 Is Nothing)
        Set UnionRange = objRange2
    Case (objRange2 Is Nothing)
        Set UnionRange = objRange1
    Case Else
        Set UnionRange = Union(objRange1, objRange2)
    End Select
End Function

'*****************************************************************************
'[ �֐��� ]�@MinusRange
'[ �T�@�v ]�@�̈悩��̈���A���O����
'�@�@�@�@�@�@�`�|�a = �`��!�a
'�@�@�@�@�@�@!�a = !(B1��B2��B3...��Bn) = !B1��!B2��!B3...��!Bn
'[ ���@�� ]�@�Ώۗ̈�
'[ �߂�l ]�@objRange1 �| objRange2
'*****************************************************************************
Public Function MinusRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Dim objRounds As Range
    Dim i As Long
    
    If objRange2 Is Nothing Then
        Set MinusRange = objRange1
        Exit Function
    End If
    
    '���O����̈�̐��������[�v
    '!�a = !B1��!B2��!B3.....��!Bn
    Set objRounds = ReverseRange(objRange2.Areas(1))
    For i = 2 To objRange2.Areas.Count
        Set objRounds = IntersectRange(objRounds, ReverseRange(objRange2.Areas(i)))
    Next
    
    '�`��!�a
    Set MinusRange = IntersectRange(objRange1, objRounds)
End Function

'*****************************************************************************
'[ �֐��� ]�@ArrangeRange
'[ �T�@�v ]�@Select�o����Z���ɐ�������A�̈�̏d�����Ȃ���
'[ ���@�� ]�@�Ώۗ̈�
'[ �߂�l ]�@���������̈�
'*****************************************************************************
Public Function ArrangeRange(ByRef objRange As Range) As Range
    Dim objArea      As Range
    
    If objRange Is Nothing Then
        Exit Function
    End If
    
    '�̈悲�Ƃɐ�������
    For Each objArea In objRange.Areas
        Set ArrangeRange = UnionRange(ArrangeRange, ArrangeRange2(objArea))
    Next
    
    '�Ō�̃Z���ȍ~�̗̈�𑫂�
    Set ArrangeRange = UnionRange(ArrangeRange, MinusRange(objRange, GetUsedRange(objRange.Worksheet)))
End Function

'*****************************************************************************
'[ �֐��� ]�@ArrangeRange2
'[ �T�@�v ]�@Select�o����Z���ɐ�������A�̈�̏d�����Ȃ���
'[ ���@�� ]�@�Ώۗ̈�
'[ �߂�l ]�@���������̈�
'*****************************************************************************
Private Function ArrangeRange2(ByRef objRange As Range) As Range
    Dim objArrange(1 To 3) As Range
    Dim i As Long
    
    If objRange.Count = 1 Then
        Set ArrangeRange2 = objRange.MergeArea
        Exit Function
    End If
    
    If IsOnlyCell(objRange) Then
        Set ArrangeRange2 = objRange
        Exit Function
    End If
    
    With objRange
        On Error Resume Next
        '���ׂẴZ���������Z���ɉ����đI������
        Set objArrange(1) = .SpecialCells(xlCellTypeConstants)
        Set objArrange(2) = .SpecialCells(xlCellTypeFormulas)
        Set objArrange(3) = .SpecialCells(xlCellTypeBlanks)
        On Error GoTo 0
    End With
    
    For i = 1 To 3
        Set ArrangeRange2 = UnionRange(ArrangeRange2, objArrange(i))
    Next
End Function

'*****************************************************************************
'[ �֐��� ]�@ReverseRange
'[ �T�@�v ]�@�̈�𔽓]����
'[ ���@�� ]�@�Ώۗ̈�
'[ �߂�l ]�@Not objRange
'*****************************************************************************
Private Function ReverseRange(ByRef objRange As Range) As Range
    Dim i As Long
    Dim objRound(1 To 4) As Range
    
    With objRange.Parent
        On Error Resume Next
        '�I��̈����̗̈悷�ׂ�
        Set objRound(1) = .Range(.Rows(1), _
                                 .Rows(objRange.Row - 1))
        '�I��̈��艺�̗̈悷�ׂ�
        Set objRound(2) = .Range(.Rows(objRange.Row + objRange.Rows.Count), _
                                 .Rows(Rows.Count))
        '�I��̈��荶�̗̈悷�ׂ�
        Set objRound(3) = .Range(.Columns(1), _
                                 .Columns(objRange.Column - 1))
        '�I��̈���E�̗̈悷�ׂ�
        Set objRound(4) = .Range(.Columns(objRange.Column + objRange.Columns.Count), _
                                 .Columns(Columns.Count))
        On Error GoTo 0
    End With
    
    '�I��̈�ȊO�̗̈��ݒ�
    For i = 1 To 4
        Set ReverseRange = UnionRange(ReverseRange, objRound(i))
    Next
End Function

'*****************************************************************************
'[ �֐��� ]�@ReSelectRange
'[ �T�@�v ]�@�V�����̈���A���̗̈�̑I�����Ƃ̃G���A�ɕ�������
'�@�@�@�@�@�@��:ReSelectRange(Range("A1,A2,A3"),Range("A1:A2")).Address��"A1,A2"
'[ ���@�� ]�@objSelection:���̗̈�AobjNewRange:�V�����̈�
'[ �߂�l ]  objNewRange�����̗̈�̑I�����Ƃ̃G���A�ɕ�����������
'*****************************************************************************
Public Function ReSelectRange(ByRef objSelection As Range, ByRef objNewRange As Range) As Range
    Dim objTmpRange As Range
    Dim i As Long
    Dim strAddress As String
    Dim strRange   As String
        
    For i = 1 To objSelection.Areas.Count
        Set objTmpRange = IntersectRange(objSelection.Areas(i), objNewRange)
        If Not (objTmpRange Is Nothing) Then
            strRange = objTmpRange.Address(False, False)
            If Not (MinusRange(objTmpRange, Range(strRange)) Is Nothing) Then
                Set ReSelectRange = objNewRange
                Exit Function
            End If
            strAddress = strAddress & strRange & ","
        End If
    Next i
    
    '�����̃J���}���폜
    strAddress = Left$(strAddress, Len(strAddress) - 1)
    If Len(strAddress) < 256 Then
        Set ReSelectRange = Range(strAddress)
    Else
        Set ReSelectRange = objNewRange
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@GetRowMergeRange
'[ �T�@�v ]�@�������ꂽ�̈���擾����
'[ ���@�� ]�@�����^�C�v�A�Ώۗ̈�
'[ �߂�l ]�@�������ꂽ�̈�
'*****************************************************************************
Public Function GetMergeRange(ByRef objSelection As Range, Optional ByVal enmMergeType As EMergeType = E_MTBOTH) As Range
    Dim objRange   As Range
    Dim objCell    As Range
    
    '�������ꂽ�Z����UsedRange�ȊO�ɂ͂Ȃ��̂�
    Set objRange = IntersectRange(objSelection, GetUsedRange())
    If objRange Is Nothing Then
        Exit Function
    End If
    
On Error GoTo ErrHandle
'    If objRange.Count > 100000 Then
'        Call Err.Raise(C_CheckErrMsg, , "�I�����ꂽ�Z�����������܂�")
'    End If
    
    '�Z���̐��������[�v
    For Each objCell In objRange
        With objCell.MergeArea
            '�����Z�����H
            If .Count > 1 Then
                '����̃Z����
                If .Row = objCell.Row And .Column = objCell.Column Then
                    Select Case enmMergeType
                    Case E_MTBOTH
                        Set GetMergeRange = UnionRange(GetMergeRange, objCell)
                    Case E_MTROW
                        If .Rows.Count > 1 Then
                            Set GetMergeRange = UnionRange(GetMergeRange, objCell)
                        End If
                    Case E_MTCOL
                        If .Columns.Count > 1 Then
                            Set GetMergeRange = UnionRange(GetMergeRange, objCell)
                        End If
                    End Select
                End If
            End If
        End With
    Next
Exit Function
ErrHandle:
    Call Err.Raise(C_CheckErrMsg, , "�I�����ꂽ�Z�����������܂�")
End Function

'*****************************************************************************
'[ �֐��� ]�@GetNearlyRange
'[ �T  �v ]  Shape�̎l���ɍł��߂��Z���͈͂��擾����
'[ ��  �� ]�@Shape�I�u�W�F�N�g
'[ �߂�l ]�@�Z���͈�
'*****************************************************************************
Public Function GetNearlyRange(ByRef objShape As Shape) As Range
    Dim objTopLeft     As Range
    Dim objBottomRight As Range
    Set objTopLeft = objShape.TopLeftCell
    Set objBottomRight = objShape.BottomRightCell
    
    '��̈ʒu�ƍ�����ݒ�
    If objShape.Height = 0 Then
        With objTopLeft
            If .Top + .Height / 2 < objShape.Top Then
                Set objTopLeft = Cells(.Row + 1, .Column)
                Set objBottomRight = Cells(.Row + 1, objBottomRight.Column)
            End If
        End With
    Else
        '���̃Z���̍Đݒ�
        With objBottomRight
            If .Top = objShape.Top + objShape.Height Then
                Set objBottomRight = Cells(.Row - 1, .Column)
            End If
        End With
            
        '��[�̍Đݒ�
        With objTopLeft
            If .Top + .Height / 2 < objShape.Top Then
                If .Row + 1 <= objBottomRight.Row Then
                    Set objTopLeft = Cells(.Row + 1, .Column)
                End If
            End If
        End With
                
        '���[�̍Đݒ�
        With objBottomRight
            If .Top + .Height / 2 > objShape.Top + objShape.Height Then
                If .Row - 1 >= objTopLeft.Row Then
                    Set objBottomRight = Cells(.Row - 1, .Column)
                End If
            End If
        End With
    End If
    
    '���̈ʒu�ƕ���ݒ�
    If objShape.Width = 0 Then
        With objTopLeft
            If .Left + .Width / 2 < objShape.Left Then
                Set objTopLeft = Cells(.Row, .Column + 1)
                Set objBottomRight = Cells(objBottomRight.Row, .Column + 1)
            End If
        End With
    Else
        '�E�̃Z���̍Đݒ�
        With objBottomRight
            If .Left = objShape.Left + objShape.Width Then
                Set objBottomRight = Cells(.Row, .Column - 1)
            End If
        End With
    
        '���[�̍Đݒ�
        With objTopLeft
            If .Left + .Width / 2 < objShape.Left Then
                If .Column + 1 <= objBottomRight.Column Then
                    Set objTopLeft = Cells(.Row, .Column + 1)
                End If
            End If
        End With
                
        '�E�[�̍Đݒ�
        With objBottomRight
            If .Left + .Width / 2 > objShape.Left + objShape.Width Then
                If .Column - 1 >= objTopLeft.Column Then
                    Set objBottomRight = Cells(.Row, .Column - 1)
                End If
            End If
        End With
    End If
    
    Set GetNearlyRange = Range(objTopLeft, objBottomRight)
End Function

'*****************************************************************************
'[ �֐��� ]�@CheckDupRange
'[ �T  �v ]�@�̈�ɏd�����Ȃ����ǂ������肷��
'[ ��  �� ]�@���肷��̈�
'[ �߂�l ]�@True�F�d������
'*****************************************************************************
Public Function CheckDupRange(ByRef objAreas As Range) As Boolean
    Dim objRange   As Range
    Dim objWkRange As Range
    
    For Each objRange In objAreas.Areas
        If IntersectRange(objWkRange, objRange) Is Nothing Then
            Set objWkRange = UnionRange(objWkRange, objRange)
        Else
            CheckDupRange = True
            Exit Function
        End If
    Next objRange
End Function

'*****************************************************************************
'[ �֐��� ]�@SearchValueCell
'[ �T  �v ]�@�l�̓��͂���Ă���Z������������
'[ ��  �� ]�@objRange�F�����͈�
'[ �߂�l ]�@�l�̓��͂���Ă���Z��
'*****************************************************************************
Public Function SearchValueCell(ByRef objRange As Range) As Range
    Dim objWkRange(0 To 1)  As Range
    
    On Error Resume Next
    With objRange
        Set objWkRange(0) = .SpecialCells(xlCellTypeConstants)
        Set objWkRange(1) = .SpecialCells(xlCellTypeFormulas)
    End With
    On Error GoTo 0
    Set SearchValueCell = UnionRange(objWkRange(0), objWkRange(1))
End Function

'*****************************************************************************
'[ �֐��� ]�@GetSheeetShapeRange
'[ �T  �v ]�@���[�N�V�[�g��Shpes�I�u�W�F�N�g��ShapeRange�I�u�W�F�N�g�ɕϊ�
'[ ��  �� ]�@���[�N�V�[�g
'[ �߂�l ]�@ShapeRange�I�u�W�F�N�g
'*****************************************************************************
Public Function GetSheeetShapeRange(ByRef objSheet As Worksheet) As ShapeRange
    Dim i As Long
    If objSheet.Shapes.Count = 0 Then
        Exit Function
    End If
    ReDim lngArray(1 To objSheet.Shapes.Count)
    For i = 1 To objSheet.Shapes.Count
        lngArray(i) = i
    Next
    Set GetSheeetShapeRange = objSheet.Shapes.Range(lngArray)
End Function

'*****************************************************************************
'[ �֐��� ]�@GetMergeAddress
'[ �T  �v ]�@�����Z��1�����̎��A����̃A�h���X�����Ԃ�Ȃ��̂ŁA�S�̂�Ԃ�
'[ ��  �� ]�@�ΏۃA�h���X
'[ �߂�l ]�@�A�h���X
'*****************************************************************************
Public Function GetMergeAddress(ByVal strAddress As String) As String
    GetMergeAddress = strAddress
    With Range(strAddress)
        If .Rows.Count = 1 And .Columns.Count = 1 Then
            With .MergeArea
                If .Count > 1 Then
                    GetMergeAddress = .Address(0, 0)
                End If
            End With
        End If
    End With
End Function

'*****************************************************************************
'[ �֐��� ]�@StrConvert
'[ �T  �v ]�@������̕ϊ����s��
'[ ��  �� ]�@�ϊ��O�̕�����A�ϊ����
'[ �߂�l ]�@�ύX��̕�����
'*****************************************************************************
Public Function StrConvert(ByVal strText As String, ByVal strCommand As String) As String
    StrConvert = strText
    Select Case strCommand
    Case "UpperCase"  '�啶���ɕϊ�
        StrConvert = StrConv(StrConvert, vbUpperCase)
    Case "LowerCase"  '�������ɕϊ�
        StrConvert = StrConv(StrConvert, vbLowerCase)
    Case "ProperCase" '�擪�̂ݑ啶���ɕϊ�
        StrConvert = StrConv(StrConvert, vbProperCase)
    Case "Hiragana" '�Ђ炪�Ȃɕϊ�
        StrConvert = StrConv(StrConvert, vbHiragana)
    Case "Katakana" '�J�^�J�i�ɕϊ�
        StrConvert = StrConv(StrConvert, vbKatakana)
    Case "Wide"     '�S�p�ɕϊ�
        StrConvert = Replace(StrConvert, """", Chr(&H8168))
        StrConvert = Replace(StrConvert, "'", "�f")
        StrConvert = Replace(StrConvert, "\", "��")
        StrConvert = StrConv(StrConvert, vbWide)
    Case "Narrow"   '���p�ɕϊ�
        StrConvert = Replace(StrConvert, "�`", Chr(1) & "~")
        StrConvert = StrConv(StrConvert, vbNarrow)
        StrConvert = Replace(StrConvert, Chr(1) & "~", "�`")
    Case "NarrowExceptKana" '�J�^�J�i�ȊO���p�ɕϊ�
        StrConvert = Replace(StrConvert, "�`", Chr(1) & "~")
        StrConvert = StrConvNarrowExceptKana(StrConvert)
        StrConvert = Replace(StrConvert, Chr(1) & "~", "�`")
    Case "WideOnlyKana" '�J�^�J�i�̂ݑS�p�ɕϊ�
        StrConvert = StrConvWideOnlyKana(StrConvert)
    Case "Trim" '�O��̋󔒂��폜
        StrConvert = Trim(StrConvert)
    Case "RTrim" '�����̋󔒂��폜
        StrConvert = Rtrim(StrConvert)
    End Select
End Function

'*****************************************************************************
'[ �֐��� ]�@StrConvNarrowExceptKana
'[ �T  �v ]�@�J�^�J�i�ȊO���p�ɕϊ�
'[ ��  �� ]�@�ϊ��O�̕�����
'[ �߂�l ]�@�ϊ���̕�����
'*****************************************************************************
Private Function StrConvNarrowExceptKana(ByVal strText As String) As String
    Dim i       As Long
    Dim strChar As String
    
    '�����������[�v
    For i = 1 To Len(strText)
        strChar = Mid$(strText, i, 1)
        Select Case AscW(strChar)
        Case AscW("�A") To AscW("��"), AscW("�@"), AscW("��"), _
             AscW("�["), AscW("�A"), AscW("�B")
            StrConvNarrowExceptKana = StrConvNarrowExceptKana & strChar
        Case Else
            StrConvNarrowExceptKana = StrConvNarrowExceptKana & StrConv(strChar, vbNarrow)
        End Select
    Next
End Function

'*****************************************************************************
'[ �֐��� ]�@StrConvWideOnlyKana
'[ �T  �v ]�@�J�^�J�i�̂ݑS�p�ɕϊ�
'[ ��  �� ]�@�ϊ��O�̕�����
'[ �߂�l ]�@�ϊ���̕�����
'*****************************************************************************
Private Function StrConvWideOnlyKana(ByVal strText As String) As String
    Dim i           As Long
    Dim strChar     As String
    Dim strWideChar As String
    
    '�擪��(��)���_�̎��̑Ή��Ƃ��āA�擪�ɓK���ȕ�����ǉ����Ă���
    strText = "?" & strText
    
    '��������������烋�[�v�@����L�Őݒ肵���擪�����͑ΏۊO
    For i = Len(strText) To 2 Step -1
        strChar = Mid$(strText, i, 1)
        Select Case AscW(strChar)
        Case AscW("�") To AscW("�"), AscW("�") To AscW("�"), AscW("�"), _
             AscW("�"), AscW("�"), AscW("�")
            StrConvWideOnlyKana = StrConv(strChar, vbWide) & StrConvWideOnlyKana
        Case AscW("�"), AscW("�")
            strWideChar = StrConv(Mid$(strText, i - 1, 2), vbWide)
            If Len(strWideChar) = 1 Then
                '��F�� �� �K
                StrConvWideOnlyKana = strWideChar & StrConvWideOnlyKana
                i = i - 1
            Else
                '��F�(���p) �� �J(�S�p)
                StrConvWideOnlyKana = StrConv(strChar, vbWide) & StrConvWideOnlyKana
            End If
        Case Else
            StrConvWideOnlyKana = strChar & StrConvWideOnlyKana
        End Select
    Next
End Function

'*****************************************************************************
'[�T�v] SortArray�z����\�[�g����
'[����] Sort�Ώ۔z��
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub SortArray(ByRef SortArray() As TSortArray)
    '�o�u���\�[�g
    Dim i As Long
    Dim j As Long
    Dim Swap As TSortArray
    For i = UBound(SortArray) To 1 Step -1
        For j = 1 To i - 1
            If CompareValue(SortArray(j), SortArray(j + 1)) Then
                Swap = SortArray(j)
                SortArray(j) = SortArray(j + 1)
                SortArray(j + 1) = Swap
            End If
        Next j
    Next i
    
    '�X�^�b�N�̈�s�����N����
'    Call QuickSort(SortArray, 1, UBound(SortArray))
End Sub

'Private Sub QuickSort(ByRef SortArray() As TSortArray, ByVal lngMin As Long, ByVal lngMax As Long)
'    Dim i As Long
'    Dim j As Long
'    Dim Base As TSortArray
'    Dim Swap As TSortArray
'    Base = SortArray(Int((lngMin + lngMax) / 2))
'    i = lngMin
'    j = lngMax
'    Do
'        Do While CompareValue(SortArray(i), Base)
'            i = i + 1
'        Loop
'        Do While Not CompareValue(SortArray(j), Base)
'            j = j - 1
'        Loop
'        If i >= j Then Exit Do
'            Swap = SortArray(i)
'            SortArray(i) = SortArray(j)
'            SortArray(j) = Swap
'        i = i + 1
'        j = j - 1
'    Loop
'    If (lngMin < i - 1) Then
'        Call QuickSort(SortArray, lngMin, i - 1)
'    End If
'    If (lngMax > j + 1) Then
'        Call QuickSort(SortArray, j + 1, lngMax)
'    End If
'End Sub

'*****************************************************************************
'[�T�v] �召��r���s��
'[����] ��r�Ώ�
'[�ߒl] True: SortArray1 > SortArray2
'*****************************************************************************
Public Function CompareValue(ByRef SortArray1 As TSortArray, ByRef SortArray2 As TSortArray) As Boolean
    If SortArray1.Key1 = SortArray2.Key1 Then
        CompareValue = (SortArray1.Key2 > SortArray2.Key2)
    Else
        CompareValue = (SortArray1.Key1 > SortArray2.Key1)
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@IsBorderMerged
'[ �T  �v ]�@Range�̋��E�������Z�����ǂ���
'[ ��  �� ]�@���肷��Range
'[ �߂�l ]�@True:���E�Ɍ����Z������AFalse:�Ȃ�
'*****************************************************************************
Public Function IsBorderMerged(ByRef objRange As Range) As Boolean
    IsBorderMerged = Not (MinusRange(ArrangeRange(objRange), objRange) Is Nothing)
End Function

'*****************************************************************************
'[�T�v] Range��(�������ꂽ)�P��̃Z�����ǂ���
'[����] ���肷��Range
'[�ߒl] True:�P��̃Z���AFalse:�����̃Z��
'*****************************************************************************
Public Function IsOnlyCell(ByRef objRange As Range) As Boolean
    If objRange.Areas.Count > 1 Then
        Exit Function
    End If
    If objRange.Count = 1 Then
        IsOnlyCell = True
        Exit Function
    End If
    
    IsOnlyCell = (objRange.Address = objRange(1, 1).MergeArea.Address)
End Function

'*****************************************************************************
'[ �֐��� ]�@GetUsedRange
'[ �T  �v ]�@�g�p����Ă���̈���擾����
'[ ��  �� ]�@���肷��Range
'[ �߂�l ]�@�g�p����Ă���̈�
'*****************************************************************************
Public Function GetUsedRange(Optional ByRef objSheet As Worksheet = Nothing) As Range
    On Error Resume Next
    If objSheet Is Nothing Then
        Set GetUsedRange = Range(Cells(1, 1), Cells.SpecialCells(xlCellTypeLastCell))
    Else
        With objSheet
            Set GetUsedRange = .Range(.Cells(1, 1), .Cells.SpecialCells(xlCellTypeLastCell))
        End With
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@OffsetRange
'[ �T�@�v ]�@Range��Offset���ړ�����
'            �����Z��������ƒP�Ȃ�Offset���\�b�h���z��O�̓�������邽��
'[ ���@�� ]�@���̗̈�A�s��Offset�A���Offset
'[ �߂�l ]�@�}�`���X���C�h������̈�
'*****************************************************************************
Public Function OffsetRange(ByRef objRange As Range, Optional ByVal lngRowOffset As Long = 0, Optional ByVal lngColOffset As Long = 0) As Range
    Dim objCell(1 To 2) As Range '1:����A2:�E��
    
    With objRange(1)
        Set objCell(1) = objRange.Worksheet.Cells(.Row + lngRowOffset, .Column + lngColOffset)
    End With
    With objRange(objRange.Count)
        Set objCell(2) = objRange.Worksheet.Cells(.Row + lngRowOffset, .Column + lngColOffset)
    End With
    
    Set OffsetRange = objRange.Worksheet.Range(objCell(1), objCell(2))
End Function

'*****************************************************************************
'[ �֐��� ]�@ReSizeRange
'[ �T�@�v ]�@Range��Offset���g��k������
'            �����Z��������ƒP�Ȃ�Resize���\�b�h���z��O�̓�������邽��
'[ ���@�� ]�@���̗̈�A�s��Offset�A���Offset
'[ �߂�l ]�@�}�`���g��k��������̈�
'*****************************************************************************
Public Function ReSizeRange(ByRef objRange As Range, Optional ByVal lngRowOffset As Long = 0, Optional ByVal lngColOffset As Long = 0) As Range
    Dim objCell(1 To 2) As Range '1:����A2:�E��
    
    Set objCell(1) = objRange(1)
    With objRange(objRange.Count)
        Set objCell(2) = objRange.Worksheet.Cells(.Row + lngRowOffset, .Column + lngColOffset)
    End With
    
    Set ReSizeRange = objRange.Worksheet.Range(objCell(1), objCell(2))
End Function

'*****************************************************************************
'[ �֐��� ]�@GetClipbordText
'[ �T  �v ]�@�N���b�v�{�[�h�̃e�L�X�g���擾����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Function GetClipbordText() As String
On Error GoTo ErrHandle
    Dim objCb As New DataObject
    Call objCb.GetFromClipboard
    
    '�e�L�X�g�`�����ێ�����Ă��鎞
    If objCb.GetFormat(1) Then
        GetClipbordText = objCb.GetText
    End If
ErrHandle:
    Set objCb = Nothing
End Function

'*****************************************************************************
'[ �֐��� ]�@SetClipbordText
'[ �T  �v ]�@�N���b�v�{�[�h�Ƀe�L�X�g��ݒ肷��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetClipbordText(ByVal strText As String)
On Error GoTo ErrHandle
    Dim objCb As New DataObject
    Call objCb.Clear
    Call objCb.SetText(strText)
    Call objCb.PutInClipboard
ErrHandle:
    Set objCb = Nothing
End Sub

'*****************************************************************************
'[ �֐��� ]�@ClearClipbord
'[ �T  �v ]�@�N���b�v�{�[�h�̃N���A
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub ClearClipbord()
On Error GoTo ErrHandle
    Dim objCb As New DataObject
    Call objCb.Clear
    Call objCb.SetText("")
    Call objCb.PutInClipboard
ErrHandle:
    Set objCb = Nothing
End Sub

'*****************************************************************************
'[ �֐��� ]�@DeleteSheet
'[ �T  �v ]�@���[�N�V�[�g�̒��g���폜����
'[ ��  �� ]�@�Ώۂ̃��[�N�V�[�g
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub DeleteSheet(ByRef objSheet As Worksheet)
    Dim objShape  As Shape
    For Each objShape In objSheet.Shapes
        Call objShape.Delete
    Next objShape
    
    With objSheet.Cells
        Call .Delete
    End With

    objSheet.StandardWidth = 8

    '�Ō�̃Z�����C������
    Call objSheet.Cells.Parent.UsedRange
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetIMEOff
'[ �T  �v ]�@�h�l�d���I�t�ɂ���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetIMEOff()
On Error GoTo ErrHandle
    Dim hIME As LongPtr
    hIME = ImmGetContext(Application.hwnd)
    Call ImmSetOpenStatus(hIME, 0)
ErrHandle:
    If hIME <> 0 Then
        Call ImmReleaseContext(Application.hwnd, hIME)
    End If
End Sub

'*****************************************************************************
'[ �T  �v ]�@DPI�̕ϊ������擾���� ��72(Excel�̃f�t�H���g��DPI)/��ʂ�DPI
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Function DPIRatio() As Double
    DPIRatio = 72 / GetDPI()
End Function

'*****************************************************************************
'[ �T  �v ]�@DPI���擾����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@DPI ���W����96
'*****************************************************************************
Public Function GetDPI() As Long
    Dim DC As LongPtr
    DC = GetDC(0)
    GetDPI = GetDeviceCaps(DC, LOGPIXELSX)
    Call ReleaseDC(0, DC)
End Function

'*****************************************************************************
'[�T�v] �R�s�[�Ώۂ�Range���擾����
'[����] �Ȃ�
'[�ߒl] �R�s�[�Ώۂ�Range
'*****************************************************************************
Public Function GetCopyRange() As Range
    If OpenClipboard(0) = 0 Then Exit Function
    Dim hMem As LongPtr
    hMem = GetClipboardData(RegisterClipboardFormat("Link"))
    If hMem = 0 Then
        Call CloseClipboard
        Exit Function
    End If
     
On Error GoTo ErrHandle
    Dim Size As Long
    Dim p As LongPtr
    Size = GlobalSize(hMem)
    p = GlobalLock(hMem)
    ReDim Data(1 To Size) As Byte
    Call CopyMemory(Data(1), ByVal p, Size)
    Call GlobalUnlock(hMem)
    Call CloseClipboard
    hMem = 0
    
    Dim strData As String
    Dim i As Long
    For i = 1 To Size
        If Data(i) = 0 Then
            Data(i) = Asc("/") '�V�[�g���ɂ��t�@�C�����ɂ��g���Ȃ�����
        End If
    Next i
    strData = StrConv(Data, vbUnicode)
    
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = False
    
    Dim strAddress As String
    Dim strFileAndSheet As String
    '��FExcel/Z:\[Book1.xlsx]Sheet1/R1C1:R2C2//
    objRegExp.Pattern = "Excel\/(.+)\/(R[0-9]+C[0-9]+(:R[0-9]+C[0-9]+)?)\/\/$"
    If Not objRegExp.Test(strData) Then Exit Function
    With objRegExp.Execute(strData)(0)
        '��FZ:\[Book1.xlsx]Sheet1
        strFileAndSheet = .SubMatches(0)
        '��FR1C1:R2C2
        strAddress = .SubMatches(1)
    End With
    
    Dim strBook  As String
    Dim strSheet As String
    '��FZ:\[Book1.xlsx]Sheet1
    objRegExp.Pattern = "\[(.+)\](.+)$"
    If objRegExp.Test(strFileAndSheet) Then
        With objRegExp.Execute(strFileAndSheet)(0)
            '��FBook1.xlsx
            strBook = .SubMatches(0)
            '��FSheet1
            strSheet = .SubMatches(1)
        End With
    Else
        If vbYes = MsgBox("�R�s�[���̃t�@�C���p�X�ƃV�[�g���̍��v�̕����񂪒������邽�߃R�s�[���̃Z���͈͂̎擾���ł��܂���ł���" & vbLf & "�t�@�C���p�X�ƃV�[�g������͂��܂����H", vbQuestion & vbYesNo) Then
            strBook = InputBox("�R�s�[���̃t�@�C��������͂��Ă��������B" & vbLf & "�\��t����Ɠ����t�@�C���̏ꍇ�͏ȗ���")
            strSheet = InputBox("�R�s�[���̃V�[�g������͂��Ă��������B" & vbLf & "�\��t����Ɠ����V�[�g�̏ꍇ�͏ȗ���")
        Else
            Exit Function
        End If
    End If
    
    Dim strRange As String
    strRange = Application.ConvertFormula(strAddress, xlR1C1, xlA1)
    If strBook = "" Then
        If strSheet = "" Then
            Set GetCopyRange = Range(strRange)
        Else
            Set GetCopyRange = Worksheets(strSheet).Range(strRange)
        End If
    Else
        Set GetCopyRange = Workbooks(strBook).Worksheets(strSheet).Range(strRange)
    End If
    
    If IsOnlyCell(GetCopyRange) Then
        Set GetCopyRange = GetCopyRange.MergeArea
    End If
    
    Application.CutCopyMode = False
    Exit Function
ErrHandle:
    If hMem <> 0 Then Call CloseClipboard
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'*****************************************************************************
'[�T�v] �g���q�̎擾
'[����] �t�@�C���p�X
'[�ߒl] �g���q(�啶��)
'*****************************************************************************
Public Function GetFileExtension(ByVal strFilename As String) As String
    With CreateObject("Scripting.FileSystemObject")
        GetFileExtension = UCase(.GetExtensionName(strFilename))
    End With
End Function

'*****************************************************************************
'[�T�v] Undo�{�^���̏����擾����
'[����] �Ȃ�
'[�ߒl] Undo�{�^����TooltipText
'*****************************************************************************
Public Function GetUndoStr() As String
    With CommandBars.FindControl(, 128) 'Undo�{�^��
        If .Enabled Then
            If .ListCount = 1 Then
                'Undo��1��ނ̎���Undo�R�}���h
                GetUndoStr = Trim(.List(1))
            End If
        End If
    End With
End Function

'*****************************************************************************
'[�T�v] 255���ȏ�̃A�h���X�ł��擾�ł���悤�ɂ���
'[����] Range
'[�ߒl] 255���ȏ�ɑΉ������A�h���X
'*****************************************************************************
Public Function GetAddress(ByRef objAreas As Range) As String
    Dim objRange As Range
    For Each objRange In objAreas.Areas
        If GetAddress = "" Then
            GetAddress = objRange.Address(0, 0)
        Else
            GetAddress = GetAddress & "," & objRange.Address(0, 0)
        End If
    Next
End Function

'*****************************************************************************
'[�T�v] 255���ȏ�ɑΉ������A�h���X��Range���擾
'[����] 255���ȏ�ɑΉ������A�h���X
'[�ߒl] Range
'*****************************************************************************
Public Function GetRange(ByVal strAddress As String) As Range
    Const MAXLEN = 250
    Dim strText As String
    Dim i As Long
    
    While Len(strAddress) > 0
        If Len(strAddress) >= MAXLEN Then
            For i = MAXLEN To 1 Step -1
                If Mid(strAddress, i, 1) = "," Then
                    strText = Left(strAddress, i - 1)
                    Set GetRange = UnionRange(GetRange, Range(strText))
                    strAddress = Mid(strAddress, i + 1)
                    Exit For
                End If
            Next
        Else
            Set GetRange = UnionRange(GetRange, Range(strAddress))
            strAddress = ""
        End If
    Wend
End Function

'*****************************************************************************
'[�T�v] �I�u�W�F�N�g�̑I����\����ʂ�\������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ShowSelectionPane()
    On Error Resume Next
    If Not CommandBars.GetPressedMso("SelectionPane") Then
        Call CommandBars.ExecuteMso("SelectionPane")
    End If
End Sub


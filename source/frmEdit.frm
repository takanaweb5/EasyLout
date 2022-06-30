VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEdit 
   Caption         =   "���񂽂񃌃C�A�E�g"
   ClientHeight    =   3972
   ClientLeft      =   48
   ClientTop       =   228
   ClientWidth     =   7464
   OleObjectBlob   =   "frmEdit.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hwnd       As LongPtr
'Private OrgWndProc As Long
Private blnZoomed  As Boolean
Private objTmpBar  As CommandBar
Private dblAnchor1T  As Double
Private dblAnchor1L  As Double
Private dblAnchor2B  As Double
Private dblAnchor2R  As Double
Private dblAnchor3T  As Double
Private dblAnchor3L  As Double

Private IsBoxSelect As Boolean
Private FText       As String
Private FSelRow     As Long '��`�I���J�n�s
Private FSelRowCnt  As Long '��`�I���s��
Private FSelCol     As Long '��`�I���J�n��(�o�C�g�P��)
Private FSelLen     As Long '��`�I������(�o�C�g�P��)

'*****************************************************************************
'[�C�x���g]�@UserForm_Initialize
'[ �T  �v ]�@�t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    Dim lngStyle As Long
    Dim i        As Long
    
    Set imgGrip.Picture = Nothing
    dblAnchor1T = Me.Height - cmdCancel.Top
    dblAnchor1L = Me.Width - cmdCancel.Left
    dblAnchor2B = Me.Height - txtEdit.Height
    dblAnchor2R = Me.Width - txtEdit.Width
    dblAnchor3T = Me.Height - imgGrip.Top
    dblAnchor3L = Me.Width - imgGrip.Height
    
    '********************************************
    '�E�B���h�E�̃T�C�Y��ύX�o����悤�ɕύX
    '********************************************
    hwnd = FindWindow("ThunderDFrame", Me.Caption)
    lngStyle = GetWindowLong(hwnd, GWL_STYLE)
    Call SetWindowLong(hwnd, GWL_STYLE, lngStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX)

'    '********************************************
'    '�T�u�N���X�����ă}�E�X�z�C�[����L���ɂ���
'    '********************************************
'    OrgWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubClassProc)
        
    '********************************************
    '�t�H�[���̏�����Ԃ�ݒ�
    '********************************************
    With txtEdit
        .MultiLine = True
        .WordWrap = False
        .ScrollBars = fmScrollBarsBoth
        .SelectionMargin = True
        If IsOnlyCell(Selection) Then
            .Text = ActiveCell.Value
        Else
            .Text = GetRangeText(Selection)
        End If
        chkWordWrap = .WordWrap
        .SelStart = 0
    End With
    
    '********************************************
    '�E�N���b�N���j���[�쐬
    '********************************************
    Set objTmpBar = CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
    With objTmpBar.Controls
        With .Add()
            .Caption = "��`�I��"
        End With
        With .Add()
            .BeginGroup = True
            .Caption = "���ɖ߂�(&U)�@�@�@Ctrl+Z"
        End With
        With .Add()
            .Caption = "��蒼��(&F)�@Ctrl+Shift+Z"
        End With
        With .Add(, 21)
            .BeginGroup = True
        End With
        With .Add(, 19)
        End With
        With .Add(, 22)
        End With
        With .Add()
            .Caption = "�폜(&D)"
        End With
        With .Add()
            .BeginGroup = True
            .Caption = "���ׂđI��(&A)�@�@Ctrl+A"
        End With
        With .Add()
            .BeginGroup = True
            .Caption = "�啶���ɕϊ�"
        End With
        With .Add()
            .Caption = "�������ɕϊ�"
        End With
        With .Add()
            .Caption = "�擪�̂ݑ啶���ɕϊ�"
        End With
        With .Add()
            .Caption = "�Ђ炪�Ȃɕϊ�"
        End With
        With .Add()
            .Caption = "�J�^�J�i�ɕϊ�"
        End With
        With .Add()
            .Caption = "�S�p�ɕϊ�"
        End With
        With .Add()
            .Caption = "���p�ɕϊ�"
        End With
        With .Add()
            .Caption = "�J�^�J�i�ȊO���p�ɕϊ�"
        End With
        With .Add()
            .Caption = "�J�^�J�i�̂ݑS�p�ɕϊ�"
        End With
    End With

    For i = 1 To objTmpBar.Controls.Count
        objTmpBar.Controls(i).onAction = "OnPopupClick2"
        objTmpBar.Controls(i).Tag = i
    Next i
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_Terminate
'[ �T  �v ]�@�f�X�g���N�^
'*****************************************************************************
Private Sub UserForm_Terminate()
'    '�E�B���h�E�v���V�W���[�����ɂ��ǂ�
'    If OrgWndProc <> 0 Then
'        Call SetWindowLong(hWnd, GWL_WNDPROC, OrgWndProc)
'    End If
    
    '�E�N���b�N���j���[�폜
    Call objTmpBar.Delete
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_QueryClose
'[ �T  �v ]�@�~�{�^���Ńt�H�[������鎞�A�t�H�[����j�������Ȃ�
'*****************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '�~�{�^���Ńt�H�[������鎞�A�t�H�[����j�������Ȃ�
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        blnZoomed = IsZoomed(hwnd)
        Me.Hide
    End If
    
    '�E�B���h�E�̃T�C�Y�����ɖ߂�
    Call ShowWindow(hwnd, SW_SHOWNORMAL)
End Sub
'*****************************************************************************
'[�C�x���g]�@cmdOK_Click
'[ �T  �v ]�@�n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrHandle
    If IsBoxSelect Then
        Call BoxSelect(False, True)
        Exit Sub
    End If
    
    Dim strText         As String
    Dim objOldSelection As Range
    Dim objNewSelection As Range
    
    '���s��CRLF��LF
    strText = Replace$(txtEdit.Text, vbCr, "")
    
    If IsOnlyCell(Selection) Then
        Call SaveUndoInfo(E_CellValue, ActiveCell.MergeArea)
        ActiveCell.Value = Replace$(strText, vbTab, "")
        Set objNewSelection = Selection
    Else
        Set objOldSelection = Selection
        Set objNewSelection = GetPasteRange(strText, Selection)
        Call SaveUndoInfo(E_CellValue, objOldSelection)
        Call objOldSelection.ClearContents
        Call PasteTabText(strText, objNewSelection)
    End If
    Call objNewSelection.Select
    Call SetOnUndo
    Call objNewSelection.Select
ErrHandle:
    blnZoomed = IsZoomed(hwnd)
    Me.Hide
    '�E�B���h�E�̃T�C�Y�����ɖ߂�
    Call ShowWindow(hwnd, SW_SHOWNORMAL)
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdCancel_Click
'[ �T  �v ]�@�L�����Z���{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    If IsBoxSelect Then
        Call BoxSelect(False, False)
        Exit Sub
    End If
    
    blnZoomed = IsZoomed(hwnd)
    Me.Hide
    '�E�B���h�E�̃T�C�Y�����ɖ߂�
    Call ShowWindow(hwnd, SW_SHOWNORMAL)
End Sub

'*****************************************************************************
'[�C�x���g]�@SpbSize_Change
'[ �T  �v ]�@�t�H���g�T�C�Y�ύX��
'*****************************************************************************
Private Sub SpbSize_Change()
    txtSize.Text = CStr(SpbSize.Value)
    txtEdit.Font.Size = SpbSize.Value
End Sub

'*****************************************************************************
'[�C�x���g]�@SpbSize_KeyDown
'[ �T  �v ]�@ESC�L�[�Ńt�H���g�T�C�Y�̕ύX������������Ȃ��悤�ɂ���
'*****************************************************************************
Private Sub SpbSize_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Call cmdCancel_Click
        Exit Sub
    End If

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        KeyCode = 0
        Call txtEdit.SetFocus
        Exit Sub
    End If
End Sub

'*****************************************************************************
'[�C�x���g]�@SpbSize_Exit
'[ �T  �v ]�@�t�H���g�T�C�Y�ύX��̐����X�N���[���o�[��\�����邽�߂̂��܂��Ȃ�
'*****************************************************************************
Private Sub SpbSize_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call txtEdit.SetFocus
    Call Me.Repaint
End Sub

'*****************************************************************************
'[�C�x���g]�@chkWordWrap_Click
'[ �T  �v ]�@�E�[�̐܂�Ԃ��̗L����ύX���܂�
'*****************************************************************************
Private Sub chkWordWrap_Change()
    txtEdit.WordWrap = chkWordWrap
    If chkWordWrap Then
        txtEdit.ScrollBars = fmScrollBarsVertical
    Else
        txtEdit.ScrollBars = fmScrollBarsBoth
    End If
    Call txtEdit.SetFocus
    Call Me.Repaint
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_Resize
'[ �T  �v ]�@�t�H�[���̃T�C�Y�ύX��
'*****************************************************************************
Private Sub UserForm_Resize()
    If Me.Width < 365 Then
        Me.Width = 365
    End If
    
    cmdCancel.Top = Me.Height - dblAnchor1T
    cmdCancel.Left = Me.Width - dblAnchor1L
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - 10 - cmdOK.Width
    txtEdit.Width = Me.Width - dblAnchor2R
    txtEdit.Height = Me.Height - dblAnchor2B
    frmFontSize.Top = cmdCancel.Top
    SpbSize.Top = cmdCancel.Top
    chkWordWrap.Top = cmdCancel.Top + 1
    
    imgGrip.Top = Me.Height - dblAnchor3T
    imgGrip.Left = Me.Width - dblAnchor3L
End Sub

'*****************************************************************************
'[�C�x���g]�@txtEdit_KeyDown
'[ �T  �v ]�@Ctrl+Return �܂��� Alt+Return �łn�j�{�^����������
'*****************************************************************************
Private Sub txtEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '�E�N���b�N���j���[��\������
    If KeyCode = 93 Then
        Call txtEdit_MouseUp(2, 0, 0, 0)
        KeyCode = 0
        Exit Sub
    End If
    
    If Shift = 2 Or Shift = 4 Then
        If KeyCode = vbKeyReturn Then
            Call cmdOK_Click
            Exit Sub
        End If
    End If

    'Ctrl+Shift+Z ��Redo
    If Shift = 3 And KeyCode = vbKeyZ Then
        Call Me.RedoAction
        KeyCode = 0
        Exit Sub
    End If
End Sub

'*****************************************************************************
'[�C�x���g]�@txtEdit_MouseUp
'[ �T  �v ]�@�E�N���b�N���j���[��\������
'*****************************************************************************
Private Sub txtEdit_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 2 Then '�E�{�^��
        objTmpBar.Controls(1).Enabled = (txtEdit.SelLength > 0 And InStr(1, txtEdit.SelText, vbCr) > 0)
        objTmpBar.Controls(2).Enabled = Me.CanUndo
        objTmpBar.Controls(3).Enabled = Me.CanRedo
        objTmpBar.Controls(4).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(5).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(6).Enabled = txtEdit.CanPaste
        objTmpBar.Controls(7).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(8).Enabled = (txtEdit.Text <> "")
        objTmpBar.Controls(9).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(10).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(11).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(12).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(13).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(14).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(15).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(16).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(17).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.ShowPopup
    End If
End Sub

'*****************************************************************************
'[�C�x���g]�@imgGrip_MouseDown
'[ �T  �v ]�@�t�H�[���̉E���Ńt�H�[���̃T�C�Y��ύX�o����悤�ɂ���
'*****************************************************************************
Private Sub imgGrip_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_SYSCOMMAND, SC_SIZE Or 8, 0)
End Sub

'*****************************************************************************
'[ �֐��� ]�@OnPopupClick
'[ �T  �v ]�@�|�b�v�A�b�v���j���[�N���b�N��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub OnPopupClick()
    Dim strNewSelText As String
On Error GoTo ErrHandle
    Select Case CLng(CommandBars.ActionControl.Tag)
    Case 8 To 15
        If Right(txtEdit.SelText, 1) = vbCr Then
            txtEdit.SelLength = txtEdit.SelLength - 1
        End If
    End Select

    Select Case CLng(CommandBars.ActionControl.Tag)
    Case 1 '��`�I��
        Call BoxSelect(True)
    Case 2 '���ɖ߂�
        Call Me.UndoAction
    Case 3 '��蒼��
        Call Me.RedoAction
    Case 4 '�؂���
        Call SendKeys("^x", True)
    Case 5 '�R�s�[
        Call SendKeys("^c", True)
    Case 6 '�\��t��
        Call SendKeys("^v", True)
    Case 7 '�폜
        Call SendKeys("{DEL}", True)
    Case 8 '���ׂđI��
        Call SendKeys("^a", True)
    Case 9  '�啶���ɕϊ�
        strNewSelText = StrConvert(txtEdit.SelText, "UpperCase")
    Case 10  '�������ɕϊ�
        strNewSelText = StrConvert(txtEdit.SelText, "LowerCase")
    Case 11 '�擪�̂ݑ啶���ɕϊ�
        strNewSelText = StrConvert(txtEdit.SelText, "ProperCase")
    Case 12 '�Ђ炪�Ȃɕϊ�
        strNewSelText = StrConvert(txtEdit.SelText, "Hiragana")
    Case 13 '�J�^�J�i�ɕϊ�
        strNewSelText = StrConvert(txtEdit.SelText, "Katakana")
    Case 14 '�S�p�ɕϊ�
        strNewSelText = StrConvert(txtEdit.SelText, "Wide")
    Case 15 '���p�ɕϊ�
        strNewSelText = StrConvert(txtEdit.SelText, "Narrow")
    Case 16 '�J�^�J�i�ȊO���p�ɕϊ�
        strNewSelText = StrConvert(txtEdit.SelText, "NarrowExceptKana")
    Case 17 '�J�^�J�i�̂ݑS�p�ɕϊ�
        strNewSelText = StrConvert(txtEdit.SelText, "WideOnlyKana")
    End Select
    
    'Undo���ł���悤�ɃN���b�v�{�[�h���g�p����
    If strNewSelText <> "" Then
        Dim lngSelStart As Long
        
        '���p�J�^�J�i�́u�ށv�Ȃǂ͕��������ς��̂Œ���
        lngSelStart = txtEdit.SelStart
        Call SetClipbordText(strNewSelText)
        Call SendKeys("^v", True)
        'Excel2019�ł͂��ꂪ�Ȃ��ƁAClearClipbord���Ctrl+V�����s����ĉ����N����Ȃ�
        DoEvents
        txtEdit.SelStart = lngSelStart
        txtEdit.SelLength = Len(strNewSelText)
        
        '�N���b�v�{�[�h�̃N���A
        Call ClearClipbord
    End If
ErrHandle:
End Sub

'*****************************************************************************
'[�v���p�e�B]�@Zoomed
'[ �T  �v ]�@�E�B���h�E�T�C�Y���ő剻����Ă��邩�H
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get Zoomed() As Boolean
    Zoomed = blnZoomed
End Property
Public Property Let Zoomed(ByVal Value As Boolean)
    '�E�B���h�E�T�C�Y���ő剻����
    If ActiveCell.HasFormula = False And Value = True Then
        Call ShowWindow(hwnd, SW_MAXIMIZE)
        Me.Hide
    End If
End Property

'*****************************************************************************
'[�T�v] ��`�I�����[�h���J�n����
'[����] True:��`�I�����[�h�J�n�ATrue:Ok�N���b�N��
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub BoxSelect(ByVal blnSet As Boolean, Optional ByVal blnOk As Boolean = False)
    If blnSet Then
        IsBoxSelect = True
        Me.Caption = "��`�I�����[�h     �P�s�����ҏW�����ꍇ�͂��ׂĂ̍s�ɓK�p����܂�"
        txtEdit.BackColor = &HC0FFFF
        objTmpBar.Controls(1).Visible = False
        
        FSelRow = txtEdit.CurLine + 1
        FText = txtEdit.Text
        txtEdit.Text = GetBoxText(txtEdit.SelText)
        txtEdit.SelStart = 0
        txtEdit.SelLength = Len(txtEdit.Text)
    Else
        IsBoxSelect = False
        Me.Caption = "���񂽂񃌃C�A�E�g"
        txtEdit.BackColor = &H80000005
        objTmpBar.Controls(1).Visible = True
        
        If blnOk Then
            txtEdit.Text = SetBoxText()
        Else
            txtEdit.Text = FText
        End If
    End If
End Sub

'*****************************************************************************
'[�T�v] ��`�I�����ꂽ�e�L�X�g���擾����
'[����] True:��`�I�����[�h�J�n�ATrue:Ok�N���b�N��
'[�ߒl] ��`�I�����ꂽ�e�L�X�g
'*****************************************************************************
Private Function GetBoxText(ByVal strSelText As String) As String
    Dim strAll() As String
    Dim strSel() As String
    Call GetStrArray(FText, strAll())
    Call GetStrArray(strSelText, strSel())
    FSelRowCnt = UBound(strSel) + 1
    
    Dim strLine As String
    Dim strWk   As String
    Dim x(1 To 2) As Long
    strLine = strAll(FSelRow - 1)  '�I��擪�s
    strWk = Mid(strLine, 1, Len(strLine) - Len(strSel(0)))
    x(1) = LenB(StrConv(strWk, vbFromUnicode)) + 1
    strWk = strSel(UBound(strSel)) '�I���ŏI�s
    x(2) = LenB(StrConv(strWk, vbFromUnicode)) + 1
    
    If x(1) <= x(2) Then
        FSelCol = x(1)
        FSelLen = x(2) - x(1)
    Else
        FSelCol = x(2)
        FSelLen = x(1) - x(2)
    End If
    
    Dim i As Long
    For i = FSelRow To FSelRow + FSelRowCnt - 1
        strWk = StrConv(strAll(i - 1), vbFromUnicode)
        GetBoxText = GetBoxText & StrConv(MidB(strWk, FSelCol, FSelLen), vbUnicode) & vbLf
    Next
End Function

'*****************************************************************************
'[�T�v] ��`�ҏW���ꂽ�e�L�X�g�����̃e�L�X�g�ɔ��f����
'[����] �Ȃ�
'[�ߒl] ���f��̃e�L�X�g
'*****************************************************************************
Private Function SetBoxText() As String
On Error Resume Next
    Dim strAll() As String
    Dim strSel() As String
    Call GetStrArray(FText, strAll())
    Call GetStrArray(txtEdit.Text, strSel())

    Dim i As Long
    Dim j As Long
    Dim strWk As String
    For i = FSelRow To FSelRow + FSelRowCnt - 1
        If j > UBound(strSel) Then
            strWk = strSel(0)
        ElseIf UBound(strSel) = 1 And strSel(1) = "" Then
            strWk = strSel(0)
        Else
            strWk = strSel(j)
        End If
        strAll(i - 1) = ExchangeStr(FSelCol, FSelLen, strAll(i - 1), strWk)
        j = j + 1
    Next
    SetBoxText = Join(strAll, vbLf)
End Function

'*****************************************************************************
'[�T�v] ����ʒu�̕������u������
'[����] �u���J�n��(�o�C�g�P��),�u��������(�o�C�g�P��),�u���O�̕�����,�u�����镶����
'[�ߒl] �u����̃e�L�X�g
'*****************************************************************************
Private Function ExchangeStr(ByVal StartCol As Long, ByVal Length As Long, ByVal SrcStr As String, ByVal DstStr As String) As String
    Dim strWk As String
    strWk = StrConv(SrcStr, vbFromUnicode)
    ExchangeStr = StrConv(LeftB(strWk, StartCol - 1), vbUnicode) & DstStr & StrConv(MidB(strWk, StartCol + Length), vbUnicode)
End Function

'*****************************************************************************
'[�T�v] ����������s�ŋ�؂��Ĕz��Ɋi�[����
'[����] ���̕�����,���s�ŋ�؂����z��(�ߒl)
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub GetStrArray(ByVal strText As String, ByRef strArray As Variant)
    strText = Replace$(strText, vbCrLf, vbCr)
    strText = Replace$(strText, vbLf, "")
    strArray = Split(strText, vbCr)
End Sub


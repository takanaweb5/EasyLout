VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSizeList 
   Caption         =   "�T�C�Y�ꗗ"
   ClientHeight    =   4560
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7392
   OleObjectBlob   =   "frmSizeList.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSizeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�^�C�v
Enum EColRowType
    E_Col = 1 '��ύX��
    E_Row = 2 '�s�ύX��
End Enum

'�T�C�Y�ꗗ�̃J�X�^�}�C�Y���u@@@mm �� @@@�s�N�Z���v
Private Type TSizeInfo
    Millimeters As Long
    Pixel       As Long
End Type

Private udtSizeInfo As TSizeInfo
Private colSizeList(1 To 3) As New Collection
Private lngSelNo   As Long '�����p�̑I�����ꂽ�G���ANo
Private blnDisplayPageBreaks As Boolean '���y�[�W�\��
Private lngRatioNo As Long '������ۑ����Ă���G���ANo
Private colAddress As Collection '�������̉򂲂Ƃ̃A�h���X�̔z��(Undo�p)
Private Const C_MAX = 256

'*****************************************************************************
'[�C�x���g]�@UserForm_Initialize
'[ �T  �v ]�@�t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    '�������̂��߉��y�[�W���\���ɂ���
    If ActiveSheet.DisplayAutomaticPageBreaks = True Then
        blnDisplayPageBreaks = True
        ActiveSheet.DisplayAutomaticPageBreaks = False
    End If
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_Terminate
'[ �T  �v ]�@���y�[�W�\���������ɖ߂�
'*****************************************************************************
Private Sub UserForm_Terminate()
    If blnDisplayPageBreaks = True Then
        ActiveSheet.DisplayAutomaticPageBreaks = True
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@Initialize
'[ �T  �v ]  �t�H�[���̏����ݒ���s��
'[ ��  �� ]�@��E�s�@�������ΏۂƂ��邩
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub Initialize(ByVal Value As EColRowType)
    Me.Tag = Value
    
    '��u100mm �� 378�s�N�Z���v
    Call SetSizeInfoLabel
    
    Select Case Value
    Case E_Col
        Call InitCol
        tbsTab(1).Caption = "1��P�ʂ̏��(�S��)"
    Case E_Row
        Call InitRow
        tbsTab(1).Caption = "1�s�P�ʂ̏��(�S��)"
    End Select
    
    Call ReCalc
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetSizeInfoLabel
'[ �T  �v ]  �u@@@mm �� @@@�s�N�Z���v�̐ݒ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub SetSizeInfoLabel()
    Dim strArray() As String
    
    strArray = Split(CommandBars.ActionControl.Tag, ",")
    If UBound(strArray) = 3 Then
        Select Case Me.Tag
        Case E_Col
            If IsNumeric(strArray(0)) And IsNumeric(strArray(1)) Then
                udtSizeInfo.Millimeters = strArray(0)
                udtSizeInfo.Pixel = strArray(1)
            End If
        Case E_Row
            If IsNumeric(strArray(2)) And IsNumeric(strArray(3)) Then
                udtSizeInfo.Millimeters = strArray(2)
                udtSizeInfo.Pixel = strArray(3)
            End If
        End Select
    End If
        
    If udtSizeInfo.Millimeters = 0 Then
        udtSizeInfo.Millimeters = 100
        udtSizeInfo.Pixel = Application.CentimetersToPoints(10) / 0.75
    End If
    
    With udtSizeInfo
        '��u@@@mm �� @@@�s�N�Z���v
        lblSizeInfo.Caption = .Millimeters & "mm �� " & .Pixel & "�s�N�Z��"
    End With
End Sub

'*****************************************************************************
'[ �֐��� ]�@InitCol
'[ �T  �v ]  �񕝃T�C�Y�ύX���̃t�H�[�������ݒ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub InitCol()
    Dim objSelection As Range
    Dim objSizeList  As CSizeList
    Dim objRange     As Range
    Dim lngMin       As Long
    Dim lngNo        As Long
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    
    '***************************************
    '�P��P�ʂ̏����쐬
    '***************************************
    '***********************
    '�G���A���\�[�g����
    '***********************
    '�I��͈͂�Columns�̘a�W�������d�����r������
    Set objSelection = Union(Selection.EntireColumn, Selection.EntireColumn) '@
    
    ReDim lngColNo(1 To objSelection.Areas.Count) As Long '@
    ReDim lngIdxNo(1 To objSelection.Areas.Count) As Long
    
    '�G���A�̐��������[�v
    For i = LBound(lngColNo) To UBound(lngColNo) '@
        lngColNo(i) = objSelection.Areas(i).Column '@
    Next i
    
    '�G���A�̐��������[�v
    For i = LBound(lngIdxNo) To UBound(lngIdxNo)
        lngMin = 9999999
        For j = LBound(lngColNo) To UBound(lngColNo) '@
            If lngColNo(j) < lngMin Then '@
                lngMin = lngColNo(j) '@
                lngNo = j
            End If
        Next j
        lngIdxNo(i) = lngNo
        lngColNo(lngNo) = 9999999 '@
    Next i
    
    '***********************
    '�R���g���[�����쐬����
    '***********************
    k = 0
    '�G���A�̐��������[�v
    For i = LBound(lngIdxNo) To UBound(lngIdxNo)
        With objSelection.Areas(lngIdxNo(i))
            '��̐��������[�v
            For j = 1 To .Columns.Count '@
                k = k + 1
                Set objSizeList = New CSizeList                      '@
                Call objSizeList.CreateSizeList(k, frmPage2)
                Call objSizeList.SetValues(.Columns(j).EntireColumn) '@
                Call colSizeList(2).Add(objSizeList)
        
                '100���ő�Ƃ���
                If k > C_MAX Then
                    Call Err.Raise(513, , C_MAX & "�ȏ�̗�ɑ΂��Ď��s�ł��܂���B")
                End If
            Next j
        End With
    Next i
    
    '***************************************
    '�G���A�P�ʂ̏����쐬
    '***************************************
    k = 0
    '�G���A�̐��������[�v
    For Each objRange In Selection.Areas
        k = k + 1
        Set objSizeList = New CSizeList
        Call objSizeList.CreateSizeList(k, frmPage1)
        Call objSizeList.SetValues(objRange.Columns.EntireColumn) '@
        Call colSizeList(1).Add(objSizeList)
    Next objRange
End Sub

'*****************************************************************************
'[ �֐��� ]�@InitRow
'[ �T  �v ]  �s���T�C�Y�ύX���̃t�H�[�������ݒ�
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub InitRow()
    Dim objSelection As Range
    Dim objSizeList  As CSizeList
    Dim objRange     As Range
    Dim lngMin       As Long
    Dim lngNo        As Long
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    
    '***************************************
    '�P�s�P�ʂ̏����쐬
    '***************************************
    '***********************
    '�G���A���\�[�g����
    '***********************
    '�I��͈͂�Columns�̘a�W�������d�����r������
    Set objSelection = Union(Selection.EntireRow, Selection.EntireRow) '@
    
    ReDim lngRowNo(1 To objSelection.Areas.Count) As Long '@
    ReDim lngIdxNo(1 To objSelection.Areas.Count) As Long
    
    '�G���A�̐��������[�v
    For i = LBound(lngRowNo) To UBound(lngRowNo) '@
        lngRowNo(i) = objSelection.Areas(i).Row  '@
    Next i
    
    '�G���A�̐��������[�v
    For i = LBound(lngIdxNo) To UBound(lngIdxNo)
        lngMin = 9999999
        For j = LBound(lngRowNo) To UBound(lngRowNo) '@
            If lngRowNo(j) < lngMin Then '@
                lngMin = lngRowNo(j) '@
                lngNo = j
            End If
        Next j
        lngIdxNo(i) = lngNo
        lngRowNo(lngNo) = 9999999  '@
    Next i
    
    '***********************
    '�R���g���[�����쐬����
    '***********************
    k = 0
    '�G���A�̐��������[�v
    For i = LBound(lngIdxNo) To UBound(lngIdxNo)
        With objSelection.Areas(lngIdxNo(i))
            '�s�̐��������[�v
            For j = 1 To .Rows.Count '@
                k = k + 1
                Set objSizeList = New CSizeList                      '@
                Call objSizeList.CreateSizeList(k, frmPage2)
                Call objSizeList.SetValues(.Rows(j).EntireRow) '@
                Call colSizeList(2).Add(objSizeList)
        
                '100���ő�Ƃ���
                If k > C_MAX Then
                    Call Err.Raise(513, , C_MAX & "�ȏ�̍s�ɑ΂��Ď��s�ł��܂���B")
                End If
            Next j
        End With
    Next i
    
    '***************************************
    '�G���A�P�ʂ̏����쐬
    '***************************************
    k = 0
    '�G���A�̐��������[�v
    For Each objRange In Selection.Areas
        k = k + 1
        Set objSizeList = New CSizeList
        Call objSizeList.CreateSizeList(k, frmPage1)
        Call objSizeList.SetValues(objRange.Rows.EntireRow)  '@
        Call colSizeList(1).Add(objSizeList)
    Next objRange
End Sub

'*****************************************************************************
'[�C�x���g]�@frmPage1_Exit
'[ �T  �v ]�@�ۑ������������N���A����
'*****************************************************************************
Private Sub frmPage1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ClearRatio
End Sub

'*****************************************************************************
'[�C�x���g]�@lblSizeInfo_Click
'[ �T  �v ]�@�u@@@mm �� @@@�s�N�Z���v�̒l���J�X�^�}�C�Y������
'*****************************************************************************
Public Sub lblSizeInfo_Click()
    With frmSizeInfo
        '�t�H�[����\��
        Call .Show
        
        '�L�����Z����
        If blnFormLoad = False Then
            Exit Sub
        End If
        
        Call Unload(frmSizeInfo)
    End With
    
    Call SetSizeInfoLabel
    Call ReCalc
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_QueryClose
'[ �T  �v ]�@�~�{�^���Ńt�H�[������鎞�A�ύX�����ɖ߂�
'*****************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    '�ύX���Ȃ���΃t�H�[�������
    If Me.Caption <> "�T�C�Y�ꗗ�i�ύX����j" Then
        Exit Sub
    End If
        
    '�~�{�^���Ńt�H�[������鎞�A�t�H�[������Ȃ�
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Exit Sub
    End If
    
    Call SetOnUndo
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdSave_Click
'[ �T  �v ]�@�m��{�^��������
'*****************************************************************************
Private Sub cmdSave_Click()
    Me.Caption = "�T�C�Y�ꗗ"
    cmdSave.Enabled = False
    cmdUndo.Enabled = False
    '����{�^����L���ɂ���
    Call EnableMenuItem(GetSystemMenu(FindWindow("ThunderDFrame", Me.Caption), False), SC_CLOSE, MF_BYCOMMAND)
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdUndo_Click
'[ �T  �v ]�@���ɖ߂��{�^��������
'*****************************************************************************
Private Sub cmdUndo_Click()
    Call ExecUndo
    Application.ScreenUpdating = True
    Call ReCalc
    
    Call cmdSave_Click
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdOK_Click
'[ �T  �v ]�@�n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@tbsTab_Change
'[ �T  �v ]�@�^�u�ύX��
'*****************************************************************************
Private Sub tbsTab_Change()
    Dim lngNo   As Long '�u�ڍ׏��v���N���b�N����No
    Dim frmPage As MSForms.Frame
    
    Select Case Me.Tag
    Case E_Col
        tbsTab(1).Caption = "1��P�ʂ̏��"
    Case E_Row
        tbsTab(1).Caption = "1�s�P�ʂ̏��"
    End Select
    
    Select Case tbsTab.Value
    Case 0 '�����̃^�u
        frmPage1.Visible = True
        frmPage2.Visible = False
        frmPage3.Visible = False
        Me.Controls("tbsTab").Tag = 0
    Case 1 '�E���̃^�u
        lngNo = Me.Controls("tbsTab").Tag
        If lngNo = 0 Then
            frmPage1.Visible = False
            frmPage2.Visible = True
            frmPage3.Visible = False
        Else '�u�ڍ׏��v���N���b�N���ă^�u���ύX���ꂽ��
            frmPage1.Visible = False
            frmPage2.Visible = False
            frmPage3.Visible = True
        End If
    End Select
        
    '�u�I���v�`�F�b�N�̃N���A
    lngSelNo = 0
    If tbsTab.Value = 1 Then
        If lngNo = 0 Then
            Call ClearSelChk(frmPage2)
        Else
            Call ClearSelChk(frmPage3)
        End If
    End If
    
    '�u�ڍ׏��v���N���b�N���ă^�u���ύX���ꂽ��
    If lngNo <> 0 Then
        tbsTab(1).Caption = tbsTab(1).Caption & "(�ڍ�)"
        Select Case Me.Tag
        Case E_Col
            Call TabChangeCol(lngNo)
        Case E_Row
            Call TabChangeRow(lngNo)
        End Select
    Else
        tbsTab(1).Caption = tbsTab(1).Caption & "(�S��)"
    End If
    
    Call ReCalc
End Sub

'*****************************************************************************
'[ �֐��� ]�@ClearSelChk
'[ �T  �v ]  �u�I���v�`�F�b�N�̃N���A
'[ ��  �� ]�@�N���A����t���[���I�u�W�F�N�g
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ClearSelChk(ByRef frmPage As MSForms.Frame)
On Error GoTo ErrHandle
    Dim i As Long
    For i = 1 To C_MAX
        frmPage.Controls("chkSelect" & i).Value = False
    Next i
ErrHandle:
End Sub
'*****************************************************************************
'[ �֐��� ]�@TabChangeCol
'[ �T  �v ]  �u�ڍ׏��v���N���b�N���ă^�u���ύX���ꂽ��(��ύX��)
'[ ��  �� ]�@�G���A��No
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub TabChangeCol(ByVal lngNo As Long)
    Dim objSizeList As CSizeList
    Dim objRange    As Range
    Dim i           As Long

    Set colSizeList(3) = New Collection
    Set objRange = Selection.Areas(lngNo)
    
    '��̐��������[�v
    For i = 1 To objRange.Columns.Count  '@
        Set objSizeList = New CSizeList
        Call objSizeList.CreateSizeList(i, frmPage3)
        Call objSizeList.SetValues(objRange.Columns(i).EntireColumn) '@
        Call colSizeList(3).Add(objSizeList)

        '100���ő�Ƃ���
        If i = C_MAX Then
            Exit For
        End If
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ]�@TabChangeRow
'[ �T  �v ]  �u�ڍ׏��v���N���b�N���ă^�u���ύX���ꂽ��(�s�ύX��)
'[ ��  �� ]�@�G���A��No
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub TabChangeRow(ByVal lngNo As Long)
    Dim objSizeList As CSizeList
    Dim objRange    As Range
    Dim i           As Long

    Set colSizeList(3) = New Collection
    Set objRange = Selection.Areas(lngNo)
    
    '��̐��������[�v
    For i = 1 To objRange.Rows.Count  '@
        Set objSizeList = New CSizeList
        Call objSizeList.CreateSizeList(i, frmPage3)
        Call objSizeList.SetValues(objRange.Rows(i).EntireRow) '@
        Call colSizeList(3).Add(objSizeList)

        '100���ő�Ƃ���
        If i = C_MAX Then
            Exit For
        End If
    Next i
End Sub

'*****************************************************************************
'[ �֐��� ]�@ReCalc
'[ �T  �v ]  ����v�,����ϣ�ɍŐV�̒l��ݒ肷��
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub ReCalc()
    Dim objSizeList As CSizeList
    Dim lngSumPixel As Long
    Dim i           As Long
    Dim strAvg      As String
    
    If tbsTab.Value = 0 Then
        i = 1
    ElseIf Me.Controls("tbsTab").Tag = 0 Then
        i = 2
    Else
        i = 3
    End If
    
    '�G���A�̐��������[�v
    For Each objSizeList In colSizeList(i)
        Call objSizeList.SetSize
        lngSumPixel = lngSumPixel + objSizeList.Pixel
    Next objSizeList

    '�G���A�̐��������[�v
    For Each objSizeList In colSizeList(i)
        Call objSizeList.SetPercent(lngSumPixel)
    Next objSizeList

    txtSum.Text = CStr(lngSumPixel)
    lblSum.Caption = Format$(PixelToMillimeter(lngSumPixel), "##0.0") & "mm"
    
    If colSizeList(i).Count <> 0 Then
        txtAvg.Text = Format$(CStr(lngSumPixel / colSizeList(i).Count), "0.0")
        strAvg = Format$(PixelToMillimeter(lngSumPixel / colSizeList(i).Count), "##0.0") & "mm"
        If colSizeList(i).Count = 1 Then
            lblAvg.Caption = strAvg
        Else
            lblAvg.Caption = strAvg & Format$(1 / colSizeList(i).Count * 100, " #0.0") & "��"
        End If
    End If

    Call Me.Controls("frmPage" & CStr(i)).SetFocus
End Sub

'*****************************************************************************
'[ �֐��� ]�@PixelToMillimeter
'[ �T  �v ]  �P�ʂ̕ϊ� Pixel �� mm
'[ ��  �� ]�@Pixel
'[ �߂�l ]�@mm
'*****************************************************************************
Public Function PixelToMillimeter(ByVal lngPixel As Long) As Double
    With udtSizeInfo
        PixelToMillimeter = lngPixel * .Millimeters / .Pixel
    End With
End Function

'*****************************************************************************
'[�v���p�e�B]�@SelectNo
'[ �T  �v ]�@�I���Ń`�F�b�N����Ă���A�G���ANo
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get SelectNo() As Long
    SelectNo = lngSelNo
End Property
Public Property Let SelectNo(ByVal Value As Long)
    lngSelNo = Value
End Property

'*****************************************************************************
'[ �֐��� ]�@SaveRatio
'[ �T  �v ]�@�䗦��ۑ�����G���ANo��ۑ�
'[ ��  �� ]�@�G���ANo
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SaveRatio(ByVal lngNo As Long)
    If tbsTab.Value = 0 Then
        lngRatioNo = lngNo
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@CheckRatio
'[ �T  �v ]�@�䗦���ۑ�����Ă��邩�`�F�b�N
'[ ��  �� ]�@�G���ANo
'[ �߂�l ]�@True:�ۑ�����Ă���
'*****************************************************************************
Public Function CheckRatio(ByVal lngNo As Long) As Boolean
    If tbsTab.Value = 0 And lngRatioNo = lngNo Then
        CheckRatio = True
    Else
        CheckRatio = False
    End If
End Function

'*****************************************************************************
'[ �֐��� ]�@ClearRatio
'[ �T  �v ]�@�䗦��ۑ����������N���A����
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub ClearRatio()
    If lngRatioNo <> 0 Then
        Call colSizeList(1).Item(lngRatioNo).ClearRatio
        lngRatioNo = 0
    End If
End Sub

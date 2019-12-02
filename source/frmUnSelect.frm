VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUnSelect 
   Caption         =   "�I�����Ă�������"
   ClientHeight    =   2136
   ClientLeft      =   108
   ClientTop       =   336
   ClientWidth     =   4668
   OleObjectBlob   =   "frmUnSelect.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmUnSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�̈�̎������ʂ̃��[�h
Public Enum EUnselectMode
    E_Unselect  '�����
    E_Reverse   '���]
    E_Union     '�ǉ�
    E_Intersect '�i�荞��
End Enum

Private lngReferenceStyle As Long
Private blnCheck As Boolean

Private strLastSheet   As String '�O��̗̈�̕����p
Private strLastAddress As String '�O��̗̈�̕����p

'*****************************************************************************
'[�C�x���g]�@�e��}�E�X���쎞
'[ �T  �v ]�@RefEdit��L���ɂ�����
'*****************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefEdit.SetFocus
End Sub
Private Sub Frame_Click()
    RefEdit.SetFocus
End Sub
Private Sub Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefEdit.SetFocus
End Sub
Private Sub lblTitle_Click()
    RefEdit.SetFocus
End Sub
Private Sub lblTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefEdit.SetFocus
End Sub
Private Sub cmdOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefEdit.SetFocus
End Sub
Private Sub cmdCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RefEdit.SetFocus
End Sub

'*****************************************************************************
'[�C�x���g]�@chkReverse_Change
'[ �T  �v ]�@���]�`�F�b�N��
'*****************************************************************************
Private Sub chkReverse_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkReverse.Value = True Then
        Call ChangeMode(E_Reverse)
    Else
        Call ChangeMode(E_Unselect)
    End If
End Sub

'*****************************************************************************
'[�C�x���g]�@chkIntersect_Change
'[ �T  �v ]�@�i�荞�݃`�F�b�N��
'*****************************************************************************
Private Sub chkIntersect_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkIntersect.Value = True Then
        Call ChangeMode(E_Intersect)
    Else
        Call ChangeMode(E_Unselect)
    End If
End Sub

'*****************************************************************************
'[�C�x���g]�@chkUnion_Change
'[ �T  �v ]�@�ǉ��`�F�b�N��
'*****************************************************************************
Private Sub chkUnion_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkUnion.Value = True Then
        Call ChangeMode(E_Union)
    Else
        Call ChangeMode(E_Unselect)
    End If
End Sub

'*****************************************************************************
'[ �֐��� ]�@ChangeMode
'[ �T  �v ]  ����]����i�荞�ݣ���������̃��[�h��ύX����
'[ ��  �� ]�@���[�h�^�C�v
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub ChangeMode(ByVal enmModeType As EUnselectMode)
    Select Case enmModeType
    Case E_Unselect
        lblTitle.Caption = "�}�E�X�őI����������������̈��I�����Ă�������"
    Case E_Reverse
        lblTitle.Caption = "�}�E�X�őI���𔽓]���������̈��I�����Ă�������"
    Case E_Intersect
        lblTitle.Caption = "�}�E�X�őI�����i�荞�݂����̈��I�����Ă�������"
    Case E_Union
        lblTitle.Caption = "�}�E�X�őI����ǉ��������̈��I�����Ă�������"
    End Select
    
    blnCheck = True
    Select Case enmModeType
    Case E_Unselect
        chkReverse.Value = False
        chkIntersect.Value = False
        chkUnion.Value = False
    Case E_Reverse
        chkReverse.Value = True
        chkIntersect.Value = False
        chkUnion.Value = False
    Case E_Intersect
        chkReverse.Value = False
        chkIntersect.Value = True
        chkUnion.Value = False
    Case E_Union
        chkReverse.Value = False
        chkIntersect.Value = False
        chkUnion.Value = True
    End Select
    blnCheck = False
    
    RefEdit.SetFocus
End Sub
    
'*****************************************************************************
'[�C�x���g]�@UserForm_Initialize
'[ �T  �v ]�@�t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()

    lngReferenceStyle = Application.ReferenceStyle
    Application.ReferenceStyle = xlA1

    '�I�v�V�������B��
    Me.Height = 128
    
    '�Ăь��ɒʒm����
    blnFormLoad = True
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_Terminate
'[ �T  �v ]�@�t�H�[���A�����[�h��
'*****************************************************************************
Private Sub UserForm_Terminate()
    Application.ReferenceStyle = lngReferenceStyle
    '�Ăь��ɒʒm����
    blnFormLoad = False
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdLastSelect_Click
'[ �T  �v ]�@�O��̗̈�̕����{�^��������
'*****************************************************************************
Private Sub cmdLastSelect_Click()
    Call Range(strLastAddress).Select
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdOK_Click
'[ �T  �v ]�@�n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
    Call cmdOK.SetFocus
    Me.Hide
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdCancel_Click
'[ �T  �v ]�@�L�����Z���{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetLastSelect
'[ �T  �v ]�@���O�̃R�}���h���s���ɑI�����ꂽ�̈�̃A�h���X��ۑ�����
'[ ��  �� ]  ���O�̗̈�̏��
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetLastSelect(ByVal strSheetName As String, ByVal strAddress As String)
    strLastSheet = strSheetName
    strLastAddress = strAddress
    
    If strLastAddress = "" Or ActiveSheet.Name <> strLastSheet Then
        cmdLastSelect.Enabled = False
    End If
End Sub

'*****************************************************************************
'[�v���p�e�B]�@SelectRange
'[ �T  �v ]�@�I�����ꂽ�̈�
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get SelectRange() As Range
    Dim objRange  As Range
    Dim strAddress As String
    
    For Each objRange In Range(RefEdit.Value).Areas
        strAddress = strAddress & Range(GetMergeAddress(objRange.Address)).Address(False, False) & ","
    Next
    
    '�����̃J���}���폜
    Set SelectRange = Range(Left$(strAddress, Len(strAddress) - 1))
End Property

'*****************************************************************************
'[�v���p�e�B]�@Mode
'[ �T  �v ]�@�I�����[�h
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get Mode() As EUnselectMode
    Select Case (True)
    Case chkReverse.Value
        Mode = E_Reverse
    Case chkIntersect.Value
        Mode = E_Intersect
    Case chkUnion.Value
        Mode = E_Union
    Case Else
        Mode = E_Unselect
    End Select
End Property

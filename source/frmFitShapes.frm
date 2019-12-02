VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFitShapes 
   Caption         =   "�I�����Ă�������"
   ClientHeight    =   2040
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3960
   OleObjectBlob   =   "frmFitShapes.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmFitShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum EFitType
    E_Default
    E_TopLeft
    E_Another
End Enum

'*****************************************************************************
'[�C�x���g]�@optXXXXXX_Enter
'[ �T  �v ]�@�t�H�[�J�X�擾��
'*****************************************************************************
Private Sub optDefault_Enter()
    optDefault.Value = True
End Sub
Private Sub optTopLeft_Enter()
    optTopLeft.Value = True
End Sub
Private Sub optAnother_Enter()
    optAnother.Value = True
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_Initialize
'[ �T  �v ]�@�t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    '�Ăь��ɒʒm����
    blnFormLoad = True
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_Terminate
'[ �T  �v ]�@�t�H�[���A�����[�h��
'*****************************************************************************
Private Sub UserForm_Terminate()
    '�Ăь��ɒʒm����
    blnFormLoad = False
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdOK_Click
'[ �T  �v ]�@�n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
    Me.Hide
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdCancel_Click
'[ �T  �v ]�@����{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@optButton_DblClick
'[ �T  �v ]�@�S�̃_�u���N���b�N
'*****************************************************************************
Private Sub optDefault_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub optTopLeft_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub optAnother_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub

'*****************************************************************************
'[�C�x���g]�@lblAnother_Click
'[ �T  �v ]�@�����N���b�N��
'*****************************************************************************
Private Sub lblAnother_Click()
    Call OpenHelpPage("Tips.htm#MoveCell3")
End Sub

'*****************************************************************************
'[�C�x���g]�@lblAnother_MouseMove
'[ �T  �v ]�@�����Ń}�E�X�J�[�\������̌`�ɂ���
'*****************************************************************************
Private Sub lblAnother_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call SetCursor(LoadCursor(0, IDC_HAND))
End Sub

'*****************************************************************************
'[�v���p�e�B]�@SelectType
'[ �T  �v ]�@�I���^�C�v
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get SelectType() As EFitType
    Select Case True
    Case optDefault
        SelectType = E_Default
    Case optTopLeft
        SelectType = E_TopLeft
    Case optAnother
        SelectType = E_Another
    End Select
End Property

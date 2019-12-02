VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAlignment 
   Caption         =   "�I�����Ă�������"
   ClientHeight    =   1812
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   2988
   OleObjectBlob   =   "frmAlignment.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmAlignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum EAlignmentType
    E_Top
    E_Center
    E_Bottom
    E_Cancel
End Enum

'*****************************************************************************
'[�C�x���g]�@optXXXXXX_Enter
'[ �T  �v ]�@�t�H�[�J�X�擾��
'*****************************************************************************
Private Sub optTop_Enter()
    optTop.Value = True
End Sub
Private Sub optCenter_Enter()
    optCenter.Value = True
End Sub
Private Sub optBottom_Enter()
    optBottom.Value = True
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
Private Sub optTop_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub optCenter_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub optBottom_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub

'*****************************************************************************
'[�v���p�e�B]�@SelectType
'[ �T  �v ]�@�I���^�C�v
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get SelectType() As EAlignmentType
    Select Case True
    Case optTop
        SelectType = E_Top
    Case optCenter
        SelectType = E_Center
    Case optBottom
        SelectType = E_Bottom
    End Select
End Property

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEraseSelect 
   Caption         =   "�I�����Ă�������"
   ClientHeight    =   2640
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4764
   OleObjectBlob   =   "frmEraseSelect.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmEraseSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�I�����ꂽ�^�C�v
Public Enum ESelectType
    E_Front
    E_Middle
    E_Back
End Enum

Private m_enmSelectType As ESelectType

'*****************************************************************************
'[�C�x���g]�@optButton_Enter
'[ �T  �v ]�@�I�v�V�����{�^���t�H�[�J�X�擾��
'*****************************************************************************
Private Sub optFront_Enter()
    optFront.Value = True
End Sub
Private Sub optMiddle_Enter()
    optMiddle.Value = True
End Sub
Private Sub optBack_Enter()
    optBack.Value = True
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
'[�C�x���g]�@optButton_DblClick
'[ �T  �v ]�@�I�v�V�����{�^���_�u���N���b�N
'*****************************************************************************
Private Sub optFront_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub optMiddle_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub optBack_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
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
'[�v���p�e�B]�@SelectType
'[ �T  �v ]�@�I���^�C�v
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get SelectType() As ESelectType
    Select Case True
    Case optFront
        SelectType = E_Front
    Case optMiddle
        SelectType = E_Middle
    Case optBack
        SelectType = E_Back
    End Select
End Property

'*****************************************************************************
'[�v���p�e�B]�@Hidden
'[ �T  �v ]�@��\���Ƀ`�F�b�N�������ǂ���
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get Hidden() As Boolean
    Hidden = chkHidden
End Property

'*****************************************************************************
'[�v���p�e�B]�@TopSelect
'[ �T  �v ]�@�擪��(�s)��I�����Ă��邩�ǂ���
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Let TopSelect(ByVal Value As Boolean)
    If Value = True Then
        optMiddle.Enabled = False
        optBack.Enabled = False
        optFront.Value = True
        Call optFront.SetFocus
    Else
        optBack.Value = True
        Call optBack.SetFocus
    End If
End Property

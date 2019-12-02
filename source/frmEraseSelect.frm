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
Enum ESelectType
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
'[�C�x���g]�@UserForm_Activate
'[ �T  �v ]�@�t�H�[�����A�N�e�B�u�ɂȂ鎞
'*****************************************************************************
Private Sub UserForm_Activate()
    
    '�f�t�H���g�I��l��ݒ肵�A�t�H�[�J�X�����Ă�
    Select Case m_enmSelectType
    Case E_Front
        optFront.Value = True
        optFront.SetFocus
    Case E_Middle
        optMiddle.Value = True
        optMiddle.SetFocus
    Case E_Back
        optBack.Value = True
        optBack.SetFocus
    End Select
End Sub

'*****************************************************************************
'[ �֐��� ]�@SetEnabled
'[ �T  �v ]�@�I���o���Ȃ��񂹕���I���o���Ȃ�����
'[ ��  �� ]�@enmSelectType : �I���o���Ȃ��񂹕�
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Public Sub SetEnabled(ByVal enmSelectType As ESelectType)
    Select Case enmSelectType
    Case E_Front
        optFront.Enabled = False
    Case E_Middle
        optMiddle.Enabled = False
    Case E_Back
        optBack.Enabled = False
    End Select
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
Public Property Let SelectType(ByVal Value As ESelectType)
    m_enmSelectType = Value
End Property

'*****************************************************************************
'[�v���p�e�B]�@Hidden
'[ �T  �v ]�@��\���Ƀ`�F�b�N�������ǂ���
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get Hidden() As Boolean
    Hidden = chkHidden
End Property

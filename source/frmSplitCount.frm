VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplitCount 
   Caption         =   "�I�����Ă�������"
   ClientHeight    =   1608
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4248
   OleObjectBlob   =   "frmSplitCount.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSplitCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum ESizeType
    E_OrgSize
    E_SameSiZe
    E_NonSize
End Enum

'*****************************************************************************
'[�C�x���g]�@UserForm_Initialize
'[ �T  �v ]�@�t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    '�Ăь��ɒʒm����
    blnFormLoad = True
    
    If CommandBars.ActionControl.TooltipText = "�s�𕪊�" Then
        chkInsert.Caption = "���̍����Ɠ��������̍s��}������"
    Else
        chkInsert.Caption = "���̕��Ɠ������̗��}������"
    End If
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
'[�C�x���g]�@txtCount_KeyDown
'[ �T  �v ]�@�����C�����L�[�ŃT�C�Y�̕ύX���o����悤�ɂ���
'            �����C�����L�[�ŏ㉺�Ɉړ�����悤�ɂ���
'*****************************************************************************
Private Sub txtCount_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case (KeyCode)
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown
        Call SpbCount.SetFocus
    Case Else
        If Shift = 0 Then
            Call MsgBox("�J�[�\���L�[�ő��삵�ĉ�����")
            Call SpbCount.SetFocus
        End If
    End Select
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
'[ �T  �v ]�@�L�����Z���{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@SpbCount_Change
'[ �T  �v ]�@�������ύX��
'*****************************************************************************
Private Sub SpbCount_Change()
    txtCount.Text = CStr(SpbCount.Value)
End Sub

'*****************************************************************************
'[�v���p�e�B]�@Count
'[ �T  �v ]�@������
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get Count() As Long
    Count = SpbCount.Value
End Property

'*****************************************************************************
'[�v���p�e�B]�@CheckInsert
'[ �T  �v ]�@���̕�(����)�Ɠ�����(����)�̗�(�s)��}�����邩�ǂ���
'[ ��  �� ]�@�Ȃ�
'*****************************************************************************
Public Property Get CheckInsert() As Boolean
    CheckInsert = chkInsert.Value
End Property

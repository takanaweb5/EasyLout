VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOption 
   Caption         =   "�I�v�V����"
   ClientHeight    =   4236
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4980
   OleObjectBlob   =   "frmOption.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'*****************************************************************************
'[�C�x���g]�@UserForm_Initialize
'[ �T  �v ]�@�t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    chkCheck1 = GetSetting(REGKEY, "KEY", "OpenEdit", True)
    chkCheck2 = GetSetting(REGKEY, "KEY", "CopyText", True)
    chkCheck3 = GetSetting(REGKEY, "KEY", "PasteText", True)
    chkCheck4 = GetSetting(REGKEY, "KEY", "BackSpace", True)
    chkCheck5 = GetSetting(REGKEY, "KEY", "FindNext", True)
    chkCheck6 = GetSetting(REGKEY, "KEY", "FindPrev", True)
    txtFont.Text = GetSetting(REGKEY, "KEY", "FontName", DEFAULTFONT)
    txtDelimiter.Text = GetSetting(REGKEY, "KEY", "Delimiter", " ")
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdOK_Click
'[ �T  �v ]�@�n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
    Call SaveSetting(REGKEY, "KEY", "OpenEdit", chkCheck1)
    Call SaveSetting(REGKEY, "KEY", "CopyText", chkCheck2)
    Call SaveSetting(REGKEY, "KEY", "PasteText", chkCheck3)
    Call SaveSetting(REGKEY, "KEY", "BackSpace", chkCheck4)
    Call SaveSetting(REGKEY, "KEY", "FindNext", chkCheck5)
    Call SaveSetting(REGKEY, "KEY", "FindPrev", chkCheck6)
    Call SaveSetting(REGKEY, "KEY", "FontName", txtFont.Text)
    Call SaveSetting(REGKEY, "KEY", "Delimiter", txtDelimiter.Text)
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdCancel_Click
'[ �T  �v ]�@�L�����Z���{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOption 
   Caption         =   "�I�v�V����"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
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
    Dim strOption As String
    strOption = CommandBars("���񂽂񃌃C�A�E�g").Controls(1).Tag
    chkCheck1 = Not (InStr(1, strOption, "{S+F2}") = 0)
    chkCheck2 = Not (InStr(1, strOption, "{C+S+C}") = 0)
    chkCheck3 = Not (InStr(1, strOption, "{C+S+V}") = 0)
    chkCheck4 = Not (InStr(1, strOption, "{BS}") = 0)
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdOK_Click
'[ �T  �v ]�@�n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
    Dim strOption As String
    
    If chkCheck1 = True Then
        strOption = strOption & "{S+F2}"
    End If
    If chkCheck2 = True Then
        strOption = strOption & "{C+S+C}"
    End If
    If chkCheck3 = True Then
        strOption = strOption & "{C+S+V}"
    End If
    If chkCheck4 = True Then
        strOption = strOption & "{BS}"
    End If
    CommandBars("���񂽂񃌃C�A�E�g").Controls(1).Tag = strOption
    
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdCancel_Click
'[ �T  �v ]�@�L�����Z���{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdHelp_Click
'[ �T  �v ]�@�w���v�{�^��������
'*****************************************************************************
Private Sub cmdHelp_Click()
    Call OpenHelpPage("Tips.htm#Option")
End Sub


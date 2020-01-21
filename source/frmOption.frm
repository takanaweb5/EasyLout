VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOption 
   Caption         =   "オプション"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   OleObjectBlob   =   "frmOption.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'*****************************************************************************
'[イベント]　UserForm_Initialize
'[ 概  要 ]　フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()
    Dim strOption As String
    strOption = CommandBars("かんたんレイアウト").Controls(1).Tag
    chkCheck1 = Not (InStr(1, strOption, "{S+F2}") = 0)
    chkCheck2 = Not (InStr(1, strOption, "{C+S+C}") = 0)
    chkCheck3 = Not (InStr(1, strOption, "{C+S+V}") = 0)
    chkCheck4 = Not (InStr(1, strOption, "{BS}") = 0)
End Sub

'*****************************************************************************
'[イベント]　cmdOK_Click
'[ 概  要 ]　ＯＫボタン押下時
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
    CommandBars("かんたんレイアウト").Controls(1).Tag = strOption
    
    Call Unload(Me)
End Sub

'*****************************************************************************
'[イベント]　cmdCancel_Click
'[ 概  要 ]　キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[イベント]　cmdHelp_Click
'[ 概  要 ]　ヘルプボタン押下時
'*****************************************************************************
Private Sub cmdHelp_Click()
    Call OpenHelpPage("Tips.htm#Option")
End Sub


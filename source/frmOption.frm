VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOption 
   Caption         =   "オプション"
   ClientHeight    =   2160
   ClientLeft      =   48
   ClientTop       =   336
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
    chkCheck1 = GetSetting(REGKEY, "KEY", "OpenEdit", True)
    chkCheck2 = GetSetting(REGKEY, "KEY", "CopyText", True)
    chkCheck3 = GetSetting(REGKEY, "KEY", "PasteText", True)
    chkCheck4 = GetSetting(REGKEY, "KEY", "BackSpace", True)
End Sub

'*****************************************************************************
'[イベント]　cmdOK_Click
'[ 概  要 ]　ＯＫボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
    Call SaveSetting(REGKEY, "KEY", "OpenEdit", chkCheck1)
    Call SaveSetting(REGKEY, "KEY", "CopyText", chkCheck2)
    Call SaveSetting(REGKEY, "KEY", "PasteText", chkCheck3)
    Call SaveSetting(REGKEY, "KEY", "BackSpace", chkCheck4)
    Call Unload(Me)
End Sub

'*****************************************************************************
'[イベント]　cmdCancel_Click
'[ 概  要 ]　キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub


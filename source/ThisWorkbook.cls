VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'*****************************************************************************
'[イベント]　AddinInstall
'[ 概  要 ]　アドインのインストール時に実行される
'*****************************************************************************
Private Sub Workbook_AddinInstall()
    On Error Resume Next
End Sub

'*****************************************************************************
'[イベント]　AddinUninstall
'[ 概  要 ]　アドインのアンインストール時に実行される
'*****************************************************************************
Private Sub Workbook_AddinUninstall()
    On Error Resume Next
    Call DeleteSetting(REGKEY)
End Sub

'*****************************************************************************
'[イベント]　Workbook_Open
'[ 概  要 ]　開く時
'*****************************************************************************
Private Sub Workbook_Open()
    On Error Resume Next
'    Call SetDPIRatio
    Call SetKeys
End Sub

'*****************************************************************************
'[イベント]　Workbook_BeforeClose
'[ 概  要 ]　閉じる時
'*****************************************************************************
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
    On Error Resume Next
    Call Application.OnKey("+{F2}")
    Call Application.OnKey("+^{c}")
    Call Application.OnKey("+^{v}")
    Call Application.OnKey("{BS}")
    Call Application.OnKey("^6")
End Sub


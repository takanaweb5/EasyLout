VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplitCount 
   Caption         =   "選択してください"
   ClientHeight    =   1608
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4248
   OleObjectBlob   =   "frmSplitCount.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSplitCount"
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
    '呼び元に通知する
    blnFormLoad = True
End Sub

'*****************************************************************************
'[イベント]　UserForm_Terminate
'[ 概  要 ]　フォームアンロード時
'*****************************************************************************
Private Sub UserForm_Terminate()
    '呼び元に通知する
    blnFormLoad = False
End Sub

'*****************************************************************************
'[イベント]　txtCount_KeyDown
'[ 概  要 ]　｢→｣，｢←｣キーでサイズの変更を出来るようにする
'            ｢↓｣，｢↑｣キーで上下に移動するようにする
'*****************************************************************************
Private Sub txtCount_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case (KeyCode)
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown
        Call SpbCount.SetFocus
    Case Else
        If Shift = 0 Then
            Call MsgBox("カーソルキーで操作して下さい")
            Call SpbCount.SetFocus
        End If
    End Select
End Sub

'*****************************************************************************
'[概要] ラベルを編集する
'[引数] True:列選択時、False:行選択時
'*****************************************************************************
Public Sub SetChkLabel(ByVal blnColumn As Boolean)
    If blnColumn Then
        chkInsert.Caption = "元の幅と同じ幅の列を挿入する"
    Else
        chkInsert.Caption = "元の高さと同じ高さの行を挿入する"
    End If
End Sub

'*****************************************************************************
'[イベント]　cmdOK_Click
'[ 概  要 ]　ＯＫボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
    Me.Hide
End Sub

'*****************************************************************************
'[イベント]　cmdCancel_Click
'[ 概  要 ]　キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[イベント]　SpbCount_Change
'[ 概  要 ]　分割数変更時
'*****************************************************************************
Private Sub SpbCount_Change()
    txtCount.Text = CStr(SpbCount.Value)
End Sub

'*****************************************************************************
'[プロパティ]　Count
'[ 概  要 ]　分割数
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get Count() As Long
    Count = SpbCount.Value
End Property

'*****************************************************************************
'[プロパティ]　CheckInsert
'[ 概  要 ]　元の幅(高さ)と同じ幅(高さ)の列(行)を挿入するかどうか
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get CheckInsert() As Boolean
    CheckInsert = chkInsert.Value
End Property

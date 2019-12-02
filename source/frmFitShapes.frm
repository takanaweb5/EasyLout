VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFitShapes 
   Caption         =   "選択してください"
   ClientHeight    =   2040
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3960
   OleObjectBlob   =   "frmFitShapes.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFitShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum EFitType
    E_Default
    E_TopLeft
    E_Another
End Enum

'*****************************************************************************
'[イベント]　optXXXXXX_Enter
'[ 概  要 ]　フォーカス取得時
'*****************************************************************************
Private Sub optDefault_Enter()
    optDefault.Value = True
End Sub
Private Sub optTopLeft_Enter()
    optTopLeft.Value = True
End Sub
Private Sub optAnother_Enter()
    optAnother.Value = True
End Sub

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
'[イベント]　cmdOK_Click
'[ 概  要 ]　ＯＫボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
    Me.Hide
End Sub

'*****************************************************************************
'[イベント]　cmdCancel_Click
'[ 概  要 ]　取消ボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[イベント]　optButton_DblClick
'[ 概  要 ]　全体ダブルクリック
'*****************************************************************************
Private Sub optDefault_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub optTopLeft_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub optAnother_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub

'*****************************************************************************
'[イベント]　lblAnother_Click
'[ 概  要 ]　説明クリック時
'*****************************************************************************
Private Sub lblAnother_Click()
    Call OpenHelpPage("Tips.htm#MoveCell3")
End Sub

'*****************************************************************************
'[イベント]　lblAnother_MouseMove
'[ 概  要 ]　説明でマウスカーソルを手の形にする
'*****************************************************************************
Private Sub lblAnother_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call SetCursor(LoadCursor(0, IDC_HAND))
End Sub

'*****************************************************************************
'[プロパティ]　SelectType
'[ 概  要 ]　選択タイプ
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get SelectType() As EFitType
    Select Case True
    Case optDefault
        SelectType = E_Default
    Case optTopLeft
        SelectType = E_TopLeft
    Case optAnother
        SelectType = E_Another
    End Select
End Property

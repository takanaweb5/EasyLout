VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAlignment 
   Caption         =   "選択してください"
   ClientHeight    =   1812
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   2988
   OleObjectBlob   =   "frmAlignment.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmAlignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum EAlignmentType
    E_Top
    E_Center
    E_Bottom
    E_Cancel
End Enum

'*****************************************************************************
'[イベント]　optXXXXXX_Enter
'[ 概  要 ]　フォーカス取得時
'*****************************************************************************
Private Sub optTop_Enter()
    optTop.Value = True
End Sub
Private Sub optCenter_Enter()
    optCenter.Value = True
End Sub
Private Sub optBottom_Enter()
    optBottom.Value = True
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
Private Sub optTop_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub optCenter_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub
Private Sub optBottom_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
End Sub

'*****************************************************************************
'[プロパティ]　SelectType
'[ 概  要 ]　選択タイプ
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get SelectType() As EAlignmentType
    Select Case True
    Case optTop
        SelectType = E_Top
    Case optCenter
        SelectType = E_Center
    Case optBottom
        SelectType = E_Bottom
    End Select
End Property

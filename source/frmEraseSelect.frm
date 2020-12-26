VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEraseSelect 
   Caption         =   "選択してください"
   ClientHeight    =   2640
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4764
   OleObjectBlob   =   "frmEraseSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmEraseSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'選択されたタイプ
Public Enum ESelectType
    E_Front
    E_Middle
    E_Back
End Enum

Private m_enmSelectType As ESelectType

'*****************************************************************************
'[イベント]　optButton_Enter
'[ 概  要 ]　オプションボタンフォーカス取得時
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
'[イベント]　optButton_DblClick
'[ 概  要 ]　オプションボタンダブルクリック
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
'[プロパティ]　SelectType
'[ 概  要 ]　選択タイプ
'[ 引  数 ]　なし
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

'*****************************************************************************
'[プロパティ]　Hidden
'[ 概  要 ]　非表示にチェックしたかどうか
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get Hidden() As Boolean
    Hidden = chkHidden
End Property

'*****************************************************************************
'[プロパティ]　TopSelect
'[ 概  要 ]　先頭列(行)を選択しているかどうか
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Let TopSelect(ByVal Value As Boolean)
    If Value = True Then
        optMiddle.Enabled = False
        optBack.Enabled = False
        optFront.Value = True
        Call optFront.SetFocus
    Else
        optBack.Value = True
        Call optBack.SetFocus
    End If
End Property

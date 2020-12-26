VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSizeInfo 
   Caption         =   "サイズ情報"
   ClientHeight    =   1800
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3924
   OleObjectBlob   =   "frmSizeInfo.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSizeInfo"
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
    Dim lngArray(1 To 4) As Long
    
    lngArray(1) = GetSetting(REGKEY, "SizeInfo", "Width_mm", 0)
    lngArray(2) = GetSetting(REGKEY, "SizeInfo", "Width_Pixel", 0)
    lngArray(3) = GetSetting(REGKEY, "SizeInfo", "Height_mm", 0)
    lngArray(4) = GetSetting(REGKEY, "SizeInfo", "Height_Pixel", 0)
    
    If lngArray(1) = 0 Or lngArray(2) = 0 Or lngArray(3) = 0 Or lngArray(4) = 0 Then
        txtmm1.Value = 100
        txtPixel1.Value = Round(Application.CentimetersToPoints(10) / DPIRatio)
        txtmm2.Value = 100
        txtPixel2.Value = Round(Application.CentimetersToPoints(10) / DPIRatio)
    Else
        txtmm1.Value = lngArray(1)
        txtPixel1.Value = lngArray(2)
        txtmm2.Value = lngArray(3)
        txtPixel2.Value = lngArray(4)
    End If
    
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
    
    'サイズ情報の保存
    Call SaveSetting(REGKEY, "SizeInfo", "Width_mm", CLng(txtmm1.Value))
    Call SaveSetting(REGKEY, "SizeInfo", "Width_Pixel", CLng(txtPixel1.Value))
    Call SaveSetting(REGKEY, "SizeInfo", "Height_mm", CLng(txtmm2.Value))
    Call SaveSetting(REGKEY, "SizeInfo", "Height_Pixel", CLng(txtPixel2.Value))
End Sub

'*****************************************************************************
'[イベント]　cmdCancel_Click
'[ 概  要 ]　キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[イベント]　txtmm_Change
'[ 概  要 ]　値を入力した時、ＯＫボタンのEnabledを制御する
'*****************************************************************************
Private Sub txtmm1_Change()
    Call ChkInput
End Sub
Private Sub txtmm2_Change()
    Call ChkInput
End Sub

'*****************************************************************************
'[イベント]　txtPixel_Change
'[ 概  要 ]　値を入力した時、ＯＫボタンのEnabledを制御する
'*****************************************************************************
Private Sub txtPixel1_Change()
    Call ChkInput
End Sub
Private Sub txtPixel2_Change()
    Call ChkInput
End Sub

'*****************************************************************************
'[ 関数名 ]　ChkInput
'[ 概  要 ]  [mm][ピクセル]ともに数値が入力された時、OKボタンを使用可能にする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ChkInput()
    If IsNumeric(txtmm1.Value) And IsNumeric(txtPixel1.Value) And _
       IsNumeric(txtmm2.Value) And IsNumeric(txtPixel2.Value) Then
        If CInt(txtmm1.Value) > 0 And CInt(txtPixel1.Value) > 0 And _
           CInt(txtmm2.Value) > 0 And CInt(txtPixel2.Value) > 0 Then
            cmdOK.Enabled = True
            Exit Sub
        End If
    End If
    cmdOK.Enabled = False
End Sub

'*****************************************************************************
'[イベント]　txtmm_KeyDown
'[ 概  要 ]　数値のみ入力可能にする
'*****************************************************************************
Private Sub txtmm1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call txt_KeyDown(KeyCode, Shift)
End Sub
Private Sub txtmm2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call txt_KeyDown(KeyCode, Shift)
End Sub

'*****************************************************************************
'[イベント]　txtPixel_KeyDown
'[ 概  要 ]　数値のみ入力可能にする
'*****************************************************************************
Private Sub txtPixel1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call txt_KeyDown(KeyCode, Shift)
End Sub
Private Sub txtPixel2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call txt_KeyDown(KeyCode, Shift)
End Sub

'*****************************************************************************
'[ 関数名 ]　txt_KeyDown
'[ 概  要 ]　数値のみ入力可能にする
'[ 引  数 ]　KeyDownイベントと同じ
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case (KeyCode)
    Case vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
    Case vbKeyLeft, vbKeyRight, vbKeyDelete, vbKeyBack
    Case vbKeyReturn, vbKeyEscape, vbKeyTab
    Case Else
        KeyCode = 0
    End Select
End Sub

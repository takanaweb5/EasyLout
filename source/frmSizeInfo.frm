VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSizeInfo 
   Caption         =   "�T�C�Y���"
   ClientHeight    =   1800
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3924
   OleObjectBlob   =   "frmSizeInfo.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSizeInfo"
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
    Dim strArray() As String
    
    strArray = Split(CommandBars.ActionControl.Tag, ",")
    If UBound(strArray) = 3 Then
        txtmm1.Value = strArray(0)
        txtPixel1.Value = strArray(1)
        txtmm2.Value = strArray(2)
        txtPixel2.Value = strArray(3)
    Else
        txtmm1.Value = 100
        txtPixel1.Value = Round(Application.CentimetersToPoints(10) / DPIRatio)
        txtmm2.Value = 100
        txtPixel2.Value = Round(Application.CentimetersToPoints(10) / DPIRatio)
    End If
    
    '�Ăь��ɒʒm����
    blnFormLoad = True
End Sub

'*****************************************************************************
'[�C�x���g]�@UserForm_Terminate
'[ �T  �v ]�@�t�H�[���A�����[�h��
'*****************************************************************************
Private Sub UserForm_Terminate()
    '�Ăь��ɒʒm����
    blnFormLoad = False
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdOK_Click
'[ �T  �v ]�@�n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
    Me.Hide
    
    '�T�C�Y���̕ۑ�
    CommandBars.ActionControl.Tag = CLng(txtmm1) & "," & CLng(txtPixel1) & "," & _
                                    CLng(txtmm2) & "," & CLng(txtPixel2)
End Sub

'*****************************************************************************
'[�C�x���g]�@cmdCancel_Click
'[ �T  �v ]�@�L�����Z���{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@txtmm_Change
'[ �T  �v ]�@�l����͂������A�n�j�{�^����Enabled�𐧌䂷��
'*****************************************************************************
Private Sub txtmm1_Change()
    Call ChkInput
End Sub
Private Sub txtmm2_Change()
    Call ChkInput
End Sub

'*****************************************************************************
'[�C�x���g]�@txtPixel_Change
'[ �T  �v ]�@�l����͂������A�n�j�{�^����Enabled�𐧌䂷��
'*****************************************************************************
Private Sub txtPixel1_Change()
    Call ChkInput
End Sub
Private Sub txtPixel2_Change()
    Call ChkInput
End Sub

'*****************************************************************************
'[ �֐��� ]�@ChkInput
'[ �T  �v ]  [mm][�s�N�Z��]�Ƃ��ɐ��l�����͂��ꂽ���AOK�{�^�����g�p�\�ɂ���
'[ ��  �� ]�@�Ȃ�
'[ �߂�l ]�@�Ȃ�
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
'[�C�x���g]�@txtmm_KeyDown
'[ �T  �v ]�@���l�̂ݓ��͉\�ɂ���
'*****************************************************************************
Private Sub txtmm1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call txt_KeyDown(KeyCode, Shift)
End Sub
Private Sub txtmm2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call txt_KeyDown(KeyCode, Shift)
End Sub

'*****************************************************************************
'[�C�x���g]�@txtPixel_KeyDown
'[ �T  �v ]�@���l�̂ݓ��͉\�ɂ���
'*****************************************************************************
Private Sub txtPixel1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call txt_KeyDown(KeyCode, Shift)
End Sub
Private Sub txtPixel2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call txt_KeyDown(KeyCode, Shift)
End Sub

'*****************************************************************************
'[ �֐��� ]�@txt_KeyDown
'[ �T  �v ]�@���l�̂ݓ��͉\�ɂ���
'[ ��  �� ]�@KeyDown�C�x���g�Ɠ���
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case (KeyCode)
    Case vbKey0 To vbKey9
    Case vbKeyLeft, vbKeyRight, vbKeyDelete, vbKeyBack
    Case vbKeyReturn, vbKeyEscape, vbKeyTab
    Case Else
        KeyCode = 0
    End Select
End Sub

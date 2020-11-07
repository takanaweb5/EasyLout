Attribute VB_Name = "Module1"
Option Explicit

Sub ClearSheets()
Attribute ClearSheets.VB_Description = "�}�N���L�^�� : 2012/7/4  ���[�U�[�� : naoki-takahashi"
Attribute ClearSheets.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim objSheet As Worksheet
    Dim objActiveSheet As Worksheet
    
    Set objActiveSheet = ActiveSheet
    
'    Application.ScreenUpdating = False
    For Each objSheet In Worksheets
        Call ClearSheet(objSheet)
    Next
    
'    Application.ScreenUpdating = True
'    For Each objSheet In Worksheets
'        'A1�̈ʒu��\��������
'        Call objSheet.Activate
'        Call objSheet.Range("A1").Select
'    Next
    
    '�J���[�p���b�g�̏�����
    Call ActiveWorkbook.ResetColors
    
    '�X�^�C���̍폜
    Call DeleteStyles(ActiveWorkbook)
    
    '���O�I�u�W�F�N�g���폜����
    Call DeleteNames(ActiveWorkbook)
    
    Call objActiveSheet.Select
End Sub

'*****************************************************************************
'[ �֐��� ]�@DeleteNames
'[ �T  �v ]�@���O�I�u�W�F�N�g���폜����
'[ ��  �� ]�@Workbook
'[ �߂�l ]�@�Ȃ�
'*****************************************************************************
Private Sub DeleteNames(ByRef objWorkbook As Workbook)
    Dim objName     As Name
    For Each objName In objWorkbook.Names
        If (Right$(objName.Name, Len("Print_Area")) <> "Print_Area") And _
           (Right$(objName.Name, Len("Print_Titles")) <> "Print_Titles") And _
           (Right$(objName.Name, Len("Database")) <> "Database") Then
'            Debug.Print objName.Name
            Call objName.Delete
        End If
    Next objName
End Sub

Private Sub ClearSheet(ByRef objSheet As Worksheet)
    Call objSheet.Activate
    
'    Call ActiveWindow.ScrollIntoView(0, 0, 1, 1)
    Call ActiveWindow.SmallScroll(, Rows.Count, , Columns.Count)
    Call objSheet.Range("A1").Select
    
    '�g���̕\��
'    ActiveWindow.DisplayGridlines = False
    
    '����������
    If ActiveWindow.Split = True And ActiveWindow.FreezePanes = False Then
        ActiveWindow.Split = False
    End If
    '���y�[�W�v���r���[����
    ActiveWindow.View = xlNormalView
    
    '���y�[�W�\��������
    objSheet.DisplayAutomaticPageBreaks = False
    
    '�{��
    ActiveWindow.Zoom = 100
End Sub

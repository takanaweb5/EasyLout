Attribute VB_Name = "RegExp"
Option Explicit

'*****************************************************************************
'[�T�v] ���K�\���e�X�g
'[����] ������C���K�\���p�^�[��
'[�ߒl] True:��v����
'*****************************************************************************
Public Function RegExpTest(ByVal ������ As String, ByVal �p�^�[�� As String) As Boolean
Attribute RegExpTest.VB_Description = "���K�\���̏����𖞂����ꍇTrue��Ԃ��܂�"
    With CreateObject("VBScript.RegExp")
        .Pattern = �p�^�[��
        RegExpTest = .Test(������)
    End With
End Function

'*****************************************************************************
'[�T�v] ���K�\���u��
'[����] ������C���K�\���p�^�[���C�u��������
'[�ߒl] �u����̕�����
'*****************************************************************************
Public Function RegExpReplace(ByVal ������ As String, ByVal �p�^�[�� As String, ByVal �u�������� As String) As String
Attribute RegExpReplace.VB_Description = "���K�\���̒u�����s���܂�"
    With CreateObject("VBScript.RegExp")
        .Pattern = �p�^�[��
        .Global = True
        RegExpReplace = .Replace(������, �u��������)
    End With
End Function

'*****************************************************************************
'[�T�v] ���K�\���}�b�`������擾
'[����] ������C���K�\���p�^�[���C���Ԗڂ̃T�u�}�b�`���擾���邩
'[�ߒl] �}�b�`������
'*****************************************************************************
Public Function RegExpMatch(ByVal ������ As String, ByVal �p�^�[�� As String, Optional ByVal �T�u�}�b�`�ԍ� As Long = 0) As Variant
Attribute RegExpMatch.VB_Description = "���K�\���Ɉ�v����������Ԃ��܂�\n��3����:�T�u�}�b�`�������擾���������̂݁A����Index���w�肵�܂�"
    With CreateObject("VBScript.RegExp")
        .Pattern = �p�^�[��
        If .Test(������) Then
            Dim objMatches
            Set objMatches = .Execute(������)
            If �T�u�}�b�`�ԍ� > 0 Then
                RegExpMatch = objMatches(0).SubMatches(�T�u�}�b�`�ԍ� - 1)
            Else
                RegExpMatch = objMatches(0).Value
            End If
        Else
            '�}�b�`���Ȃ��ꍇ #N/A�G���[��Ԃ�
            RegExpMatch = CVErr(xlErrNA)
        End If
    End With
End Function

'*****************************************************************************
'[�T�v] ���K�\���}�b�`������擾
'[����] ������C���K�\���p�^�[���C�z��̕���(0=�s����:1=�����)
'[�ߒl] �}�b�`������̔z��
'*****************************************************************************
Public Function RegExpMatches(ByVal ������ As String, ByVal �p�^�[�� As String, Optional ByVal ���� As Long = 0) As Variant
Attribute RegExpMatches.VB_Description = "���K�\���Ɉ�v�����ӏ����ׂĂ�z��ŕԂ��܂�\n�z�񐔎��`��(Ctrl+Shift+Enter)�Ŏ��o���Ă�������\n��3����:�z��̌������w�肵�܂� ��0=�s����(�ȗ���) or 1=�����"
    With CreateObject("VBScript.RegExp")
        .Pattern = �p�^�[��
        .Global = True
        If .Test(������) Then
            Dim objMatches
            Set objMatches = .Execute(������)
            ReDim Result(0 To objMatches.Count - 1)
            Dim i As Long
            For i = 0 To objMatches.Count - 1
                Result(i) = objMatches(i).Value
            Next
            If ���� = 0 Then
                RegExpMatches = WorksheetFunction.Transpose(Result)
            Else
                RegExpMatches = Result
            End If
        Else
            '�}�b�`���Ȃ��ꍇ #N/A�G���[��Ԃ�
            RegExpMatches = CVErr(xlErrNA)
        End If
    End With
End Function


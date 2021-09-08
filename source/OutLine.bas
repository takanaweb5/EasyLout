Attribute VB_Name = "OutLine"
Option Explicit

'*****************************************************************************
'[�T�v] ���O�̒i���ԍ��̎��̒i���ԍ����擾����
'[����] �Ȃ�
'[�ߒl] ��F�i�A�j���i�C�j
'*****************************************************************************
Public Function OutLineNext() As Variant
    Dim i As Long
    Dim Value As Variant
    Application.Volatile '���ꂪ�Ȃ��ƍČv�Z����Ȃ�

    '������1�s���k��A�l�̐ݒ肳�ꂽ�Z��������
    For i = Application.ThisCell.Row - 1 To 1 Step -1
        Value = Application.ThisCell.EntireColumn.Rows(i).Value
        If Value <> "" Then
            '���O�̒i���ԍ����玟�̒i���ԍ����擾
            If VarType(Value) = vbDouble Then
                OutLineNext = CDbl(GetNext(Value))
            Else
                OutLineNext = GetNext(Value)
            End If
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[�T�v] �i���ԍ��̎��̒i���ԍ����擾����
'[����] ���O�̒i���ԍ��@��F�i�A�j
'[�ߒl] ��F�i�A�j���i�C�j
'*****************************************************************************
Private Function GetNext(ByVal strOutLine As String) As String
    '���[�̕���
    Dim strL As String
    strL = Left(strOutLine, 1)
    If InStr(1, "�i(", strL) = 0 Then
        strL = ""
    End If

    '�E�[�̕���
    Dim strR As String
    strR = Right(strOutLine, 1)
    If InStr(1, ".)�D�j", strR) = 0 Then
        strR = ""
    End If

    '���[�̕������폜
    Dim strNum As String
    strNum = Mid(strOutLine, Len(strL) + 1, Len(strOutLine) - Len(strL & strR))

    '�����ȊO�̎���1�����łȂ���
    If Not IsNumeric(strNum) And Len(strNum) > 1 Then
        GetNext = GetNum(strOutLine)
        Exit Function
    End If
    If InStr(1, strNum, "�|") > 0 Or InStr(1, strNum, "�C") > 0 Or InStr(1, strNum, "�D") > 0 Or InStr(1, strNum, "�@") > 0 Or _
       InStr(1, strNum, "-") > 0 Or InStr(1, strNum, ",") > 0 Or InStr(1, strNum, ".") > 0 Or InStr(1, strNum, " ") > 0 Then
        GetNext = GetNum(strOutLine)
        Exit Function
    End If

    '�S�p�� ���� �� �J�i �̎��� �C �̎��� �B�A�J �̎��� �K �ƂȂ邽��
    '���p�łŎ��̕������擾���đS�p�ɖ߂�
    Dim blnHiragana As Boolean
    Dim blnWide As Boolean

    '�S�p�Ђ炪�Ȃ̎��A�J�^�J�i�ɕϊ�
    If StrConv(strNum, vbKatakana) <> strNum Then
        blnHiragana = True
        strNum = StrConv(strNum, vbKatakana)
    End If

    '�S�p�̐����E�J�^�J�i�̎��͔��p�ɕϊ�
    If StrConv(strNum, vbNarrow) <> strNum Then
        blnWide = True
        strNum = StrConv(strNum, vbNarrow)
    End If

    '���̒l
    If IsNumeric(strNum) Then
        strNum = CLng(strNum) + 1
    Else
        strNum = Chr(Asc(strNum) + 1)
    End If

    '�S�p�̎��͑S�p�ɖ߂�
    If blnWide Then
        strNum = StrConv(strNum, vbWide)
    End If

    '�Ђ炪�Ȃ̎��͂Ђ炪�Ȃɖ߂�
    If blnHiragana Then
        strNum = StrConv(strNum, vbHiragana)
    End If

    '���[�̕�����A������
    GetNext = strL & strNum & strR
End Function

'*****************************************************************************
'[�T�v] �i���ԍ��̎��̒i���ԍ����擾����(�A�ԕ������Z�p�����̎��̂�)
'[����] ��F��P��
'[�ߒl] ��F��P�� �� ��Q��
'*****************************************************************************
Private Function GetNum(ByVal strOutLine As String) As String
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True '�����ӏ��̈�v�ɑΉ�
    objRegExp.Pattern = "[0-9]+|[�O-�X]+" '�S�p�܂��͔��p�̐��l���܂ނƂ�
    If Not objRegExp.Test(strOutLine) Then
        GetNum = strOutLine
        Exit Function
    End If

    Dim strL As String
    Dim strR As String
    Dim strNum As String
    Dim objMatches As Object
    Set objMatches = objRegExp.Execute(strOutLine)

    '�Z�p�����̉ӏ��������̎��́A��ԉE���̌���ΏۂƂ���
    With objMatches(objMatches.Count - 1)
        strL = Left(strOutLine, .FirstIndex)
        strR = Mid(strOutLine, .FirstIndex + .Length + 1)
        strNum = .Value
    End With

    Dim lngNum As Long
    Dim blnWide As Boolean

    '�S�p�����̎��͔��p�ɕϊ�
    If StrConv(strNum, vbNarrow) <> strNum Then
        blnWide = True
        lngNum = CLng(StrConv(strNum, vbNarrow))
    Else
        lngNum = CLng(strNum)
    End If

    strNum = CStr(lngNum + 1)

    '�S�p�����̎��͑S�p�ɖ߂�
    If blnWide Then
        strNum = StrConv(strNum, vbWide)
    End If

    '���[�̕�����A������
    GetNum = strL & strNum & strR
End Function


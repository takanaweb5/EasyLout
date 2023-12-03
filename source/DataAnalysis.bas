Attribute VB_Name = "DataAnalysis"
Option Explicit
Private Const C_CONNECTSTR = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={FileName};Extended Properties=""Excel 12.0;{Headers}"";"
Private Const adOpenStatic = 3
Private Const adLockReadOnly = 1

'***********************************************************************
'[�T�v] Google�X�v���b�h�V�[�g��QUERY()�֐����ǂ�
'[����] �f�[�^�͈̔́A�N�G��������
'       True:�ŏ��̍s���w�b�_�[�Ƃ��Ĉ���
'       NULL�\��(���l���ڂ�Null�̈���)�F1=��,2=0,3=#NULL!,4=�������#NULL!
'[�ߒl] ���s����(2�����z��)���X�s���Ŏ��o��
'*****************************************************************************
Public Function Query(ByRef �f�[�^�͈� As Range, ByVal �N�G�������� As String, Optional ���o�� As Boolean = True, Optional NULL�\�� As Byte = 1) As Variant
Attribute Query.VB_Description = "Google�X�v���b�h�V�[�g�Ɏ�������Ă���AQUERY()�֐����ǂ����������܂�\n��3�������ȗ����邩TRUE��ݒ肷���1�s�ڂ����o���Ƃ��Ĉ����A���o�������̂܂܃J�������ƂȂ�܂�\n��3������FALSE��ݒ肷���1�s�ڂ��疾�ׂƂ��Ĉ����A�J�������͍����珇�Ԃ�F1,F2..�ƂȂ�܂�\n��4�����͐��l���ڂ�NULL�ɂȂ������̏o�͕��@���w�肵�܂�\n�@1(�ȗ���)=��, 2=0, 3=#NULL!, 4=�������#NULL!"
On Error GoTo ErrHandle
    Dim strCon As String
    strCon = C_CONNECTSTR
    strCon = Replace(strCon, "{FileName}", �f�[�^�͈�.Worksheet.Parent.FullName)
    If ���o�� Then
        strCon = Replace(strCon, "{Headers}", "HDR=YES")
    Else
        strCon = Replace(strCon, "{Headers}", "HDR=NO")
    End If
    
    Dim strSQL As String
    strSQL = MakeSQL(�f�[�^�͈�, Trim(�N�G��������))
    
    'SQL�̎��s���ʂ̃��R�[�h�Z�b�g���擾
    Dim objRecordset As Object
    Set objRecordset = CreateObject("ADODB.Recordset")
    Call objRecordset.Open(strSQL, strCon, adOpenStatic, adLockReadOnly)
    
    Dim i As Long: Dim j As Long
    If ���o�� Then
        ReDim vData(0 To objRecordset.RecordCount, 0 To objRecordset.Fields.Count - 1) '(�s,��)
        '0�s�ڂɌ��o����ݒ肷��
        For i = 0 To objRecordset.Fields.Count - 1
            vData(0, i) = objRecordset.Fields(i).Name
        Next
    Else
        ReDim vData(1 To objRecordset.RecordCount, 0 To objRecordset.Fields.Count - 1) '(�s,��)
    End If
    
    '���ׂ̐ݒ�
    For j = 1 To objRecordset.RecordCount
        For i = 0 To objRecordset.Fields.Count - 1
            If IsNull(objRecordset.Fields(i).Value) Then
                '�����񂩂ǂ�������
                If objRecordset.Fields(i).Type >= 200 Then
                    If NULL�\�� = 4 Then
                        vData(j, i) = CVErr(xlErrNull)
                    Else
                        vData(j, i) = ""
                    End If
                Else
                    Select Case NULL�\��
                    Case 1
                        vData(j, i) = ""
                    Case 2
                        vData(j, i) = 0
                    Case Else
                        vData(j, i) = CVErr(xlErrNull)
                    End Select
                End If
            Else
                vData(j, i) = objRecordset.Fields(i).Value
            End If
        Next
        objRecordset.MoveNext
    Next
    
    Call objRecordset.Close
    Query = vData()
    Exit Function
ErrHandle:
    If �f�[�^�͈�.Worksheet.Parent.Path = "" Then
        Query = "��x���ۑ�����Ă��Ȃ��t�@�C���̓G���[�ɂȂ�܂�"
    Else
        '�G���[���b�Z�[�W��\��
        Query = Err.Description
    End If
End Function

'*****************************************************************************
'[�T�v] Query������ƃZ���͈͂��SQL�𐶐�����
'[����] �Z���͈́A�N�G��������
'[�ߒl] SQL
'*****************************************************************************
Private Function MakeSQL(ByRef objRange As Range, ByVal strQuery As String) As String
    'FROM��̐ݒ�
    Dim strFrom As String
    strFrom = Replace(" FROM [{Sheet}${Range}] ", "{Sheet}", objRange.Worksheet.Name)
    strFrom = Replace(strFrom, "{Range}", objRange.AddressLocal(False, False, xlA1))
    
    Dim objRegExp As Object
    Dim objSubMatches As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = "^SELECT"
    objRegExp.IgnoreCase = True '�啶������������ʂ��Ȃ�
    
    If objRegExp.test(strQuery) Then
        objRegExp.Pattern = "^(SELECT .*?)(WHERE|GROUP BY|HAVING|ORDER BY)(.*)$"
        If objRegExp.test(strQuery) Then
            Set objSubMatches = objRegExp.Execute(strQuery)(0).SubMatches
            MakeSQL = objSubMatches(0) & strFrom & objSubMatches(1) & objSubMatches(2)
        Else
            MakeSQL = strQuery & strFrom
        End If
    Else
        MakeSQL = "SELECT *" & strFrom & strQuery
    End If
End Function

'*****************************************************************************
'[�T�v] �̈���c�����Ɍ�������
'[����] �Ώۗ̈�
'[�ߒl] �������ʂ�2�����z��
'*****************************************************************************
Public Function UnionV(�̈�1, ParamArray �̈�2()) As Variant
Attribute UnionV.VB_Description = "�̈�Ɨ̈���c�����Ɍ������܂�"
    Dim arr As Variant
    Dim ColCount As Long
    Dim RowCount As Long
    
    '���͂�2�����z��ɕϊ�
    arr = ConvertTo2DArray(�̈�1)
    RowCount = UBound(arr, 1)
    ColCount = UBound(arr, 2)
    
    Dim vRange
    For Each vRange In �̈�2
        '���͂�2�����z��ɕϊ�
        arr = ConvertTo2DArray(vRange)
        RowCount = RowCount + UBound(arr, 1)
        If ColCount <> UBound(arr, 2) Then Call Err.Raise(513)
    Next
    
    ReDim Result(1 To RowCount, 1 To ColCount)
    Dim CurrentRow As Long
    '���͂�2�����z��ɕϊ�
    arr = ConvertTo2DArray(�̈�1)
    Call AppenResult(arr, CurrentRow, Result)
    
    For Each vRange In �̈�2
        '���͂�2�����z��ɕϊ�
        arr = ConvertTo2DArray(vRange)
        Call AppenResult(arr, CurrentRow, Result)
    Next
    UnionV = Result
End Function

'*****************************************************************************
'[�T�v] �̈���������Ɍ�������
'[����] �Ώۗ̈�
'[�ߒl] �������ʂ�2�����z��
'*****************************************************************************
Public Function UnionH(�̈�1, ParamArray �̈�2()) As Variant
Attribute UnionH.VB_Description = "�̈�Ɨ̈���������Ɍ������܂�"
    Dim arr  As Variant
    Dim ColCount As Long
    Dim RowCount As Long
    
    '���͂�2�����z��ɕϊ�
    arr = ConvertTo2DArray(�̈�1)
    RowCount = UBound(arr, 1)
    ColCount = UBound(arr, 2)
    
    Dim vRange
    For Each vRange In �̈�2
        '���͂�2�����z��ɕϊ�
        arr = ConvertTo2DArray(vRange)
        ColCount = ColCount + UBound(arr, 2)
        If RowCount <> UBound(arr, 1) Then Call Err.Raise(513)
    Next
    
    '�s������ւ������ʂ����߂�
    ReDim Result(1 To ColCount, 1 To RowCount)
    
    Dim CurrentCol As Long
    '���͂�2�����z��ɕϊ�
    arr = ConvertTo2DArray(�̈�1)
    '�s������ւ��Ď��s
    Call AppenResult(Transpose(arr), CurrentCol, Result)
    For Each vRange In �̈�2
        '���͂�2�����z��ɕϊ�
        arr = ConvertTo2DArray(vRange)
        '�s������ւ��Ď��s
        Call AppenResult(Transpose(arr), CurrentCol, Result)
    Next
    
    '�s������ɖ߂�
    UnionH = Transpose(Result)
End Function

'*****************************************************************************
'[�T�v] ���͂�2�����z��ɕϊ�
'[����] �ϊ���
'[�ߒl] 2�����z��
'*****************************************************************************
Private Function ConvertTo2DArray(ByRef vRange) As Variant
    'Range�̎��AVariant�z��ɕϊ�
    Dim arr
    arr = vRange
        
    If IsArray(arr) Then
        Select Case ArrayDims(arr)
        Case 1
            ReDim arr2(1 To 1, 1 To UBound(arr))
            Dim i As Long
            For i = 1 To UBound(arr)
                arr2(1, i) = arr(i)
            Next
            ConvertTo2DArray = arr2
        Case 2
            ConvertTo2DArray = arr
        Case Else
            Call Err.Raise(513)
        End Select
    Else
        ReDim arr2(1 To 1, 1 To 1)
        arr2(1, 1) = arr
        ConvertTo2DArray = arr2
    End If
End Function

'*****************************************************************************
'[�T�v] �s�Ɨ��u������(WorksheetFunction.Transpose�͒ᑬ�̂���)
'[����] �ϊ��O
'[�ߒl] �ϊ���
'*****************************************************************************
Private Function Transpose(ByRef arr) As Variant
    ReDim work(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 1) To UBound(arr, 1))
    Dim x As Long, y As Long
    For x = LBound(arr, 1) To UBound(arr, 1)
        For y = LBound(arr, 2) To UBound(arr, 2)
            work(y, x) = arr(x, y)
        Next
    Next
    Transpose = work
End Function

'*****************************************************************************
'[�T�v] �Ώۂ�2�����z��̒l��Result�z��Ɋi�[
'[����] arr:�Ώۂ̔z��Arow:�J�����g�s��(�l��i�߂ĕԂ�)�AResult:���ʔz��(�l��ǉ����ĕԂ�)
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub AppenResult(ByRef arr, ByRef CurrentRow As Long, ByRef Result)
    Dim CurrentCol As Long
    Dim row As Long
    Dim col As Long
    For row = LBound(arr, 1) To UBound(arr, 1)
        CurrentRow = CurrentRow + 1
        CurrentCol = 0
        For col = LBound(arr, 2) To UBound(arr, 2)
            CurrentCol = CurrentCol + 1
            If IsEmpty(arr(row, col)) Then
                Result(CurrentRow, CurrentCol) = ""
            Else
                Result(CurrentRow, CurrentCol) = arr(row, col)
            End If
        Next
    Next
End Sub

'*****************************************************************************
'[�T�v] �o���A���g�z��̎��������擾
'[����] �o���A���g�z��
'[�ߒl] ������
'*****************************************************************************
Private Function ArrayDims(arr) As Long
    Dim i As Long
    Dim tmp As Long
    On Error Resume Next
    Do While Err.Number = 0
        i = i + 1
        tmp = UBound(arr, i)
    Loop
    On Error GoTo 0
    ArrayDims = i - 1
End Function


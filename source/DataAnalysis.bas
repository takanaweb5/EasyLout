Attribute VB_Name = "DataAnalysis"
Option Explicit
Private Const C_CONNECTSTR = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={FileName};Extended Properties=""Excel 12.0;{Headers}"";"
Private Const adOpenStatic = 3
Private Const adLockReadOnly = 1

'***********************************************************************
'[�T�v] Google�X�v���b�h�V�[�g��QUERY()�֐����ǂ�
'[����] �f�[�^�͈̔́A�N�G��������
'       True:�ŏ��̍s���w�b�_�[�Ƃ��Ĉ���
'[�ߒl] ���s����(2�����z��)���X�s���Ŏ��o��
'*****************************************************************************
Public Function Query(ByRef �f�[�^�͈� As Range, ByVal �N�G�������� As String, Optional ���o�� As Boolean = True) As Variant
Attribute Query.VB_Description = "Google�X�v���b�h�V�[�g�Ɏ�������Ă���AQUERY()�֐����ǂ����������܂�\n1�s�ڂ����o���Ƃ��Ĉ������͑�3�������ȗ����邩TRUE��ݒ肵�A���o�������̂܂܃J�������ƂȂ�܂�\n1�s�ڂ��疾�ׂƂ��Ĉ������͑�3������FALSE��ݒ肵�A�J�������͍����珇�Ԃ�F1,F2..�ƂȂ�܂�"
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
            If IsNull(objRecordset.Fields(i).value) Then
                vData(j, i) = CVErr(xlErrNull)
            Else
                vData(j, i) = objRecordset.Fields(i).value
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
Public Function UnionV(�̈�1, ParamArray �̈�2())
Attribute UnionV.VB_Description = "�̈�Ɨ̈���c�����Ɍ������܂�"
    Dim arr As Variant
    Dim i As Long
    
    Dim ColCount As Long
    Dim RowCount As Long
    ReDim RowCntList(0 To 1 + UBound(�̈�2)) As Long
    
    'Range�̎��AVariant�z��ɕϊ�
    arr = �̈�1
        
    '�������𔻒�
    Select Case ArrayDims(arr)
    Case 1
        RowCntList(0) = 1
        ColCount = UBound(arr) - LBound(arr) + 1
    Case 2
        RowCntList(0) = UBound(arr, 1) - LBound(arr, 1) + 1
        ColCount = UBound(arr, 2) - LBound(arr, 2) + 1
    Case Else
        Call Err.Raise(513)
    End Select
    
    Dim vRange
    i = 1
    For Each vRange In �̈�2
        'Range�̎��AVariant�z��ɕϊ�
        arr = vRange
        '�������𔻒�
        Select Case ArrayDims(arr)
        Case 1
            RowCntList(i) = 1
            If ColCount <> UBound(arr) - LBound(arr) + 1 Then
                Call Err.Raise(513)
            End If
        Case 2
            RowCntList(i) = UBound(arr, 1) - LBound(arr, 1) + 1
            If ColCount <> UBound(arr, 2) - LBound(arr, 2) + 1 Then
                Call Err.Raise(513)
            End If
        Case Else
            Call Err.Raise(513)
        End Select
        i = i + 1
    Next
    
    'RowCount��z��̍��v���狁�߂�
    For i = 0 To UBound(RowCntList)
        RowCount = RowCount + RowCntList(i)
    Next

    ReDim Result(1 To RowCount, 1 To ColCount)
    Dim CurrentRow As Long
    'Range�̎��AVariant�z��ɕϊ�
    arr = �̈�1
    Call AppenResult(arr, CurrentRow, Result)
    
    For Each vRange In �̈�2
        'Range�̎��AVariant�z��ɕϊ�
        arr = vRange
        Call AppenResult(arr, CurrentRow, Result)
    Next
    UnionV = Result
End Function

'*****************************************************************************
'[�T�v] �̈���������Ɍ�������
'[����] �Ώۗ̈�
'[�ߒl] �������ʂ�2�����z��
'*****************************************************************************
Public Function UnionH(�̈�1, ParamArray �̈�2())
Attribute UnionH.VB_Description = "�̈�Ɨ̈���������Ɍ������܂�"
    Dim arr As Variant
    Dim i As Long
    
    Dim ColCount As Long
    Dim RowCount As Long
    ReDim ColCntList(0 To 1 + UBound(�̈�2)) As Long
    
    'Range�̎��AVariant�z��ɕϊ�
    arr = �̈�1
        
    '�������𔻒�
    Select Case ArrayDims(arr)
    Case 2
        ColCntList(0) = UBound(arr, 2) - LBound(arr, 2) + 1
        RowCount = UBound(arr, 1) - LBound(arr, 1) + 1
    Case Else
        Call Err.Raise(513)
    End Select
    
    Dim vRange
    i = 1
    For Each vRange In �̈�2
        'Range�̎��AVariant�z��ɕϊ�
        arr = vRange
        ColCntList(i) = UBound(arr, 2) - LBound(arr, 2) + 1
        If RowCount <> UBound(arr, 1) - LBound(arr, 1) + 1 Then
            Call Err.Raise(513)
        End If
        i = i + 1
    Next
    
    'ColCount��z��̍��v���狁�߂�
    For i = 0 To UBound(ColCntList)
        ColCount = ColCount + ColCntList(i)
    Next
    
    '�s������ւ������ʂ����߂�
    ReDim Result(1 To ColCount, 1 To RowCount)
    
    Dim CurrentCol As Long
    'Range�̎��AVariant�z��ɕϊ�
    arr = �̈�1
    '�s������ւ��Ď��s
    Call AppenResult(Transpose(arr), CurrentCol, Result)
    For Each vRange In �̈�2
        'Range�̎��AVariant�z��ɕϊ�
        arr = vRange
        '�s������ւ��Ď��s
        Call AppenResult(Transpose(arr), CurrentCol, Result)
    Next
    
    '�s������ɖ߂�
    UnionH = Transpose(Result)
End Function

'*****************************************************************************
'[�T�v] �s�Ɨ��u������(WorksheetFunction.Transpose�͒ᑬ�̂���)
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Function Transpose(ByRef arr)
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
'[�T�v] �̈�̒l��Result�z��Ɋi�[
'[����] arr:�Ώۗ̈�Arow:�J�����g�s��(�l��i�߂ĕԂ�)�AResult:���ʔz��(�l��ǉ����ĕԂ�)
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Function AppenResult(ByRef arr, ByRef CurrentRow As Long, ByRef Result)
    Dim col As Long
    Dim row As Long
    Dim i As Long
    Dim value
    '�z��̎������𔻒�
    Select Case ArrayDims(arr)
    Case 1
        CurrentRow = CurrentRow + 1
        For Each value In arr
            col = col + 1
            Result(CurrentRow, col) = value
        Next
    Case 2
        For row = LBound(arr, 1) To UBound(arr, 1)
            CurrentRow = CurrentRow + 1
            col = 0
            For i = LBound(arr, 2) To UBound(arr, 2)
                col = col + 1
                Result(CurrentRow, col) = arr(row, i)
            Next
        Next
    End Select
End Function

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


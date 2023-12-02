Attribute VB_Name = "DataAnalysis"
Option Explicit
Private Const C_CONNECTSTR = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={FileName};Extended Properties=""Excel 12.0;{Headers}"";"
Private Const adOpenStatic = 3
Private Const adLockReadOnly = 1

'***********************************************************************
'[概要] GoogleスプレッドシートのQUERY()関数もどき
'[引数] データの範囲、クエリ文字列
'       True:最初の行をヘッダーとして扱う
'[戻値] 実行結果(2次元配列)※スピルで取り出す
'*****************************************************************************
Public Function Query(ByRef データ範囲 As Range, ByVal クエリ文字列 As String, Optional 見出し As Boolean = True) As Variant
Attribute Query.VB_Description = "Googleスプレッドシートに実装されている、QUERY()関数もどきを実装します\n1行目を見出しとして扱う時は第3引数を省略するかTRUEを設定し、見出しがそのままカラム名となります\n1行目から明細として扱う時は第3引数にFALSEを設定し、カラム名は左から順番にF1,F2..となります"
On Error GoTo ErrHandle
    Dim strCon As String
    strCon = C_CONNECTSTR
    strCon = Replace(strCon, "{FileName}", データ範囲.Worksheet.Parent.FullName)
    If 見出し Then
        strCon = Replace(strCon, "{Headers}", "HDR=YES")
    Else
        strCon = Replace(strCon, "{Headers}", "HDR=NO")
    End If
    
    Dim strSQL As String
    strSQL = MakeSQL(データ範囲, Trim(クエリ文字列))
    
    'SQLの実行結果のレコードセットを取得
    Dim objRecordset As Object
    Set objRecordset = CreateObject("ADODB.Recordset")
    Call objRecordset.Open(strSQL, strCon, adOpenStatic, adLockReadOnly)
    
    Dim i As Long: Dim j As Long
    If 見出し Then
        ReDim vData(0 To objRecordset.RecordCount, 0 To objRecordset.Fields.Count - 1) '(行,列)
        '0行目に見出しを設定する
        For i = 0 To objRecordset.Fields.Count - 1
            vData(0, i) = objRecordset.Fields(i).Name
        Next
    Else
        ReDim vData(1 To objRecordset.RecordCount, 0 To objRecordset.Fields.Count - 1) '(行,列)
    End If
    
    '明細の設定
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
    If データ範囲.Worksheet.Parent.Path = "" Then
        Query = "一度も保存されていないファイルはエラーになります"
    Else
        'エラーメッセージを表示
        Query = Err.Description
    End If
End Function

'*****************************************************************************
'[概要] Query文字列とセル範囲よりSQLを生成する
'[引数] セル範囲、クエリ文字列
'[戻値] SQL
'*****************************************************************************
Private Function MakeSQL(ByRef objRange As Range, ByVal strQuery As String) As String
    'FROM句の設定
    Dim strFrom As String
    strFrom = Replace(" FROM [{Sheet}${Range}] ", "{Sheet}", objRange.Worksheet.Name)
    strFrom = Replace(strFrom, "{Range}", objRange.AddressLocal(False, False, xlA1))
    
    Dim objRegExp As Object
    Dim objSubMatches As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = "^SELECT"
    objRegExp.IgnoreCase = True '大文字小文字を区別しない
    
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
'[概要] 領域を縦方向に結合する
'[引数] 対象領域
'[戻値] 結合結果の2次元配列
'*****************************************************************************
Public Function UnionV(領域1, ParamArray 領域2())
Attribute UnionV.VB_Description = "領域と領域を縦方向に結合します"
    Dim arr As Variant
    Dim i As Long
    
    Dim ColCount As Long
    Dim RowCount As Long
    ReDim RowCntList(0 To 1 + UBound(領域2)) As Long
    
    'Rangeの時、Variant配列に変換
    arr = 領域1
        
    '次元数を判定
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
    For Each vRange In 領域2
        'Rangeの時、Variant配列に変換
        arr = vRange
        '次元数を判定
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
    
    'RowCountを配列の合計から求める
    For i = 0 To UBound(RowCntList)
        RowCount = RowCount + RowCntList(i)
    Next

    ReDim Result(1 To RowCount, 1 To ColCount)
    Dim CurrentRow As Long
    'Rangeの時、Variant配列に変換
    arr = 領域1
    Call AppenResult(arr, CurrentRow, Result)
    
    For Each vRange In 領域2
        'Rangeの時、Variant配列に変換
        arr = vRange
        Call AppenResult(arr, CurrentRow, Result)
    Next
    UnionV = Result
End Function

'*****************************************************************************
'[概要] 領域を横方向に結合する
'[引数] 対象領域
'[戻値] 結合結果の2次元配列
'*****************************************************************************
Public Function UnionH(領域1, ParamArray 領域2())
Attribute UnionH.VB_Description = "領域と領域を横方向に結合します"
    Dim arr As Variant
    Dim i As Long
    
    Dim ColCount As Long
    Dim RowCount As Long
    ReDim ColCntList(0 To 1 + UBound(領域2)) As Long
    
    'Rangeの時、Variant配列に変換
    arr = 領域1
        
    '次元数を判定
    Select Case ArrayDims(arr)
    Case 2
        ColCntList(0) = UBound(arr, 2) - LBound(arr, 2) + 1
        RowCount = UBound(arr, 1) - LBound(arr, 1) + 1
    Case Else
        Call Err.Raise(513)
    End Select
    
    Dim vRange
    i = 1
    For Each vRange In 領域2
        'Rangeの時、Variant配列に変換
        arr = vRange
        ColCntList(i) = UBound(arr, 2) - LBound(arr, 2) + 1
        If RowCount <> UBound(arr, 1) - LBound(arr, 1) + 1 Then
            Call Err.Raise(513)
        End If
        i = i + 1
    Next
    
    'ColCountを配列の合計から求める
    For i = 0 To UBound(ColCntList)
        ColCount = ColCount + ColCntList(i)
    Next
    
    '行列を入れ替えた結果を求める
    ReDim Result(1 To ColCount, 1 To RowCount)
    
    Dim CurrentCol As Long
    'Rangeの時、Variant配列に変換
    arr = 領域1
    '行列を入れ替えて実行
    Call AppenResult(Transpose(arr), CurrentCol, Result)
    For Each vRange In 領域2
        'Rangeの時、Variant配列に変換
        arr = vRange
        '行列を入れ替えて実行
        Call AppenResult(Transpose(arr), CurrentCol, Result)
    Next
    
    '行列を元に戻す
    UnionH = Transpose(Result)
End Function

'*****************************************************************************
'[概要] 行と列を置換する(WorksheetFunction.Transposeは低速のため)
'[引数] なし
'[戻値] なし
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
'[概要] 領域の値をResult配列に格納
'[引数] arr:対象領域、row:カレント行数(値を進めて返す)、Result:結果配列(値を追加して返す)
'[戻値] なし
'*****************************************************************************
Private Function AppenResult(ByRef arr, ByRef CurrentRow As Long, ByRef Result)
    Dim col As Long
    Dim row As Long
    Dim i As Long
    Dim value
    '配列の次元数を判定
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
'[概要] バリアント配列の次元数を取得
'[引数] バリアント配列
'[戻値] 次元数
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


Attribute VB_Name = "DataAnalysis"
Option Explicit
Private Const C_CONNECTSTR = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={FileName};Extended Properties=""Excel 12.0;{Headers}"";"
Private Const adOpenStatic = 3
Private Const adLockReadOnly = 1

'***********************************************************************
'[概要] GoogleスプレッドシートのQUERY()関数もどき
'[引数] データの範囲、クエリ文字列
'       True:最初の行をヘッダーとして扱う
'       NULL表現(数値項目のNullの扱い)：1=空白,2=0,3=#NULL!,4=文字列も#NULL!
'[戻値] 実行結果(2次元配列)※スピルで取り出す
'*****************************************************************************
Public Function Query(ByRef データ範囲 As Range, ByVal クエリ文字列 As String, Optional 見出し As Boolean = True, Optional NULL表現 As Byte = 1) As Variant
Attribute Query.VB_Description = "Googleスプレッドシートに実装されている、QUERY()関数もどきを実装します\n第3引数を省略するかTRUEを設定すると1行目を見出しとして扱い、見出しがそのままカラム名となります\n第3引数にFALSEを設定すると1行目から明細として扱い、カラム名は左から順番にF1,F2..となります\n第4引数は数値項目がNULLになった時の出力方法を指定します\n　1(省略時)=空白, 2=0, 3=#NULL!, 4=文字列も#NULL!"
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
            If IsNull(objRecordset.Fields(i).Value) Then
                '文字列かどうか判定
                If objRecordset.Fields(i).Type >= 200 Then
                    If NULL表現 = 4 Then
                        vData(j, i) = CVErr(xlErrNull)
                    Else
                        vData(j, i) = ""
                    End If
                Else
                    Select Case NULL表現
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
Public Function UnionV(領域1, ParamArray 領域2()) As Variant
Attribute UnionV.VB_Description = "領域と領域を縦方向に結合します"
    Dim arr As Variant
    Dim ColCount As Long
    Dim RowCount As Long
    
    '入力を2次元配列に変換
    arr = ConvertTo2DArray(領域1)
    RowCount = UBound(arr, 1)
    ColCount = UBound(arr, 2)
    
    Dim vRange
    For Each vRange In 領域2
        '入力を2次元配列に変換
        arr = ConvertTo2DArray(vRange)
        RowCount = RowCount + UBound(arr, 1)
        If ColCount <> UBound(arr, 2) Then Call Err.Raise(513)
    Next
    
    ReDim Result(1 To RowCount, 1 To ColCount)
    Dim CurrentRow As Long
    '入力を2次元配列に変換
    arr = ConvertTo2DArray(領域1)
    Call AppenResult(arr, CurrentRow, Result)
    
    For Each vRange In 領域2
        '入力を2次元配列に変換
        arr = ConvertTo2DArray(vRange)
        Call AppenResult(arr, CurrentRow, Result)
    Next
    UnionV = Result
End Function

'*****************************************************************************
'[概要] 領域を横方向に結合する
'[引数] 対象領域
'[戻値] 結合結果の2次元配列
'*****************************************************************************
Public Function UnionH(領域1, ParamArray 領域2()) As Variant
Attribute UnionH.VB_Description = "領域と領域を横方向に結合します"
    Dim arr  As Variant
    Dim ColCount As Long
    Dim RowCount As Long
    
    '入力を2次元配列に変換
    arr = ConvertTo2DArray(領域1)
    RowCount = UBound(arr, 1)
    ColCount = UBound(arr, 2)
    
    Dim vRange
    For Each vRange In 領域2
        '入力を2次元配列に変換
        arr = ConvertTo2DArray(vRange)
        ColCount = ColCount + UBound(arr, 2)
        If RowCount <> UBound(arr, 1) Then Call Err.Raise(513)
    Next
    
    '行列を入れ替えた結果を求める
    ReDim Result(1 To ColCount, 1 To RowCount)
    
    Dim CurrentCol As Long
    '入力を2次元配列に変換
    arr = ConvertTo2DArray(領域1)
    '行列を入れ替えて実行
    Call AppenResult(Transpose(arr), CurrentCol, Result)
    For Each vRange In 領域2
        '入力を2次元配列に変換
        arr = ConvertTo2DArray(vRange)
        '行列を入れ替えて実行
        Call AppenResult(Transpose(arr), CurrentCol, Result)
    Next
    
    '行列を元に戻す
    UnionH = Transpose(Result)
End Function

'*****************************************************************************
'[概要] 入力を2次元配列に変換
'[引数] 変換元
'[戻値] 2次元配列
'*****************************************************************************
Private Function ConvertTo2DArray(ByRef vRange) As Variant
    'Rangeの時、Variant配列に変換
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
'[概要] 行と列を置換する(WorksheetFunction.Transposeは低速のため)
'[引数] 変換前
'[戻値] 変換後
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
'[概要] 対象の2次元配列の値をResult配列に格納
'[引数] arr:対象の配列、row:カレント行数(値を進めて返す)、Result:結果配列(値を追加して返す)
'[戻値] なし
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


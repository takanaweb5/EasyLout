Attribute VB_Name = "GeneralTools"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal uIDEnableItem As Long, ByVal uEnable As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal PROCESS As Long, lpExitCode As Long) As Long
'Public Declare Function CloseHandle Lib "KERNEL32.DLL" (ByVal hObject As Long) As Long
'Public Declare Function TerminateProcess Lib "KERNEL32.DLL" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' 定数の定義
Public Const IDC_HAND = 32649
Public Const SC_CLOSE = 61536
Public Const GWL_WNDPROC = (-4)
Public Const WM_RBUTTONDOWN = &H204 '右マウスボタンを押した
Public Const WM_MOUSEWHEEL = &H20A  'ホイールが回された（Win98,NT4.0以降）
Public Const MF_BYCOMMAND = 0
Public Const MF_GRAYED = 1
Public Const GWL_STYLE = (-16)
Public Const WS_THICKFRAME = &H40000 'ウィンドウのサイズ変更
Public Const WS_MINIMIZEBOX = &H20000 '最小化ボタン
Public Const WS_MAXIMIZEBOX = &H10000 '最大化ボタン
Public Const SW_SHOWNORMAL = 1
Public Const SW_MAXIMIZE = 3
Public Const SYNCHRONIZE       As Long = &H100000
Public Const PROCESS_TERMINATE As Long = &H1
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103

'選択タイプ
Public Enum ESelectionType
    E_Range
    E_Shape
    E_Other
End Enum

'*****************************************************************************
'[ 関数名 ]　CheckSelection
'[ 概  要 ]　選択されているかオブジェクトの種類を判定する
'[ 引  数 ]　なし
'[ 戻り値 ]　Range、Shape、その他　のいずれか
'*****************************************************************************
Public Function CheckSelection() As ESelectionType
On Error GoTo ErrHandle
    If Selection Is Nothing Then
        CheckSelection = E_Other
        Exit Function
    End If
    
    If TypeOf Selection Is Range Then
        CheckSelection = E_Range
    ElseIf TypeOf Selection.ShapeRange Is ShapeRange Then
        CheckSelection = E_Shape
    Else
        CheckSelection = E_Other
    End If
Exit Function
ErrHandle:
    CheckSelection = E_Other
End Function

'*****************************************************************************
'[ 関数名 ]　GetRangeText
'[ 概  要 ]　各行の文字列を改行で、各列の文字列を空白で区切って連結する
'[ 引  数 ]　対象の領域
'[ 戻り値 ]　連結された文字列
'*****************************************************************************
Public Function GetRangeText(ByRef objRange As Range) As String
    Dim i   As Long
    
    '行の数だけループ
    For i = 1 To objRange.Rows.Count
        GetRangeText = GetRangeText & GetRowText(objRange.Rows(i)) & vbLf
    Next i
    
    '先頭と末尾の空白行を削除
    GetRangeText = TrimChr(GetRangeText)
End Function

'*****************************************************************************
'[ 関数名 ]　GetRowText
'[ 概  要 ]　各列の文字列を空白で区切って連結する
'[ 引  数 ]　対象の１行
'[ 戻り値 ]　連結された文字列
'*****************************************************************************
Private Function GetRowText(ByRef objRange As Range) As String
    Dim i       As Long
    Dim strText As String
    
    '列の数だけループ
    For i = 1 To objRange.Columns.Count
        strText = GetCellText(objRange.Cells(1, i))
        GetRowText = GetRowText & strText & vbTab
    Next i

    '末尾のTabを削除
    GetRowText = RTrimChr(GetRowText, vbTab)
End Function

'*****************************************************************************
'[ 関数名 ]　GetCellText
'[ 概  要 ]　Cellの文字列を表示された書式で取得する
'[ 引  数 ]　対象のセル
'[ 戻り値 ]　文字列
'*****************************************************************************
Public Function GetCellText(ByRef objCell As Range) As String
On Error GoTo ErrHandle
    If objCell.Text <> "" Then
        Select Case objCell.NumberFormat
        Case "General", "@"
            GetCellText = objCell.Value
        Case Else
            GetCellText = WorksheetFunction.Text(objCell, objCell.NumberFormat)
        End Select
    End If
    GetCellText = RTrim$(GetCellText)
Exit Function
ErrHandle:
    GetCellText = RTrim$(objCell.Value)
End Function

'*****************************************************************************
'[ 関数名 ]　TrimChr
'[ 概  要 ]　文字列の先頭と末尾の改行やタブ文字を削除する
'[ 引  数 ]　削除する文字
'[ 戻り値 ]　文字列
'*****************************************************************************
Public Function TrimChr(ByVal strText As String, Optional ByVal strChr As String = vbLf) As String
    TrimChr = LTrimChr(strText, strChr)
    TrimChr = RTrimChr(TrimChr, strChr)
End Function

'*****************************************************************************
'[ 関数名 ]　LTrimChr
'[ 概  要 ]　文字列の先頭の改行やタブ文字を削除する
'[ 引  数 ]　削除する文字
'[ 戻り値 ]　文字列
'*****************************************************************************
Public Function LTrimChr(ByVal strText As String, Optional ByVal strChr As String = " ") As String
    Dim i        As Long
    Dim lngStart As Long
    
    '前方よりループ
    For i = 1 To Len(strText)
        If Mid$(strText, i, 1) <> strChr Then
            lngStart = i
            Exit For
        End If
    Next
    
    If lngStart > 0 Then
        LTrimChr = Mid$(strText, lngStart)
    End If
End Function

'*****************************************************************************
'[ 関数名 ]　RTrimChr
'[ 概  要 ]　文字列の末尾の改行やタブ文字を削除する
'[ 引  数 ]　削除する文字
'[ 戻り値 ]　文字列
'*****************************************************************************
Public Function RTrimChr(ByVal strText As String, Optional ByVal strChr As String = " ") As String
    Dim i        As Long
    Dim lngEnd   As Long
    
    '後方よりループ
    For i = Len(strText) To 1 Step -1
        If Mid$(strText, i, 1) <> strChr Then
            lngEnd = i
            Exit For
        End If
    Next
    
    If lngEnd > 0 Then
        RTrimChr = Left$(strText, lngEnd)
    End If
End Function

'*****************************************************************************
'[ 関数名 ]　GetStrArray
'[ 概  要 ]　文字列を改行でばらして１行ごとの配列で返す
'[ 引  数 ]　strText:元の文字列、StrArray:1行ごとの配列
'[ 戻り値 ]　行数
'*****************************************************************************
Public Function GetStrArray(ByVal strText As String, Optional ByRef strArray As Variant) As Long
    '改行または空白だけ
    If Trim$(Replace$(strText, vbLf, "")) = "" Then
        GetStrArray = 0
        Exit Function
    End If
    
    '１行ごとに配列に格納
    strArray = Split(TrimChr(strText), vbLf)
    
    GetStrArray = UBound(strArray) + 1
End Function

'*****************************************************************************
'[ 関数名 ]　IntersectRange
'[ 概　要 ]　領域と領域の重なる領域を取得する
'　　　　　　Ａ∩Ｂ
'[ 引　数 ]　対象領域(Nothingも可)
'[ 戻り値 ]　objRange1 ∩ objRange2
'*****************************************************************************
Public Function IntersectRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Select Case True
    Case (objRange1 Is Nothing) Or (objRange2 Is Nothing)
        Set IntersectRange = Nothing
    Case Else
        Set IntersectRange = Intersect(objRange1, objRange2)
    End Select
End Function

'*****************************************************************************
'[ 関数名 ]　UnionRange
'[ 概　要 ]　領域に領域を加える
'　　　　　　Ａ∪Ｂ
'[ 引　数 ]　対象領域(Nothingも可)
'[ 戻り値 ]　objRange1 ∪ objRange2
'*****************************************************************************
Public Function UnionRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Select Case True
    Case (objRange1 Is Nothing) And (objRange2 Is Nothing)
        Set UnionRange = Nothing
    Case (objRange1 Is Nothing)
        Set UnionRange = objRange2
    Case (objRange2 Is Nothing)
        Set UnionRange = objRange1
    Case Else
        Set UnionRange = Union(objRange1, objRange2)
    End Select
End Function

'*****************************************************************************
'[ 関数名 ]　MinusRange
'[ 概　要 ]　領域から領域を、除外する
'　　　　　　Ａ－Ｂ = Ａ∩!Ｂ
'　　　　　　!Ｂ = !(B1∪B2∪B3...∪Bn) = !B1∩!B2∩!B3...∩!Bn
'[ 引　数 ]　対象領域
'[ 戻り値 ]　objRange1 － objRange2
'*****************************************************************************
Public Function MinusRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Dim objRounds As Range
    Dim i As Long
    
    If objRange2 Is Nothing Then
        Set MinusRange = objRange1
        Exit Function
    End If
    
    '除外する領域の数だけループ
    '!Ｂ = !B1∩!B2∩!B3.....∩!Bn
    Set objRounds = ReverseRange(objRange2.Areas(1))
    For i = 2 To objRange2.Areas.Count
        Set objRounds = IntersectRange(objRounds, ReverseRange(objRange2.Areas(i)))
    Next
    
    'Ａ∩!Ｂ
    Set MinusRange = IntersectRange(objRange1, objRounds)
End Function

'*****************************************************************************
'[ 関数名 ]　ArrangeRange
'[ 概　要 ]　Select出来るセルに整理する、領域の重複をなくす
'[ 引　数 ]　対象領域
'[ 戻り値 ]　整理した領域
'*****************************************************************************
Public Function ArrangeRange(ByRef objRange As Range) As Range
    Dim objUsedRange As Range
    Dim objArea      As Range
    
    If objRange Is Nothing Then
        Exit Function
    End If
    
    '領域ごとに整理する
    For Each objArea In objRange.Areas
        Set ArrangeRange = UnionRange(ArrangeRange, ArrangeRange2(objArea))
    Next
    
    '最後のセル以降の領域を足す
    With objRange.Worksheet
        Set objUsedRange = .Range(.Cells(1, 1), .Cells.SpecialCells(xlLastCell))
    End With
    Set ArrangeRange = UnionRange(ArrangeRange, MinusRange(objRange, objUsedRange))
End Function

'*****************************************************************************
'[ 関数名 ]　ArrangeRange2
'[ 概　要 ]　Select出来るセルに整理する、領域の重複をなくす
'[ 引　数 ]　対象領域
'[ 戻り値 ]　整理した領域
'*****************************************************************************
Private Function ArrangeRange2(ByRef objRange As Range) As Range
    Dim objArrange(1 To 3) As Range
    Dim i As Long
    
    If objRange.Count = 1 Then
        Set ArrangeRange2 = objRange.MergeArea
        Exit Function
    End If
    
    If IsOnlyCell(objRange) Then
        Set ArrangeRange2 = objRange
        Exit Function
    End If
    
    With objRange
        On Error Resume Next
        'すべてのセルを結合セルに応じて選択する
        Set objArrange(1) = .SpecialCells(xlCellTypeConstants)
        Set objArrange(2) = .SpecialCells(xlCellTypeFormulas)
        Set objArrange(3) = .SpecialCells(xlCellTypeBlanks)
        On Error GoTo 0
    End With
    
    For i = 1 To 3
        Set ArrangeRange2 = UnionRange(ArrangeRange2, objArrange(i))
    Next
End Function

'*****************************************************************************
'[ 関数名 ]　ReverseRange
'[ 概　要 ]　領域を反転する
'[ 引　数 ]　対象領域
'[ 戻り値 ]　Not objRange
'*****************************************************************************
Private Function ReverseRange(ByRef objRange As Range) As Range
    Dim i As Long
    Dim objRound(1 To 4) As Range
    
    With objRange.Parent
        On Error Resume Next
        '選択領域より上の領域すべて
        Set objRound(1) = .Range(.Rows(1), _
                                 .Rows(objRange.Row - 1))
        '選択領域より下の領域すべて
        Set objRound(2) = .Range(.Rows(objRange.Row + objRange.Rows.Count), _
                                 .Rows(Rows.Count))
        '選択領域より左の領域すべて
        Set objRound(3) = .Range(.Columns(1), _
                                 .Columns(objRange.Column - 1))
        '選択領域より右の領域すべて
        Set objRound(4) = .Range(.Columns(objRange.Column + objRange.Columns.Count), _
                                 .Columns(Columns.Count))
        On Error GoTo 0
    End With
    
    '選択領域以外の領域を設定
    For i = 1 To 4
        Set ReverseRange = UnionRange(ReverseRange, objRound(i))
    Next
End Function

'*****************************************************************************
'[ 関数名 ]　GetNearlyRange
'[ 概  要 ]  Shapeの四方に最も近いセル範囲を取得する
'[ 引  数 ]　Shapeオブジェクト
'[ 戻り値 ]　セル範囲
'*****************************************************************************
Public Function GetNearlyRange(ByRef objShape As Shape) As Range
    Dim objTopLeft     As Range
    Dim objBottomRight As Range
    Set objTopLeft = objShape.TopLeftCell
    Set objBottomRight = objShape.BottomRightCell
    
    '上の位置と高さを設定
    If objShape.Height = 0 Then
        With objTopLeft
            If .Top + .Height / 2 < objShape.Top Then
                Set objTopLeft = Cells(.Row + 1, .Column)
                Set objBottomRight = Cells(.Row + 1, objBottomRight.Column)
            End If
        End With
    Else
        '下のセルの再設定
        With objBottomRight
            If .Top = objShape.Top + objShape.Height Then
                Set objBottomRight = Cells(.Row - 1, .Column)
            End If
        End With
            
        '上端の再設定
        With objTopLeft
            If .Top + .Height / 2 < objShape.Top Then
                If .Row + 1 <= objBottomRight.Row Then
                    Set objTopLeft = Cells(.Row + 1, .Column)
                End If
            End If
        End With
                
        '下端の再設定
        With objBottomRight
            If .Top + .Height / 2 > objShape.Top + objShape.Height Then
                If .Row - 1 >= objTopLeft.Row Then
                    Set objBottomRight = Cells(.Row - 1, .Column)
                End If
            End If
        End With
    End If
    
    '左の位置と幅を設定
    If objShape.Width = 0 Then
        With objTopLeft
            If .Left + .Width / 2 < objShape.Left Then
                Set objTopLeft = Cells(.Row, .Column + 1)
                Set objBottomRight = Cells(objBottomRight.Row, .Column + 1)
            End If
        End With
    Else
        '右のセルの再設定
        With objBottomRight
            If .Left = objShape.Left + objShape.Width Then
                Set objBottomRight = Cells(.Row, .Column - 1)
            End If
        End With
    
        '左端の再設定
        With objTopLeft
            If .Left + .Width / 2 < objShape.Left Then
                If .Column + 1 <= objBottomRight.Column Then
                    Set objTopLeft = Cells(.Row, .Column + 1)
                End If
            End If
        End With
                
        '右端の再設定
        With objBottomRight
            If .Left + .Width / 2 > objShape.Left + objShape.Width Then
                If .Column - 1 >= objTopLeft.Column Then
                    Set objBottomRight = Cells(.Row, .Column - 1)
                End If
            End If
        End With
    End If
    
    Set GetNearlyRange = Range(objTopLeft, objBottomRight)
End Function

'*****************************************************************************
'[ 関数名 ]　GetCopyRangeAddress
'[ 概  要 ]　Copy対象のRangeのAddressを取得
'[ 引  数 ]　なし
'[ 戻り値 ]　例：[Book1]Sheet1!$A$1:$B$1
'*****************************************************************************
Public Function GetCopyRangeAddress() As String
On Error GoTo ErrHandle
    Application.DisplayAlerts = False
    
    Dim objWorksheet As Worksheet
    Set objWorksheet = ThisWorkbook.Worksheets("Workarea1")
    With objWorksheet.Pictures.Paste(Link:=True)
        GetCopyRangeAddress = .Formula
        Call .Delete
    End With
    
    GetCopyRangeAddress = GetMergeAddress(GetCopyRangeAddress)
    
    Application.DisplayAlerts = True
Exit Function
ErrHandle:
    Application.DisplayAlerts = True
    Dim strMsg As String
    strMsg = "コピー元のセルの取得に失敗しました。以下の点を確認してください。" & vbCrLf
    strMsg = strMsg & "複数の範囲をコピーして実行できません。" & vbCrLf
    strMsg = strMsg & "ファイルのパスが長すぎると実行できません。"
    Call Err.Raise(Err.Number, Err.Source, strMsg)
End Function

'*****************************************************************************
'[ 関数名 ]　CheckDupRange
'[ 概  要 ]　領域に重複がないかどうか判定する
'[ 引  数 ]　判定する領域
'[ 戻り値 ]　True：重複あり
'*****************************************************************************
Public Function CheckDupRange(ByRef objAreas As Range) As Boolean
    Dim objRange   As Range
    Dim objWkRange As Range
    
    For Each objRange In objAreas.Areas
        If IntersectRange(objWkRange, objRange) Is Nothing Then
            Set objWkRange = UnionRange(objWkRange, objRange)
        Else
            CheckDupRange = True
            Exit Function
        End If
    Next objRange
End Function

'*****************************************************************************
'[ 関数名 ]　SearchValueCell
'[ 概  要 ]　値の入力されているセルを検索する
'[ 引  数 ]　objRange：検索範囲
'[ 戻り値 ]　値の入力されているセル
'*****************************************************************************
Public Function SearchValueCell(ByRef objRange As Range) As Range
    Dim objWkRange(0 To 1)  As Range
    
    On Error Resume Next
    With objRange
        Set objWkRange(0) = .SpecialCells(xlCellTypeConstants)
        Set objWkRange(1) = .SpecialCells(xlCellTypeFormulas)
    End With
    On Error GoTo 0
    Set SearchValueCell = UnionRange(objWkRange(0), objWkRange(1))
End Function

'*****************************************************************************
'[ 関数名 ]　GetSheeetShapeRange
'[ 概  要 ]　ワークシートのShpesオブジェクトをShapeRangeオブジェクトに変換
'[ 引  数 ]　ワークシート
'[ 戻り値 ]　ShapeRangeオブジェクト
'*****************************************************************************
Public Function GetSheeetShapeRange(ByRef objSheet As Worksheet) As ShapeRange
    Dim i As Long
    If objSheet.Shapes.Count = 0 Then
        Exit Function
    End If
    ReDim lngArray(1 To objSheet.Shapes.Count)
    For i = 1 To objSheet.Shapes.Count
        lngArray(i) = i
    Next
    Set GetSheeetShapeRange = objSheet.Shapes.Range(lngArray)
End Function

'*****************************************************************************
'[ 関数名 ]　DeleteSheet
'[ 概  要 ]　ワークシートの中身を削除する
'[ 引  数 ]　対象のワークシート
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub DeleteSheet(ByRef objSheet As Worksheet)
    Dim objShape  As Shape
    For Each objShape In objSheet.Shapes
        Call objShape.Delete
    Next objShape
    
    With objSheet.Cells
        Call .Clear
        Call .Delete
    End With

    '最後のセルを修正する
    Call objSheet.Cells.Parent.UsedRange
End Sub

'*****************************************************************************
'[ 関数名 ]　GetMergeAddress
'[ 概  要 ]　結合セル1つだけの時、左上のアドレスしか返らないので、全体を返す
'[ 引  数 ]　対象アドレス
'[ 戻り値 ]　アドレス
'*****************************************************************************
Public Function GetMergeAddress(ByVal strAddress As String) As String
    GetMergeAddress = strAddress
    With Range(strAddress)
        If .Rows.Count = 1 And .Columns.Count = 1 Then
            If .MergeCells = True Then
                With .MergeArea
                    GetMergeAddress = strAddress & ":" & .Cells(.Rows.Count, .Columns.Count).Address
                End With
            End If
        End If
    End With
End Function

'*****************************************************************************
'[ 関数名 ]　OffsetRange
'[ 概　要 ]　RangeをOffset分移動する
'            結合セルがあると単なるOffsetメソッドが想定外の動作をするため
'[ 引　数 ]　元の領域、行のOffset、列のOffset
'[ 戻り値 ]　図形をスライドさせる領域
'*****************************************************************************
Public Function OffsetRange(ByRef objRange As Range, Optional ByVal lngRowOffset As Long = 0, Optional ByVal lngColOffset As Long = 0) As Range
    Dim objCell(1 To 2) As Range '1:左上、2:右下
    
    With objRange(1)
        Set objCell(1) = objRange.Worksheet.Cells(.Row + lngRowOffset, .Column + lngColOffset)
    End With
    With objRange(objRange.Count)
        Set objCell(2) = objRange.Worksheet.Cells(.Row + lngRowOffset, .Column + lngColOffset)
    End With
    
    Set OffsetRange = objRange.Worksheet.Range(objCell(1), objCell(2))
End Function

'*****************************************************************************
'[ 関数名 ]　ReSizeRange
'[ 概　要 ]　RangeをOffset分拡大縮小する
'            結合セルがあると単なるResizeメソッドが想定外の動作をするため
'[ 引　数 ]　元の領域、行のOffset、列のOffset
'[ 戻り値 ]　図形を拡大縮小させる領域
'*****************************************************************************
Public Function ReSizeRange(ByRef objRange As Range, Optional ByVal lngRowOffset As Long = 0, Optional ByVal lngColOffset As Long = 0) As Range
    Dim objCell(1 To 2) As Range '1:左上、2:右下
    
    Set objCell(1) = objRange(1)
    With objRange(objRange.Count)
        Set objCell(2) = objRange.Worksheet.Cells(.Row + lngRowOffset, .Column + lngColOffset)
    End With
    
    Set ReSizeRange = objRange.Worksheet.Range(objCell(1), objCell(2))
End Function

'*****************************************************************************
'[ 関数名 ]　IsBorderMerged
'[ 概  要 ]　Rangeの境界が結合セルかどうか
'[ 引  数 ]　判定するRange
'[ 戻り値 ]　True:境界に結合セルあり、False:なし
'*****************************************************************************
Public Function IsBorderMerged(ByRef objRange As Range) As Boolean
    IsBorderMerged = Not (MinusRange(ArrangeRange(objRange), objRange) Is Nothing)
End Function

'*****************************************************************************
'[ 関数名 ]　IsOnlyCell
'[ 概  要 ]　Rangeが(結合された)単一のセルかどうか
'[ 引  数 ]　判定するRange
'[ 戻り値 ]　True:単一のセル、False:複数のセル
'*****************************************************************************
Public Function IsOnlyCell(ByRef objRange As Range) As Boolean
    IsOnlyCell = (objRange.Address = objRange(1, 1).MergeArea.Address)
End Function

'*****************************************************************************
'[ 関数名 ]　StrConvWideExceptKana
'[ 概  要 ]　カタカナ以外半角に変換
'[ 引  数 ]　変換前の文字列
'[ 戻り値 ]　変換後の文字列
'*****************************************************************************
Public Function StrConvWideExceptKana(ByVal strText As String) As String
    Dim i           As Long
    Dim strWideChar As String
    
    '文字数分ループ
    For i = 1 To Len(strText)
        strWideChar = Mid$(strText, i, 1)
        Select Case AscW(strWideChar)
        Case AscW("ア") To AscW("ン"), AscW("ァ"), AscW("ヴ"), AscW("ー")
            StrConvWideExceptKana = StrConvWideExceptKana & strWideChar
        Case Else
            StrConvWideExceptKana = StrConvWideExceptKana & StrConv(strWideChar, vbNarrow)
        End Select
    Next
End Function

'*****************************************************************************
'[ 関数名 ]　GetClipbordText
'[ 概  要 ]　クリップボードのテキストを取得する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Function GetClipbordText() As String
On Error GoTo ErrHandle
    Dim objCb As New DataObject
    Call objCb.GetFromClipboard
    
    'テキスト形式が保持されている時
    If objCb.GetFormat(1) Then
        GetClipbordText = objCb.GetText
    End If
ErrHandle:
    Set objCb = Nothing
End Function

'*****************************************************************************
'[ 関数名 ]　SetClipbordText
'[ 概  要 ]　クリップボードにテキストを設定する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetClipbordText(ByVal strText As String)
On Error GoTo ErrHandle
    Dim objCb As New DataObject
    Call objCb.Clear
    Call objCb.SetText(strText)
    Call objCb.PutInClipboard
ErrHandle:
    Set objCb = Nothing
End Sub

'*****************************************************************************
'[ 関数名 ]　ClearClipbord
'[ 概  要 ]　クリップボードのクリア
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ClearClipbord()
    ThisWorkbook.Worksheets("Commands").Range("I1").Copy
    Application.CutCopyMode = False
End Sub


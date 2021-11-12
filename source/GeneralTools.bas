Attribute VB_Name = "GeneralTools"
Option Explicit
Option Private Module

Public Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr
Public Declare PtrSafe Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As LongPtr, ByVal uIDEnableItem As Long, ByVal uEnable As Long) As Long
Public Declare PtrSafe Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
'Public Declare PtrSafe Function CloseHandle Lib "KERNEL32.DLL" (ByVal hObject As Longptr) As Long
'Public Declare PtrSafe Function TerminateProcess Lib "KERNEL32.DLL" (ByVal hProcess As Longptr, ByVal uExitCode As Long) As Long
Public Declare PtrSafe Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
Public Declare PtrSafe Function SetCursor Lib "user32.dll" (ByVal hCursor As LongPtr) As LongPtr
Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
'Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Longptr, ByVal hwnd As Longptr, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare PtrSafe Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ImmSetOpenStatus Lib "imm32.dll" (ByVal himc As LongPtr, ByVal b As Long) As Long
Public Declare PtrSafe Function ImmReleaseContext Lib "imm32.dll" (ByVal hwnd As LongPtr, ByVal himc As LongPtr) As Long
Public Declare PtrSafe Function ReleaseCapture Lib "user32.dll" () As Long
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Public Declare PtrSafe Function GlobalAlloc Lib "Kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As LongPtr
Public Declare PtrSafe Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Public Declare PtrSafe Function GlobalSize Lib "Kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GlobalLock Lib "Kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "Kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

' 定数の定義
Public Const IDC_HAND = 32649
Public Const IDC_SIZENWSE = 32642
Public Const SC_CLOSE = 61536
Public Const SC_SIZE = &HF000&
Public Const GWL_WNDPROC = (-4)
Public Const WM_SYSCOMMAND = &H112
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
Public Const MAXROWCOLCNT = 1000
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90

Public Const REGKEY = "EasyLout"
Public Const DEFAULTFONT = "ＭＳ ゴシック"

'Public DPIRatio As Double

'選択タイプ
Public Enum ESelectionType
    E_Range
    E_Shape
    E_Non
    E_Other
End Enum

'結合タイプ
Public Enum EMergeType
    E_MTROW
    E_MTCOL
    E_MTBOTH
End Enum

'ソート用構造体
Public Type TSortArray
    Key1  As Long
    Key2  As Long
    Key3  As Long
End Type

'*****************************************************************************
'[ 関数名 ]　CheckSelection
'[ 概  要 ]　選択されているかオブジェクトの種類を判定する
'[ 引  数 ]　なし
'[ 戻り値 ]　Range、Shape、その他　のいずれか
'*****************************************************************************
Public Function CheckSelection() As ESelectionType
On Error GoTo ErrHandle
    If ActiveWorkbook Is Nothing Then
        CheckSelection = E_Non
        Exit Function
    End If
    
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
    Dim lngLast    As Long
    
    'すべての行の選択時
    If objRange.Rows.Count = Rows.Count Then
        '使用されている最後の行
        lngLast = Cells.SpecialCells(xlCellTypeLastCell).Row
    Else
        lngLast = objRange.Rows.Count
    End If
    
    '行の数だけループ
    For i = 1 To lngLast
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
    Select Case objCell.NumberFormat
    Case "General", "@"
        GetCellText = Rtrim$(objCell.Value)
        Exit Function
    End Select
                
    If objCell.Text <> WorksheetFunction.Rept("#", Len(objCell.Text)) Then
        GetCellText = Rtrim$(objCell.Text)
        Exit Function
    End If

    If IsDate(objCell.Value) Then
        GetCellText = WorksheetFunction.Text(objCell.Value, objCell.NumberFormatLocal)
        Exit Function
    End If
    
    If IsNumeric(objCell.Value) Then
        GetCellText = objCell.Value
        Exit Function
    End If
ErrHandle:
    GetCellText = Rtrim$(objCell.Value)
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
    Dim objArea      As Range
    
    If objRange Is Nothing Then
        Exit Function
    End If
    
    '領域ごとに整理する
    For Each objArea In objRange.Areas
        Set ArrangeRange = UnionRange(ArrangeRange, ArrangeRange2(objArea))
    Next
    
    '最後のセル以降の領域を足す
    Set ArrangeRange = UnionRange(ArrangeRange, MinusRange(objRange, GetUsedRange(objRange.Worksheet)))
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
'[ 関数名 ]　ReSelectRange
'[ 概　要 ]　新しい領域を、元の領域の選択ごとのエリアに分割する
'　　　　　　例:ReSelectRange(Range("A1,A2,A3"),Range("A1:A2")).Address→"A1,A2"
'[ 引　数 ]　objSelection:元の領域、objNewRange:新しい領域
'[ 戻り値 ]  objNewRangeを元の領域の選択ごとのエリアに分割したもの
'*****************************************************************************
Public Function ReSelectRange(ByRef objSelection As Range, ByRef objNewRange As Range) As Range
    Dim objTmpRange As Range
    Dim i As Long
    Dim strAddress As String
    Dim strRange   As String
        
    For i = 1 To objSelection.Areas.Count
        Set objTmpRange = IntersectRange(objSelection.Areas(i), objNewRange)
        If Not (objTmpRange Is Nothing) Then
            strRange = objTmpRange.Address(False, False)
            If Not (MinusRange(objTmpRange, Range(strRange)) Is Nothing) Then
                Set ReSelectRange = objNewRange
                Exit Function
            End If
            strAddress = strAddress & strRange & ","
        End If
    Next i
    
    '末尾のカンマを削除
    strAddress = Left$(strAddress, Len(strAddress) - 1)
    If Len(strAddress) < 256 Then
        Set ReSelectRange = Range(strAddress)
    Else
        Set ReSelectRange = objNewRange
    End If
End Function

'*****************************************************************************
'[ 関数名 ]　GetRowMergeRange
'[ 概　要 ]　結合された領域を取得する
'[ 引　数 ]　結合タイプ、対象領域
'[ 戻り値 ]　結合された領域
'*****************************************************************************
Public Function GetMergeRange(ByRef objSelection As Range, Optional ByVal enmMergeType As EMergeType = E_MTBOTH) As Range
    Dim objRange   As Range
    Dim objCell    As Range
    
    '結合されたセルはUsedRange以外にはないので
    Set objRange = IntersectRange(objSelection, GetUsedRange())
    If objRange Is Nothing Then
        Exit Function
    End If
    
On Error GoTo ErrHandle
'    If objRange.Count > 100000 Then
'        Call Err.Raise(C_CheckErrMsg, , "選択されたセルが多すぎます")
'    End If
    
    'セルの数だけループ
    For Each objCell In objRange
        With objCell.MergeArea
            '結合セルか？
            If .Count > 1 Then
                '左上のセルか
                If .Row = objCell.Row And .Column = objCell.Column Then
                    Select Case enmMergeType
                    Case E_MTBOTH
                        Set GetMergeRange = UnionRange(GetMergeRange, objCell)
                    Case E_MTROW
                        If .Rows.Count > 1 Then
                            Set GetMergeRange = UnionRange(GetMergeRange, objCell)
                        End If
                    Case E_MTCOL
                        If .Columns.Count > 1 Then
                            Set GetMergeRange = UnionRange(GetMergeRange, objCell)
                        End If
                    End Select
                End If
            End If
        End With
    Next
Exit Function
ErrHandle:
    Call Err.Raise(C_CheckErrMsg, , "選択されたセルが多すぎます")
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
'[ 関数名 ]　GetMergeAddress
'[ 概  要 ]　結合セル1つだけの時、左上のアドレスしか返らないので、全体を返す
'[ 引  数 ]　対象アドレス
'[ 戻り値 ]　アドレス
'*****************************************************************************
Public Function GetMergeAddress(ByVal strAddress As String) As String
    GetMergeAddress = strAddress
    With Range(strAddress)
        If .Rows.Count = 1 And .Columns.Count = 1 Then
            With .MergeArea
                If .Count > 1 Then
                    GetMergeAddress = .Address(0, 0)
                End If
            End With
        End If
    End With
End Function

'*****************************************************************************
'[ 関数名 ]　StrConvert
'[ 概  要 ]　文字種の変換を行う
'[ 引  数 ]　変換前の文字列、変換種類
'[ 戻り値 ]　変更後の文字列
'*****************************************************************************
Public Function StrConvert(ByVal strText As String, ByVal strCommand As String) As String
    StrConvert = strText
    Select Case strCommand
    Case "UpperCase"  '大文字に変換
        StrConvert = StrConv(StrConvert, vbUpperCase)
    Case "LowerCase"  '小文字に変換
        StrConvert = StrConv(StrConvert, vbLowerCase)
    Case "ProperCase" '先頭のみ大文字に変換
        StrConvert = StrConv(StrConvert, vbProperCase)
    Case "Hiragana" 'ひらがなに変換
        StrConvert = StrConv(StrConvert, vbHiragana)
    Case "Katakana" 'カタカナに変換
        StrConvert = StrConv(StrConvert, vbKatakana)
    Case "Wide"     '全角に変換
        StrConvert = Replace(StrConvert, """", Chr(&H8168))
        StrConvert = Replace(StrConvert, "'", "’")
        StrConvert = Replace(StrConvert, "\", "￥")
        StrConvert = StrConv(StrConvert, vbWide)
    Case "Narrow"   '半角に変換
        StrConvert = Replace(StrConvert, "～", Chr(1) & "~")
        StrConvert = StrConv(StrConvert, vbNarrow)
        StrConvert = Replace(StrConvert, Chr(1) & "~", "～")
    Case "NarrowExceptKana" 'カタカナ以外半角に変換
        StrConvert = Replace(StrConvert, "～", Chr(1) & "~")
        StrConvert = StrConvNarrowExceptKana(StrConvert)
        StrConvert = Replace(StrConvert, Chr(1) & "~", "～")
    Case "WideOnlyKana" 'カタカナのみ全角に変換
        StrConvert = StrConvWideOnlyKana(StrConvert)
    Case "Trim" '前後の空白を削除
        StrConvert = Trim(StrConvert)
    Case "RTrim" '末尾の空白を削除
        StrConvert = Rtrim(StrConvert)
    End Select
End Function

'*****************************************************************************
'[ 関数名 ]　StrConvNarrowExceptKana
'[ 概  要 ]　カタカナ以外半角に変換
'[ 引  数 ]　変換前の文字列
'[ 戻り値 ]　変換後の文字列
'*****************************************************************************
Private Function StrConvNarrowExceptKana(ByVal strText As String) As String
    Dim i       As Long
    Dim strChar As String
    
    '文字数分ループ
    For i = 1 To Len(strText)
        strChar = Mid$(strText, i, 1)
        Select Case AscW(strChar)
        Case AscW("ア") To AscW("ン"), AscW("ァ"), AscW("ヴ"), _
             AscW("ー"), AscW("、"), AscW("。")
            StrConvNarrowExceptKana = StrConvNarrowExceptKana & strChar
        Case Else
            StrConvNarrowExceptKana = StrConvNarrowExceptKana & StrConv(strChar, vbNarrow)
        End Select
    Next
End Function

'*****************************************************************************
'[ 関数名 ]　StrConvWideOnlyKana
'[ 概  要 ]　カタカナのみ全角に変換
'[ 引  数 ]　変換前の文字列
'[ 戻り値 ]　変換後の文字列
'*****************************************************************************
Private Function StrConvWideOnlyKana(ByVal strText As String) As String
    Dim i           As Long
    Dim strChar     As String
    Dim strWideChar As String
    
    '先頭が(半)濁点の時の対応として、先頭に適当な文字を追加しておく
    strText = "?" & strText
    
    '文字数分後方からループ　※上記で設定した先頭文字は対象外
    For i = Len(strText) To 2 Step -1
        strChar = Mid$(strText, i, 1)
        Select Case AscW(strChar)
        Case AscW("ｱ") To AscW("ﾝ"), AscW("ｧ") To AscW("ｯ"), AscW("ｦ"), _
             AscW("ｰ"), AscW("､"), AscW("｡")
            StrConvWideOnlyKana = StrConv(strChar, vbWide) & StrConvWideOnlyKana
        Case AscW("ﾞ"), AscW("ﾟ")
            strWideChar = StrConv(Mid$(strText, i - 1, 2), vbWide)
            If Len(strWideChar) = 1 Then
                '例：ｶﾞ → ガ
                StrConvWideOnlyKana = strWideChar & StrConvWideOnlyKana
                i = i - 1
            Else
                '例：ﾞ(半角) → ゛(全角)
                StrConvWideOnlyKana = StrConv(strChar, vbWide) & StrConvWideOnlyKana
            End If
        Case Else
            StrConvWideOnlyKana = strChar & StrConvWideOnlyKana
        End Select
    Next
End Function

'*****************************************************************************
'[概要] SortArray配列をソートする
'[引数] Sort対象配列
'[戻値] なし
'*****************************************************************************
Public Sub SortArray(ByRef SortArray() As TSortArray)
    'バブルソート
    Dim i As Long
    Dim j As Long
    Dim Swap As TSortArray
    For i = UBound(SortArray) To 1 Step -1
        For j = 1 To i - 1
            If CompareValue(SortArray(j), SortArray(j + 1)) Then
                Swap = SortArray(j)
                SortArray(j) = SortArray(j + 1)
                SortArray(j + 1) = Swap
            End If
        Next j
    Next i
    
    'スタック領域不足が起きる
'    Call QuickSort(SortArray, 1, UBound(SortArray))
End Sub

'Private Sub QuickSort(ByRef SortArray() As TSortArray, ByVal lngMin As Long, ByVal lngMax As Long)
'    Dim i As Long
'    Dim j As Long
'    Dim Base As TSortArray
'    Dim Swap As TSortArray
'    Base = SortArray(Int((lngMin + lngMax) / 2))
'    i = lngMin
'    j = lngMax
'    Do
'        Do While CompareValue(SortArray(i), Base)
'            i = i + 1
'        Loop
'        Do While Not CompareValue(SortArray(j), Base)
'            j = j - 1
'        Loop
'        If i >= j Then Exit Do
'            Swap = SortArray(i)
'            SortArray(i) = SortArray(j)
'            SortArray(j) = Swap
'        i = i + 1
'        j = j - 1
'    Loop
'    If (lngMin < i - 1) Then
'        Call QuickSort(SortArray, lngMin, i - 1)
'    End If
'    If (lngMax > j + 1) Then
'        Call QuickSort(SortArray, j + 1, lngMax)
'    End If
'End Sub

'*****************************************************************************
'[概要] 大小比較を行う
'[引数] 比較対象
'[戻値] True: SortArray1 > SortArray2
'*****************************************************************************
Public Function CompareValue(ByRef SortArray1 As TSortArray, ByRef SortArray2 As TSortArray) As Boolean
    If SortArray1.Key1 = SortArray2.Key1 Then
        CompareValue = (SortArray1.Key2 > SortArray2.Key2)
    Else
        CompareValue = (SortArray1.Key1 > SortArray2.Key1)
    End If
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
'[概要] Rangeが(結合された)単一のセルかどうか
'[引数] 判定するRange
'[戻値] True:単一のセル、False:複数のセル
'*****************************************************************************
Public Function IsOnlyCell(ByRef objRange As Range) As Boolean
    If objRange.Areas.Count > 1 Then
        Exit Function
    End If
    If objRange.Count = 1 Then
        IsOnlyCell = True
        Exit Function
    End If
    
    IsOnlyCell = (objRange.Address = objRange(1, 1).MergeArea.Address)
End Function

'*****************************************************************************
'[ 関数名 ]　GetUsedRange
'[ 概  要 ]　使用されている領域を取得する
'[ 引  数 ]　判定するRange
'[ 戻り値 ]　使用されている領域
'*****************************************************************************
Public Function GetUsedRange(Optional ByRef objSheet As Worksheet = Nothing) As Range
    On Error Resume Next
    If objSheet Is Nothing Then
        Set GetUsedRange = Range(Cells(1, 1), Cells.SpecialCells(xlCellTypeLastCell))
    Else
        With objSheet
            Set GetUsedRange = .Range(.Cells(1, 1), .Cells.SpecialCells(xlCellTypeLastCell))
        End With
    End If
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
On Error GoTo ErrHandle
    Dim objCb As New DataObject
    Call objCb.Clear
    Call objCb.SetText("")
    Call objCb.PutInClipboard
ErrHandle:
    Set objCb = Nothing
End Sub

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
        Call .Delete
    End With

    objSheet.StandardWidth = 8

    '最後のセルを修正する
    Call objSheet.Cells.Parent.UsedRange
End Sub

'*****************************************************************************
'[ 関数名 ]　SetIMEOff
'[ 概  要 ]　ＩＭＥをオフにする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetIMEOff()
On Error GoTo ErrHandle
    Dim hIME As LongPtr
    hIME = ImmGetContext(Application.hwnd)
    Call ImmSetOpenStatus(hIME, 0)
ErrHandle:
    If hIME <> 0 Then
        Call ImmReleaseContext(Application.hwnd, hIME)
    End If
End Sub

'*****************************************************************************
'[ 概  要 ]　DPIの変換率を取得する ※72(ExcelのデフォルトのDPI)/画面のDPI
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Function DPIRatio() As Double
    DPIRatio = 72 / GetDPI()
End Function

'*****************************************************************************
'[ 概  要 ]　DPIを取得する
'[ 引  数 ]　なし
'[ 戻り値 ]　DPI ※標準は96
'*****************************************************************************
Public Function GetDPI() As Long
    Dim DC As LongPtr
    DC = GetDC(0)
    GetDPI = GetDeviceCaps(DC, LOGPIXELSX)
    Call ReleaseDC(0, DC)
End Function

'*****************************************************************************
'[概要] コピー対象のRangeを取得する
'[引数] なし
'[戻値] コピー対象のRange
'*****************************************************************************
Public Function GetCopyRange() As Range
    If OpenClipboard(0) = 0 Then Exit Function
    Dim hMem As LongPtr
    hMem = GetClipboardData(RegisterClipboardFormat("Link"))
    If hMem = 0 Then
        Call CloseClipboard
        Exit Function
    End If
     
On Error GoTo ErrHandle
    Dim Size As Long
    Dim p As LongPtr
    Size = GlobalSize(hMem)
    p = GlobalLock(hMem)
    ReDim Data(1 To Size) As Byte
    Call CopyMemory(Data(1), ByVal p, Size)
    Call GlobalUnlock(hMem)
    Call CloseClipboard
    hMem = 0
    
    Dim strData As String
    Dim i As Long
    For i = 1 To Size
        If Data(i) = 0 Then
            Data(i) = Asc("/") 'シート名にもファイル名にも使えない文字
        End If
    Next i
    strData = StrConv(Data, vbUnicode)
    
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = False
    
    Dim strAddress As String
    Dim strFileAndSheet As String
    '例：Excel/Z:\[Book1.xlsx]Sheet1/R1C1:R2C2//
    objRegExp.Pattern = "Excel\/(.+)\/(R[0-9]+C[0-9]+(:R[0-9]+C[0-9]+)?)\/\/$"
    If Not objRegExp.Test(strData) Then Exit Function
    With objRegExp.Execute(strData)(0)
        '例：Z:\[Book1.xlsx]Sheet1
        strFileAndSheet = .SubMatches(0)
        '例：R1C1:R2C2
        strAddress = .SubMatches(1)
    End With
    
    Dim strBook  As String
    Dim strSheet As String
    '例：Z:\[Book1.xlsx]Sheet1
    objRegExp.Pattern = "\[(.+)\](.+)$"
    If objRegExp.Test(strFileAndSheet) Then
        With objRegExp.Execute(strFileAndSheet)(0)
            '例：Book1.xlsx
            strBook = .SubMatches(0)
            '例：Sheet1
            strSheet = .SubMatches(1)
        End With
    Else
        If vbYes = MsgBox("コピー元のファイルパスとシート名の合計の文字列が長すぎるためコピー元のセル範囲の取得ができませんでした" & vbLf & "ファイルパスとシート名を入力しますか？", vbQuestion & vbYesNo) Then
            strBook = InputBox("コピー元のファイル名を入力してください。" & vbLf & "貼り付け先と同じファイルの場合は省略可")
            strSheet = InputBox("コピー元のシート名を入力してください。" & vbLf & "貼り付け先と同じシートの場合は省略可")
        Else
            Exit Function
        End If
    End If
    
    Dim strRange As String
    strRange = Application.ConvertFormula(strAddress, xlR1C1, xlA1)
    If strBook = "" Then
        If strSheet = "" Then
            Set GetCopyRange = Range(strRange)
        Else
            Set GetCopyRange = Worksheets(strSheet).Range(strRange)
        End If
    Else
        Set GetCopyRange = Workbooks(strBook).Worksheets(strSheet).Range(strRange)
    End If
    
    If IsOnlyCell(GetCopyRange) Then
        Set GetCopyRange = GetCopyRange.MergeArea
    End If
    
    Application.CutCopyMode = False
    Exit Function
ErrHandle:
    If hMem <> 0 Then Call CloseClipboard
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'*****************************************************************************
'[概要] 拡張子の取得
'[引数] ファイルパス
'[戻値] 拡張子(大文字)
'*****************************************************************************
Public Function GetFileExtension(ByVal strFilename As String) As String
    With CreateObject("Scripting.FileSystemObject")
        GetFileExtension = UCase(.GetExtensionName(strFilename))
    End With
End Function

'*****************************************************************************
'[概要] Undoボタンの情報を取得する
'[引数] なし
'[戻値] UndoボタンのTooltipText
'*****************************************************************************
Public Function GetUndoStr() As String
    With CommandBars.FindControl(, 128) 'Undoボタン
        If .Enabled Then
            If .ListCount = 1 Then
                'Undoが1種類の時のUndoコマンド
                GetUndoStr = Trim(.List(1))
            End If
        End If
    End With
End Function

'*****************************************************************************
'[概要] 255字以上のアドレスでも取得できるようにする
'[引数] Range
'[戻値] 255字以上に対応したアドレス
'*****************************************************************************
Public Function GetAddress(ByRef objAreas As Range) As String
    Dim objRange As Range
    For Each objRange In objAreas.Areas
        If GetAddress = "" Then
            GetAddress = objRange.Address(0, 0)
        Else
            GetAddress = GetAddress & "," & objRange.Address(0, 0)
        End If
    Next
End Function

'*****************************************************************************
'[概要] 255字以上に対応したアドレスのRangeを取得
'[引数] 255字以上に対応したアドレス
'[戻値] Range
'*****************************************************************************
Public Function GetRange(ByVal strAddress As String) As Range
    Const MAXLEN = 250
    Dim strText As String
    Dim i As Long
    
    While Len(strAddress) > 0
        If Len(strAddress) >= MAXLEN Then
            For i = MAXLEN To 1 Step -1
                If Mid(strAddress, i, 1) = "," Then
                    strText = Left(strAddress, i - 1)
                    Set GetRange = UnionRange(GetRange, Range(strText))
                    strAddress = Mid(strAddress, i + 1)
                    Exit For
                End If
            Next
        Else
            Set GetRange = UnionRange(GetRange, Range(strAddress))
            strAddress = ""
        End If
    Wend
End Function

'*****************************************************************************
'[概要] オブジェクトの選択を表示画面を表示する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub ShowSelectionPane()
    On Error Resume Next
    If Not CommandBars.GetPressedMso("SelectionPane") Then
        Call CommandBars.ExecuteMso("SelectionPane")
    End If
End Sub


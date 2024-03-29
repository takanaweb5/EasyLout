VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUndoObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EUndoType  'アンドゥタイプ
    E_None        '初期値
    E_ColSize     '列幅変更
    E_ColSize2    '列幅変更
    E_ColHide     '列非表示
    E_ColReShow   '列再表示
    E_RowSize     '行高変更
    E_RowSize2    '行高変更
    E_RowHide     '行非表示
    E_RowReShow   '行再表示
    E_CellBorder  'セルの境界移動
    E_MoveCell    '領域の移動
    E_CopyCell    '領域のコピー
    E_PasteValue  '値の貼り付け
    E_ApplyFormat '書式の反映
    E_MergeCell   'セルの結合
    E_EraseCol    '列の消去
    E_EraseRow    '行の消去
    E_SplitCol    '行の分割
    E_SplitRow    '列の分割
    E_CellToText  'セルをテキストボックスに変換
    E_TextToCell  'テキストボックスをセルに変換
    E_CommentToRule 'コメントを入力規則に変換
    E_CommentToText 'コメントをテキストボックスに変換
    E_TextToComment 'テキストボックスをコメントに変換
    E_RuleToComment '入力規則をコメントに変換
    E_ShapeSize   '図形のサイズ変更
    E_ShapeSize2  '図形のサイズ変更
    E_CellValue   'セルの値の変更
    E_FillRange   '塗りつぶし(セル)
    E_FillShape   '塗りつぶし(図形)
    E_FilterERR   'フィルタ設定時はエラー
    E_UndoERR     '例外発生時はエラー
End Enum

Private Type TShapes  'Shape(Comment以外)
    Type     As Long
    ID       As Long
    Name     As String
    Top      As Double
    Height   As Double
    Left     As Double
    Width    As Double
    Rotation As Long
    Color    As Long
    Recover  As Boolean
End Type

Private Type TComments  'Comment
    Name     As String
    Address  As String
    Top      As Double
    Height   As Double
    Left     As Double
    Width    As Double
    Recover  As Boolean
End Type

Private Type TRangeSize  '幅か高さ
    Address  As String
    Size     As Double 'Width or Height
End Type

Private Type TUndoInfo  'Undo情報
    Worksheet      As Worksheet
    UndoSheet      As Worksheet
    UndoType       As EUndoType
    Shapes()       As TShapes
    ShapeCount     As Long
    Comments()     As TComments
    CommentCount   As Long
    RangeSize()    As TRangeSize
    RangeCount     As Long
    Address        As String
    SelectCell     As String
    SelectShapes() As Long
    SpiltCount     As Long
End Type

Private strClipText  As String  'クリップボードの中身(テキスト形式)
Private udtUndoInfo  As TUndoInfo

'*****************************************************************************
'[イベント]　Class_Terminate
'[ 概  要 ]　デストラクタ
'*****************************************************************************
Private Sub Class_Terminate()
    Call ClearBook
End Sub

'*****************************************************************************
'[概要] Workbookをクリアする
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ClearBook()
    Call DeleteSheet(ThisWorkbook.Worksheets("Undo"))

    '名前オブジェクトをすべて削除する
    Call DeleteNames(ThisWorkbook)
    'スタイルをすべて削除する
    Call DeleteStyles(ThisWorkbook)
End Sub

'*****************************************************************************
'[プロパティ]　UndoType
'[ 概  要 ]　タイプ
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get UndoType() As EUndoType
    UndoType = udtUndoInfo.UndoType
End Property

'*****************************************************************************
'[ 関数名 ]　SaveUndoInfo
'[ 概  要 ]　Undo情報を保存する
'[ 引  数 ]　enmType:Undoタイプ、objObject:復元対象の選択されたオブジェクト
'            varInfo:付加情報
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SaveUndoInfo(ByVal enmType As EUndoType, ByRef vSelection As Variant, Optional ByRef varInfo As Variant = Nothing)
On Error GoTo ErrHandle
    Dim objRange As Range
    Dim strAddress As String
    Dim objObject As Object
    
    If TypeOf vSelection Is Range Then
        Set objRange = vSelection
        strAddress = GetAddress(objRange)
    Else
        If VarType(vSelection) = vbString Then
            strAddress = vSelection
        Else
            Set objObject = vSelection
        End If
    End If
    
    'クリップボードの中身を保存
    Call SaveClipbordText
    
    If udtUndoInfo.UndoType = enmType Then
        Select Case enmType
        Case E_ColSize2, E_RowSize2, E_CellBorder
            If IsBeat(strAddress) = True Then
                Exit Sub
            End If
        Case E_ShapeSize2
            If IsBeat(objObject) = True Then
                Exit Sub
            End If
        End Select
    End If
    
    '最後のセルを修正する
    Call ActiveSheet.UsedRange  'Undo出来なくなります
    
    'Workbookをクリアする
    Call ClearBook
    
    udtUndoInfo.UndoType = enmType
    Set udtUndoInfo.Worksheet = ActiveSheet
    Set udtUndoInfo.UndoSheet = ThisWorkbook.Worksheets("Undo")
    
    Select Case (enmType)
    Case E_SplitCol, E_SplitRow
        udtUndoInfo.SelectCell = strAddress
        udtUndoInfo.SpiltCount = varInfo
        'ワークシ−トを保存
        Call BackupSheet(ActiveSheet)
        '図形の位置･サイズを保存
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
        'コメントの位置･サイズを保存
        Call BackupComment
        If enmType = E_SplitCol Then
            '列の幅を保存
            Call BackupRangeSize(GetSameWidthAddresses(objRange))
        Else
            '列の高さを保存
            Call BackupRangeSize(GetSameHeightAddresses(objRange))
        End If
    Case E_EraseCol, E_EraseRow
        udtUndoInfo.SelectCell = strAddress
        'ワークシ−トを保存
        Call BackupSheet(ActiveSheet)
        '図形の位置･サイズを保存
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
        'コメントの位置･サイズを保存
        Call BackupComment
        '列の幅(高さ)を保存
        Call BackupRangeSize(varInfo)
    Case E_ColSize, E_RowSize, _
         E_ColSize2, E_RowSize2
        udtUndoInfo.SelectCell = strAddress
        '列の幅を保存
        Call BackupRangeSize(varInfo)
        '図形の位置･サイズを保存
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
    Case E_ColHide, E_ColReShow, _
         E_RowHide, E_RowReShow
        udtUndoInfo.SelectCell = strAddress
        udtUndoInfo.Address = varInfo
        '図形の位置･サイズを保存
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
    Case E_CopyCell, E_MoveCell
        udtUndoInfo.SelectCell = strAddress
        '図形の位置･サイズを保存
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
        '変更前のセルを保存
        Call BackupSheet(ActiveSheet)
        'コメントの位置･サイズを保存
        Call BackupComment
    Case E_MergeCell, E_CellBorder, E_CommentToRule, E_RuleToComment
        udtUndoInfo.SelectCell = strAddress
        '変更前のセルを保存
        Call BackupSheet(ActiveSheet)
        'コメントの位置･サイズを保存
        Call BackupComment
    Case E_CellToText
        udtUndoInfo.SelectCell = strAddress
        '図形の名前を保存(追加されたテキストボックスを削除出来るように)
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
        '変更前のセルを保存
        Call BackupSheet(ActiveSheet)
    Case E_TextToCell, E_TextToComment
        '変更前のセルを保存
        Call BackupSheet(ActiveSheet)
        'テキストボックスを保存
        Call BackupTextbox(objObject)
    Case E_CommentToText
        udtUndoInfo.SelectCell = strAddress
        '図形の名前を保存(追加されたテキストボックスを削除出来るように)
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
        '変更前のセルを保存
        Call BackupSheet(ActiveSheet)
        'コメントの位置･サイズを保存
        Call BackupComment
    Case E_ShapeSize, E_ShapeSize2
        '図形の位置･サイズを保存
        Call BackupShapeSize(objObject)
    Case E_CellValue, E_PasteValue, E_ApplyFormat, E_FillRange
        udtUndoInfo.SelectCell = strAddress
        'セルを保存
        Call BackupSheet(ActiveSheet)
    Case E_FillShape
        '図形の塗りつぶし色を保存
        Call BackupShapeColor(objObject)
    Case E_FilterERR
    '何もしない
    End Select
Exit Sub
ErrHandle:
    udtUndoInfo.UndoType = E_UndoERR
End Sub

'*****************************************************************************
'[ 関数名 ]　IsBeat
'[ 概  要 ]　コマンドが連続で実行されたかどうか判定
'[ 引  数 ]　選択されたオブジェクト
'[ 戻り値 ]　True:連続で実行された時
'*****************************************************************************
Private Function IsBeat(ByRef vSelection As Variant) As Boolean
On Error GoTo ErrHandle
    Dim i As Long
    
    If GetUndoCount() <> 1 Then
        Exit Function
    End If
    
    If VarType(vSelection) = vbString Then
        If udtUndoInfo.SelectCell <> vSelection Then
            Exit Function
        End If
    ElseIf TypeOf vSelection Is ShapeRange Then
        If udtUndoInfo.ShapeCount <> vSelection.Count Then
            Exit Function
        End If
        '図形の数だけループ
        For i = 1 To vSelection.Count
            With udtUndoInfo
                If .SelectShapes(i) <> vSelection(i).ID Then
                    Exit Function
                End If
            End With
        Next
    End If
    
    IsBeat = True
    Exit Function
ErrHandle:
End Function

'*****************************************************************************
'[ 関数名 ]　GetUndoCount
'[ 概  要 ]　UndoボタンのUndo数を取得する
'[ 引  数 ]　なし
'[ 戻り値 ]　Undo数
'*****************************************************************************
Private Function GetUndoCount() As Long
    With CommandBars.FindControl(, 128) 'Undoボタン
        If .Enabled Then
            GetUndoCount = .ListCount
        End If
    End With
End Function

'*****************************************************************************
'[ 関数名 ]　SaveClipbordText
'[ 概  要 ]　クリップボードのテキスト形式を保存する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SaveClipbordText()
On Error GoTo ErrHandle
    Dim objCb As New DataObject
    strClipText = ""
    If Application.CutCopyMode = False Then
        strClipText = GetClipbordText()
    End If
ErrHandle:
End Sub

'*****************************************************************************
'[ 関数名 ]　RecoverClipbordText
'[ 概  要 ]　クリップボードのテキスト形式を元に戻す
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub RecoverClipbordText()
On Error GoTo ErrHandle
    Dim objCb As New DataObject
    If strClipText <> "" Then
        Call objCb.GetFromClipboard
        If objCb.GetFormat(1) Then
            If strClipText <> objCb.GetText Then
                Call SetClipbordText(strClipText)
            End If
        Else
            Call SetClipbordText(strClipText)
        End If
    End If
ErrHandle:
End Sub

'*****************************************************************************
'[ 関数名 ]　BackupTextbox
'[ 概  要 ]　変更前のテキストボックスを保存する
'[ 引  数 ]　objTextbox:テキストボックス
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub BackupTextbox(ByRef objTextbox As ShapeRange)
    Dim objShape As Shape
    
    With udtUndoInfo
        'Undoシートに変更前のテキストボックスを保存
        .ShapeCount = objTextbox.Count
        If .ShapeCount > 1 Then
            Set objShape = objTextbox.Group
        Else
            Set objShape = objTextbox(1)
        End If
        ReDim .Shapes(1 To 1)
        .Shapes(1).Top = objShape.Top
        .Shapes(1).Left = objShape.Left
        
        Call objShape.Copy
        Call .UndoSheet.Paste
        If .ShapeCount > 1 Then
            Call objShape.Ungroup
        End If
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　BackupSheet
'[ 概  要 ]　Undoシートに変更前のセルを保存
'[ 引  数 ]　保存対象のWorksheet
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub BackupSheet(ByRef objWorksheet As Worksheet)
On Error GoTo ErrHandle
    Dim objRange As Range
    Dim blnCopyObjectsWithCells  As Boolean
    
    'WorkSheetの標準スタイルを実行シートとあわせる
    Call SetPixelInfo 'Undoできなくなります
    
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells
    '図形をコピーの対象外にする
    Application.CopyObjectsWithCells = False
 
    With udtUndoInfo
        Call DeleteSheet(.UndoSheet)
        '重たくなるのはクリップボード監視のapiが原因
        Call objWorksheet.Cells.Copy(.UndoSheet.Cells)
    End With
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
Exit Sub
ErrHandle:
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub

'*****************************************************************************
'[ 関数名 ]　BackupRangeSize
'[ 概  要 ]　幅または高さの情報を保存する
'[ 引  数 ]　保存対象の列または行のアドレス(Collection)
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub BackupRangeSize(ByRef colSelection As Variant)
    Dim i     As Long
    With udtUndoInfo
        ReDim Preserve .RangeSize(1 To colSelection.Count)
        For i = 1 To colSelection.Count
            .RangeSize(i).Address = colSelection(i)
            Select Case (Me.UndoType)
            Case E_SplitCol, E_EraseCol, E_ColSize, E_ColSize2
                .RangeSize(i).Size = Range(colSelection(i)).Columns(1).ColumnWidth
            Case E_SplitRow, E_EraseRow, E_RowSize, E_RowSize2
                .RangeSize(i).Size = Range(colSelection(i)).Rows(1).RowHeight
            End Select
        Next i
        .RangeCount = colSelection.Count
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　BackupShapeSize
'[ 概  要 ]　図形のサイズを保存する
'[ 引  数 ]　保存対象の図形
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub BackupShapeSize(ByRef objSelection As ShapeRange)
    Dim i          As Long
    Dim j          As Long
   
    With udtUndoInfo
        If objSelection Is Nothing Then
            .ShapeCount = 0
            Exit Sub
        End If
        
        .ShapeCount = objSelection.Count
        If .ShapeCount = 0 Then
            Exit Sub
        End If
        
        ReDim .Shapes(1 To .ShapeCount)
        ReDim .SelectShapes(1 To .ShapeCount)
    End With
    
    '図形の数だけループ
    For j = 1 To objSelection.Count 'For each構文だとExcel2007で型違いとなる(たぶんバグ)
        With objSelection(j)
            If .Type <> msoComment Then
                i = i + 1
                udtUndoInfo.SelectShapes(i) = .ID
                udtUndoInfo.Shapes(i).ID = .ID
                udtUndoInfo.Shapes(i).Name = .Name
                udtUndoInfo.Shapes(i).Type = .Type
                udtUndoInfo.Shapes(i).Rotation = .Rotation
                udtUndoInfo.Shapes(i).Height = .Height
                udtUndoInfo.Shapes(i).Width = .Width
                udtUndoInfo.Shapes(i).Top = .Top
                udtUndoInfo.Shapes(i).Left = .Left
            End If
        End With
    Next
    udtUndoInfo.ShapeCount = i
End Sub

'*****************************************************************************
'[ 関数名 ]　BackupComment
'[ 概  要 ]　コメントの対象セルのアドレスとサイズを保存する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub BackupComment()
    Dim i            As Long
    Dim objWorksheet As Worksheet
    Set objWorksheet = udtUndoInfo.UndoSheet
    
    With udtUndoInfo
        .CommentCount = objWorksheet.Comments.Count
        If .CommentCount = 0 Then
            Exit Sub
        End If
        ReDim .Comments(1 To .CommentCount)
    End With
    
    'Undoシート上のコメント単位にループ
    For i = 1 To objWorksheet.Comments.Count
        udtUndoInfo.Comments(i).Address = objWorksheet.Comments(i).Parent.Address
        With udtUndoInfo.Worksheet.Range(objWorksheet.Comments(i).Parent.Address).Comment.Shape
            udtUndoInfo.Comments(i).Name = .Name
            udtUndoInfo.Comments(i).Height = .Height
            udtUndoInfo.Comments(i).Width = .Width
            udtUndoInfo.Comments(i).Top = .Top
            udtUndoInfo.Comments(i).Left = .Left
        End With
    Next i
End Sub

'*****************************************************************************
'[概要] 図形の塗りつぶし色を保存する
'[引数] 保存対象の図形
'[戻値] なし
'*****************************************************************************
Private Sub BackupShapeColor(ByRef objSelection As ShapeRange)
    Dim i As Long
    With udtUndoInfo
        .ShapeCount = objSelection.Count
        
        ReDim .Shapes(1 To .ShapeCount)
        ReDim .SelectShapes(1 To .ShapeCount)
    End With
    
    '図形の数だけループ
    For i = 1 To objSelection.Count 'For each構文だとExcel2007で型違いとなる(たぶんバグ)
        With objSelection(i)
            udtUndoInfo.SelectShapes(i) = .ID
            udtUndoInfo.Shapes(i).ID = .ID
            udtUndoInfo.Shapes(i).Name = .Name
            udtUndoInfo.Shapes(i).Type = .Type
            If .DrawingObject.Interior.ColorIndex = xlNone Then
                udtUndoInfo.Shapes(i).Color = -1
            Else
                udtUndoInfo.Shapes(i).Color = .DrawingObject.Interior.Color
            End If
        End With
    Next
End Sub

'*****************************************************************************
'[ 関数名 ]　ExecUndo
'[ 概  要 ]　Undoを実行する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ExecUndo()
    Dim lngNo    As Long
    
    Select Case Me.UndoType
    Case E_FilterERR, E_UndoERR
        Application.ScreenUpdating = True '残像が残らないようにするおまじない
        Select Case Me.UndoType
        Case E_FilterERR
            Call MsgBox("フィルタ使用中の操作はUndo出来ません", vbInformation, "かんたんレイアウト")
        Case E_UndoERR
            Call MsgBox("変更前の状態の保存に失敗しました" & vbCrLf & "Undo出来ません", vbExclamation, "かんたんレイアウト")
        End Select
        Exit Sub
    End Select
    
    'クリップボードの中身を保存
    Call SaveClipbordText
    
    'クリップボードのクリア
    Call ClearClipbord
    
    Select Case (Me.UndoType)
    Case E_ColSize, E_RowSize, _
         E_ColSize2, E_RowSize2
        '列の幅または行の高さを復元する
        Call RecoverRangeSize
        '図形のサイズを復元する
        Call RecoverShapeSize
    Case E_ColHide
        '列を再表示する
        GetRange(udtUndoInfo.Address).EntireColumn.Hidden = False
        '図形のサイズを復元する
        Call RecoverShapeSize
    Case E_RowHide
        '行を再表示する
        GetRange(udtUndoInfo.Address).EntireRow.Hidden = False
        '図形のサイズを復元する
        Call RecoverShapeSize
    Case E_ColReShow
        '列を非表示する
        GetRange(udtUndoInfo.Address).EntireColumn.Hidden = True
        '図形のサイズを復元する
        Call RecoverShapeSize
    Case E_RowReShow
        '行を非表示する
        GetRange(udtUndoInfo.Address).EntireRow.Hidden = True
        '図形のサイズを復元する
        Call RecoverShapeSize
    Case E_EraseCol, E_EraseRow
        Dim i As Long
        '名前オブジェクトがセル参照をしているときに復元できるように回復させる
        If Me.UndoType = E_EraseCol Then
            With udtUndoInfo
                lngNo = .Worksheet.Range(.SelectCell).Column
                '元のシートの削除された列の数だけ挿入
                For i = lngNo To lngNo + .Worksheet.Range(.SelectCell).Columns.Count - 1
                    Call .Worksheet.Columns(lngNo).Insert
                Next i
            End With
        Else
            With udtUndoInfo
                lngNo = .Worksheet.Range(.SelectCell).Row
                '元のシートの削除された列の数だけ挿入
                For i = lngNo To lngNo + .Worksheet.Range(.SelectCell).Rows.Count - 1
                    Call .Worksheet.Rows(lngNo).Insert
                Next i
            End With
        End If
        'ワークシートを復元する
        Call RecoverSheet
        'コメントを復元
        Call RecoverComments
        '列の幅または行の高さを復元する
        Call RecoverRangeSize
        '図形のサイズを復元する
        Call RecoverShapeSize
    Case E_SplitCol, E_SplitRow
        '名前オブジェクトがセル参照をしているときに復元できるように回復させる
        If Me.UndoType = E_SplitCol Then
            With udtUndoInfo
                lngNo = .Worksheet.Range(.SelectCell).Column
                '元のシートの挿入された列を削除
                Call .Worksheet.Range(.Worksheet.Columns(lngNo + 1), .Worksheet.Columns(lngNo + .SpiltCount - 1)).Delete
            End With
        Else
            With udtUndoInfo
                lngNo = .Worksheet.Range(.SelectCell).Row
                '元のシートの挿入された行を削除
                Call .Worksheet.Range(.Worksheet.Rows(lngNo + 1), .Worksheet.Rows(lngNo + .SpiltCount - 1)).Delete
            End With
        End If
        'セルを復元
        Call RecoverSheet
        'コメントを復元
        Call RecoverComments
        '列の幅または行の高さを復元する
        Call RecoverRangeSize
        '図形のサイズを復元する
        Call RecoverShapeSize
    Case E_MoveCell, E_CopyCell
        'セルを復元
        Call RecoverSheet
        'コメントを復元
        Call RecoverComments
        '図形のサイズを復元する
        Call RecoverShapeSize
        '追加された図形を削除
        Call DeleteNewShapes
    Case E_MergeCell, E_CellBorder, E_CommentToRule, E_RuleToComment
        'セルを復元
        Call RecoverSheet
        'コメントを復元
        Call RecoverComments
        If Me.UndoType = E_CommentToRule Then
            Call DeleteInputRules
        End If
    Case E_CellToText
        'セルを復元
        Call RecoverSheet
        '追加されたテキストボックスを削除
        Call DeleteNewShapes
    Case E_TextToCell, E_TextToComment
        'テキストボックスを復元する
        Call RecoverTextbox
        'セルを復元
        Call RecoverSheet
        'コメントを復元
        Call RecoverComments
    Case E_CommentToText
        'セルを復元
        Call RecoverSheet
        'コメントを復元
        Call RecoverComments
        '追加されたテキストボックスを削除
        Call DeleteNewShapes
    Case E_ShapeSize, E_ShapeSize2
        '図形のサイズを復元する
        Call RecoverShapeSize
    Case E_CellValue, E_PasteValue, E_ApplyFormat, E_FillRange
        'セルを復元
        Call RecoverSheet
    Case E_FillShape
        '図形の塗りつぶし色を復元する
        Call RecoverShapeColor
    End Select
    
    '復元対象になったセルやオブジェクトの選択
    With udtUndoInfo
        Call .Worksheet.Activate
        Select Case (Me.UndoType)
        Case E_ColSize, E_RowSize, E_ColSize2, E_RowSize2, _
             E_ColHide, E_RowHide, _
             E_ColReShow, E_RowReShow, _
             E_MergeCell, E_MoveCell, E_CopyCell, E_PasteValue, _
             E_CellToText, E_ApplyFormat, _
             E_SplitCol, E_EraseCol, _
             E_SplitRow, E_EraseRow, _
             E_CellValue, _
             E_CommentToText, _
             E_CommentToRule, _
             E_RuleToComment, _
             E_FillRange
            Call GetRange(.SelectCell).Select
        Case E_ShapeSize, E_ShapeSize2, E_FillShape
            Call GetShapeRangeFromID(.SelectShapes).Select
        Case E_TextToCell
        End Select
    End With

    'クリップボードの中身を復元
    Call RecoverClipbordText

    '最後のセルを修正する
    Call ActiveSheet.UsedRange  'Undo出来なくなります
    
    udtUndoInfo.UndoType = E_None
End Sub

'*****************************************************************************
'[ 関数名 ]　RecoverTextbox
'[ 概  要 ]　テキストボックスを復元する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub RecoverTextbox()
    Dim objTextbox   As Object
    
    With udtUndoInfo
        'テキストボックスを復元
        Set objTextbox = .UndoSheet.Shapes(1)
        Call objTextbox.Cut
        Call .Worksheet.Paste
    End With
    
'    With udtUndoInfo.Worksheet.Shapes
'        Set objTextbox = .Item(.Count)
'    End With

    Set objTextbox = Selection
    
    With udtUndoInfo
        'テキストボックスの位置を復元
        objTextbox.Top = .Shapes(1).Top
        objTextbox.Left = .Shapes(1).Left
    End With
    
    If udtUndoInfo.ShapeCount > 1 Then
        Call objTextbox.Ungroup
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　RecoverSheet
'[ 概  要 ]　保存対象のセルを復元する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub RecoverSheet()
    Dim objRange As Range
    Dim blnCopyObjectsWithCells  As Boolean
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells
    '図形をコピーの対象外にする
    Application.CopyObjectsWithCells = False
 
    With udtUndoInfo
        'Undoシートから変更前のセルに移動
'        Call .Worksheet.Cells.UnMerge
        Call .UndoSheet.Cells.Copy(.Worksheet.Cells) 'CutだとExcel2007で幅が復元されない
        Call .UndoSheet.Cells.Clear 'コメントオブジェクトを削除する必要がある
    End With
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
Exit Sub
ErrHandle:
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub

'*****************************************************************************
'[ 関数名 ]　DeleteNewShapes
'[ 概  要 ]　ワークシートの追加された図形を削除
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub DeleteNewShapes()
    Dim i As Long
    Dim objWorksheet As Worksheet
    Dim objShape     As Shape
    Dim blnHit       As Boolean
    
    Set objWorksheet = udtUndoInfo.Worksheet
    
    'コメント以外のオブジェクトで保存対象でなかったものは削除する
    For Each objShape In objWorksheet.Shapes
        If objShape.Type <> msoComment Then
            blnHit = False
            For i = 1 To udtUndoInfo.ShapeCount
                If objShape.ID = udtUndoInfo.Shapes(i).ID Then
                    blnHit = True
                    Exit For
                End If
            Next i
            If blnHit = False Then
                Call objShape.Delete
            End If
        End If
    Next objShape
End Sub

'*****************************************************************************
'[ 関数名 ]　RecoverRangeSize
'[ 概  要 ]　列の幅または行の高さを復元
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub RecoverRangeSize()
    Dim i            As Long
    Dim objWorksheet As Worksheet
    
    Set objWorksheet = udtUndoInfo.Worksheet
    
    For i = 1 To udtUndoInfo.RangeCount
        With udtUndoInfo.RangeSize(i)
            Select Case (Me.UndoType)
            Case E_SplitCol, E_EraseCol, E_ColSize, E_ColSize2
                objWorksheet.Range(.Address).ColumnWidth = .Size
            Case E_SplitRow, E_EraseRow, E_RowSize, E_RowSize2
                objWorksheet.Range(.Address).RowHeight = .Size
            End Select
        End With
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　RecoverShapeSize
'[ 概  要 ]　図形のサイズと位置を復元
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub RecoverShapeSize()
    Dim i As Long
    Dim objWorksheet As Worksheet
    Dim objShape     As Shape
    
    Set objWorksheet = udtUndoInfo.Worksheet
    
    For Each objShape In objWorksheet.Shapes
        For i = 1 To udtUndoInfo.ShapeCount
            With udtUndoInfo.Shapes(i)
                If .ID = objShape.ID Then
                    objShape.Height = .Height
                    objShape.Width = .Width
                    objShape.Top = .Top
                    objShape.Left = .Left
                    If .Top < 0 Then
                        '図形が回転して座標がマイナスになった時ゼロになるため補正する
                        Call objShape.IncrementTop(.Top)
                    End If
                    If .Left < 0 Then
                        '図形が回転して座標がマイナスになった時ゼロになるため補正する
                        Call objShape.IncrementLeft(.Left)
                    End If
                    
                    .Recover = True
                    Exit For
                End If
            End With
        Next
    Next

    For i = 1 To udtUndoInfo.ShapeCount
        With udtUndoInfo.Shapes(i)
            If .Recover = False Then
                Call MsgBox(.Name & "が復元されませんでした")
            End If
        End With
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　RecoverComments
'[ 概  要 ]　コメントのサイズと位置を復元
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub RecoverComments()
    Dim i            As Long
    Dim objWorksheet As Worksheet
    Dim objComment   As Comment
    
    Set objWorksheet = udtUndoInfo.Worksheet
    
    '保存されたコメントの数だけループ
    For i = 1 To udtUndoInfo.CommentCount
        For Each objComment In objWorksheet.Comments
            If udtUndoInfo.Comments(i).Address = objComment.Parent.Address Then
                udtUndoInfo.Comments(i).Recover = True
                Exit For
            End If
        Next objComment
        
        If udtUndoInfo.Comments(i).Recover = True Then
            With objComment.Shape
                .Height = udtUndoInfo.Comments(i).Height
                .Width = udtUndoInfo.Comments(i).Width
                .Top = udtUndoInfo.Comments(i).Top
                .Left = udtUndoInfo.Comments(i).Left
            End With
        End If
    Next i
    
    '復元もれのチェック(Debug用)
    For i = 1 To udtUndoInfo.CommentCount
        With udtUndoInfo.Comments(i)
            If .Recover = False Then
                Call MsgBox(.Name & "(" & .Address & ")" & "が復元されませんでした")
            End If
        End With
    Next i
End Sub

'*****************************************************************************
'[概要] 図形の塗りつぶし色を復元する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub RecoverShapeColor()
    Dim i As Long
    Dim objWorksheet As Worksheet
    Dim objShape     As Shape
    Set objWorksheet = udtUndoInfo.Worksheet
    For Each objShape In objWorksheet.Shapes
        For i = 1 To udtUndoInfo.ShapeCount
            With udtUndoInfo.Shapes(i)
                If .ID = objShape.ID Then
                    If .Color = -1 Then
                        objShape.DrawingObject.Interior.ColorIndex = xlNone
                    Else
                        objShape.DrawingObject.Interior.Color = .Color
                    End If
                    Exit For
                End If
            End With
        Next
    Next
End Sub

'*****************************************************************************
'[概要] 入力規則を削除する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub DeleteInputRules()
    With udtUndoInfo
        Call .Worksheet.Range(.SelectCell).Validation.Delete
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　SetOnUndo
'[ 概  要 ]　ApplicationオブジェクトのOnUndoイベントを設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetOnUndo()
    Select Case (Me.UndoType)
    Case E_ColSize, E_ColSize2, E_ColHide, E_ColReShow
        Call Application.OnUndo("幅の変更", MacroName("ExecUndo"))
    Case E_RowSize, E_RowSize2, E_RowHide, E_RowReShow
        Call Application.OnUndo("高さの変更", MacroName("ExecUndo"))
    Case E_ShapeSize, E_ShapeSize2
        Call Application.OnUndo("図形のサイズの変更", MacroName("ExecUndo"))
    Case E_CellBorder
        Call Application.OnUndo("セルの境界移動", MacroName("ExecUndo"))
    Case E_MergeCell
        Call Application.OnUndo("セルの結合", MacroName("ExecUndo"))
    Case E_MoveCell
        Call Application.OnUndo("領域の移動", MacroName("ExecUndo"))
    Case E_CopyCell
        Call Application.OnUndo("領域のコピー", MacroName("ExecUndo"))
    Case E_PasteValue
        Call Application.OnUndo("値の貼り付け", MacroName("ExecUndo"))
    Case E_ApplyFormat
        Call Application.OnUndo("書式の反映", MacroName("ExecUndo"))
    Case E_SplitCol
        Call Application.OnUndo("列の分割", MacroName("ExecUndo"))
    Case E_EraseCol
        Call Application.OnUndo("列の消去", MacroName("ExecUndo"))
    Case E_SplitRow
        Call Application.OnUndo("行の分割", MacroName("ExecUndo"))
    Case E_EraseRow
        Call Application.OnUndo("行の消去", MacroName("ExecUndo"))
    Case E_CellToText
        Call Application.OnUndo("セルをテキストボックスに変換", MacroName("ExecUndo"))
    Case E_TextToCell
        Call Application.OnUndo("テキストボックスをセルに変換", MacroName("ExecUndo"))
    Case E_CommentToText
        Call Application.OnUndo("コメントをテキストボックスに変換", MacroName("ExecUndo"))
    Case E_TextToComment
        Call Application.OnUndo("テキストボックスをコメントに戻す", MacroName("ExecUndo"))
    Case E_CommentToRule
        Call Application.OnUndo("コメントを入力規則に変換", MacroName("ExecUndo"))
    Case E_RuleToComment
        Call Application.OnUndo("入力規則をコメントに変換", MacroName("ExecUndo"))
    Case E_CellValue
        Call Application.OnUndo("セルの値を変更", MacroName("ExecUndo"))
    Case E_FillRange, E_FillShape
        Call Application.OnUndo("塗りつぶし", MacroName("ExecUndo"))
    Case E_FilterERR
        Call Application.OnUndo("フィルタ使用中の操作はUndo出来ません", MacroName("ExecUndo"))
    Case E_UndoERR
        Call Application.OnUndo("変更前の状態の保存に失敗しました", MacroName("ExecUndo"))
    End Select
    
    'クリップボードの中身を復元
    Call RecoverClipbordText
End Sub

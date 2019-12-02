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
    E_ColSize     '列幅変更
    E_ColSize2    '列幅変更
    E_RowSize     '行高変更
    E_RowSize2    '行高変更
    E_CellBorder  'セルの境界移動
    E_MoveCell    '領域の移動
    E_CopyCell    '領域のコピー
    E_PasteValue  '値の貼り付け
    E_MergeCell   'セルの結合
    E_EraseCol    '列の消去
    E_EraseRow    '行の消去
    E_SplitCol    '行の分割
    E_SplitRow    '列の分割
    E_CellToText  'セルをテキストボックスに変換
    E_TextToCell  'テキストボックスをセルに変換
    E_ShapeSize   '図形のサイズ変更
    E_ShapeSize2  '図形のサイズ変更
    E_CellValue   'セルの値の変更
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
    SelectShapes() As Variant
    SpiltCount     As Long
End Type

Private strClipText   As String  'クリップボードの中身(テキスト形式)
Private bytNo         As Byte '1 or 2
Private udtUndoInfo(1 To 2) As TUndoInfo

'*****************************************************************************
'[イベント]　Class_Terminate
'[ 概  要 ]　デストラクタ
'*****************************************************************************
Private Sub Class_Terminate()
    Call DeleteSheet(ThisWorkbook.Worksheets("Undo1"))
    Call DeleteSheet(ThisWorkbook.Worksheets("Undo2"))
    
    Call ClearBook(ThisWorkbook)
End Sub

'*****************************************************************************
'[プロパティ]　UndoType
'[ 概  要 ]　タイプ
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get UndoType() As EUndoType
    UndoType = udtUndoInfo(bytNo).UndoType
End Property

'*****************************************************************************
'[ 関数名 ]　SaveUndoInfo
'[ 概  要 ]　Undo情報を保存する
'[ 引  数 ]　enmType:Undoタイプ、objObject:復元対象の選択されたオブジェクト
'            varInfo:付加情報
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SaveUndoInfo(ByVal enmType As EUndoType, ByRef objObject As Object, Optional ByRef varInfo As Variant = Nothing)
On Error GoTo ErrHandle
    
    'クリップボードの中身を保存
    Call SaveClipbordText
    
    If bytNo <> 0 Then
        If udtUndoInfo(bytNo).UndoType = enmType Then
            Select Case enmType
            Case E_ColSize2, E_RowSize2, E_ShapeSize2, E_CellBorder
                If IsBeat(objObject) = True Then
                    Exit Sub
                End If
            End Select
        End If
    End If
    
    If bytNo = 1 Then
        bytNo = 2
    Else
        bytNo = 1
    End If
    
    '最後のセルを修正する
    Call ActiveSheet.UsedRange  'Undo出来なくなります
    
    'Workbookをクリアする
    Call ClearBook(ThisWorkbook)
    
    udtUndoInfo(bytNo).UndoType = enmType
    Set udtUndoInfo(bytNo).Worksheet = ActiveSheet
    Set udtUndoInfo(bytNo).UndoSheet = ThisWorkbook.Worksheets("Undo" & bytNo)
    
    Select Case (enmType)
    Case E_SplitCol, E_SplitRow
        udtUndoInfo(bytNo).SelectCell = objObject.Address
        udtUndoInfo(bytNo).SpiltCount = varInfo
        'ワークシ－トを保存
        Call BackupSheet(ActiveSheet)
        '図形の位置･サイズを保存
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
        'コメントの位置･サイズを保存
        Call BackupComment
        If enmType = E_SplitCol Then
            '列の幅を保存
            Call BackupRangeSize(GetSameWidthAddresses(objObject))
        Else
            '列の高さを保存
            Call BackupRangeSize(GetSameHeightAddresses(objObject))
        End If
    Case E_EraseCol, E_EraseRow
        udtUndoInfo(bytNo).SelectCell = objObject.Address
        'ワークシ－トを保存
        Call BackupSheet(ActiveSheet)
        '図形の位置･サイズを保存
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
        'コメントの位置･サイズを保存
        Call BackupComment
        '列の幅(高さ)を保存
        Call BackupRangeSize(varInfo)
    Case E_ColSize, E_RowSize, _
         E_ColSize2, E_RowSize2
        udtUndoInfo(bytNo).SelectCell = objObject.Address
        '列の幅を保存
        Call BackupRangeSize(varInfo)
        '図形の位置･サイズを保存
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
    Case E_CopyCell, E_MoveCell
        udtUndoInfo(bytNo).SelectCell = objObject.Address
        '図形の位置･サイズを保存
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
        '変更前のセルを保存
        Call BackupSheet(ActiveSheet)
        'コメントの位置･サイズを保存
        Call BackupComment
    Case E_MergeCell, E_CellBorder
        udtUndoInfo(bytNo).SelectCell = objObject.Address
        '変更前のセルを保存
        Call BackupSheet(ActiveSheet)
        'コメントの位置･サイズを保存
        Call BackupComment
    Case E_CellToText
        udtUndoInfo(bytNo).SelectCell = objObject.Address
        '図形の名前を保存(追加されたテキストボックスを削除出来るように)
        Call BackupShapeSize(GetSheeetShapeRange(ActiveSheet))
        '変更前のセルを保存
        Call BackupSheet(ActiveSheet)
    Case E_TextToCell
        '変更前のセルを保存
        Call BackupSheet(ActiveSheet)
        'テキストボックスを保存
        Call BackupTextbox(objObject)
    Case E_ShapeSize, E_ShapeSize2
        '図形の位置･サイズを保存
        Call BackupShapeSize(objObject)
    Case E_CellValue, E_PasteValue
        udtUndoInfo(bytNo).SelectCell = objObject.Address
        'セルを保存
        Call BackupSheet(ActiveSheet)
    Case E_FilterERR
    '何もしない
    End Select
Exit Sub
ErrHandle:
    bytNo = 1
    udtUndoInfo(1).UndoType = E_UndoERR
End Sub

'*****************************************************************************
'[ 関数名 ]　IsBeat
'[ 概  要 ]　コマンドが連続で実行されたかどうか判定
'[ 引  数 ]　選択されたオブジェクト
'[ 戻り値 ]　True:連続で実行された時
'*****************************************************************************
Private Function IsBeat(ByRef objObject As Object) As Boolean
On Error GoTo ErrHandle
    Dim i As Long
    
    If GetUndoCount() <> 1 Then
        Exit Function
    End If
    
    If TypeOf objObject Is Range Then
        If udtUndoInfo(bytNo).SelectCell <> objObject.Address Then
            Exit Function
        End If
    ElseIf TypeOf objObject Is ShapeRange Then
        If udtUndoInfo(bytNo).ShapeCount <> objObject.Count Then
            Exit Function
        End If
        '図形の数だけループ
        For i = 1 To objObject.Count
            With udtUndoInfo(bytNo)
                If .SelectShapes(i) <> objObject(i).ID Then
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
    'Undo情報の取得のため、Undoボタン作成
    With CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
        With .Controls.Add(, 128) 'Undoボタン
            If .Enabled = True Then
                GetUndoCount = .ListCount
            Else
                GetUndoCount = 0
            End If
        End With
        
        Call .Delete
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
    
    With udtUndoInfo(bytNo)
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
 
    With udtUndoInfo(bytNo)
        Call DeleteSheet(.UndoSheet)
        Call objWorksheet.Cells.Copy(.UndoSheet.Cells)
    End With
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
Exit Sub
ErrHandle:
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
    Call Err.Raise(Err.NUMBER, Err.Source, Err.Description)
End Sub

'*****************************************************************************
'[ 関数名 ]　BackupRangeSize
'[ 概  要 ]　幅または高さの情報を保存する
'[ 引  数 ]　保存対象の列または行のアドレス(Collection)
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub BackupRangeSize(ByRef colSelection As Variant)
    Dim i     As Long
    With udtUndoInfo(bytNo)
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
   
    With udtUndoInfo(bytNo)
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
                udtUndoInfo(bytNo).SelectShapes(i) = .ID
                udtUndoInfo(bytNo).Shapes(i).ID = .ID
                udtUndoInfo(bytNo).Shapes(i).Name = .Name
                udtUndoInfo(bytNo).Shapes(i).Type = .Type
                udtUndoInfo(bytNo).Shapes(i).Rotation = .Rotation
                udtUndoInfo(bytNo).Shapes(i).Height = .Height
                udtUndoInfo(bytNo).Shapes(i).Width = .Width
                udtUndoInfo(bytNo).Shapes(i).Top = .Top
                udtUndoInfo(bytNo).Shapes(i).Left = .Left
            End If
        End With
    Next
    udtUndoInfo(bytNo).ShapeCount = i
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
    Set objWorksheet = udtUndoInfo(bytNo).UndoSheet
    
    With udtUndoInfo(bytNo)
        .CommentCount = objWorksheet.Comments.Count
        If .CommentCount = 0 Then
            Exit Sub
        End If
        ReDim .Comments(1 To .CommentCount)
    End With
    
    'Undoシート上のコメント単位にループ
    For i = 1 To objWorksheet.Comments.Count
        udtUndoInfo(bytNo).Comments(i).Address = objWorksheet.Comments(i).Parent.Address
        With udtUndoInfo(bytNo).Worksheet.Range(objWorksheet.Comments(i).Parent.Address).Comment.Shape
            udtUndoInfo(bytNo).Comments(i).Name = .Name
            udtUndoInfo(bytNo).Comments(i).Height = .Height
            udtUndoInfo(bytNo).Comments(i).Width = .Width
            udtUndoInfo(bytNo).Comments(i).Top = .Top
            udtUndoInfo(bytNo).Comments(i).Left = .Left
        End With
    Next i
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
    Case E_EraseCol, E_EraseRow
        Dim i As Long
        '名前オブジェクトがセル参照をしているときに復元できるように回復させる
        If Me.UndoType = E_EraseCol Then
            With udtUndoInfo(bytNo)
                lngNo = .Worksheet.Range(.SelectCell).Column
                '元のシートの削除された列の数だけ挿入
                For i = lngNo To lngNo + .Worksheet.Range(.SelectCell).Columns.Count - 1
                    Call .Worksheet.Columns(lngNo).Insert
                Next i
            End With
        Else
            With udtUndoInfo(bytNo)
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
            With udtUndoInfo(bytNo)
                lngNo = .Worksheet.Range(.SelectCell).Column
                '元のシートの挿入された列を削除
                Call .Worksheet.Range(.Worksheet.Columns(lngNo + 1), .Worksheet.Columns(lngNo + .SpiltCount - 1)).Delete
            End With
        Else
            With udtUndoInfo(bytNo)
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
    Case E_MergeCell, E_CellBorder
        'セルを復元
        Call RecoverSheet
        'コメントを復元
        Call RecoverComments
    Case E_CellToText
        'セルを復元
        Call RecoverSheet
        '追加されたテキストボックスを削除
        Call DeleteNewShapes
    Case E_TextToCell
        'セルを復元
        Call RecoverSheet
        'コメントを復元
        Call RecoverComments
        'テキストボックスを復元する
        Call RecoverTextbox
    Case E_ShapeSize, E_ShapeSize2
        '図形のサイズを復元する
        Call RecoverShapeSize
    Case E_CellValue, E_PasteValue
        'セルを復元
        Call RecoverSheet
    End Select
    
    '復元対象になったセルやオブジェクトの選択
    With udtUndoInfo(bytNo)
        Call .Worksheet.Activate
        Select Case (Me.UndoType)
        Case E_ColSize, E_RowSize, E_ColSize2, E_RowSize2, _
             E_MergeCell, E_MoveCell, E_CopyCell, E_PasteValue, _
             E_CellToText, _
             E_SplitCol, E_EraseCol, _
             E_SplitRow, E_EraseRow, _
             E_CellValue
            Call .Worksheet.Range(.SelectCell).Select
        Case E_ShapeSize, E_ShapeSize2
            Call GetShapeRangeFromID(.SelectShapes).Select
        Case E_TextToCell
        End Select
    End With

    'クリップボードの中身を復元
    Call RecoverClipbordText

    '最後のセルを修正する
    Call ActiveSheet.UsedRange  'Undo出来なくなります
End Sub

'*****************************************************************************
'[ 関数名 ]　RecoverTextbox
'[ 概  要 ]　テキストボックスを復元する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub RecoverTextbox()
    Dim objTextbox   As Object
    
    With udtUndoInfo(bytNo)
        'テキストボックスを復元
        Set objTextbox = .UndoSheet.Shapes(1)
        Call objTextbox.Cut
        Call .Worksheet.Paste
    End With
    
'    With udtUndoInfo(bytNo).Worksheet.Shapes
'        Set objTextbox = .Item(.Count)
'    End With

    Set objTextbox = Selection
    
    With udtUndoInfo(bytNo)
        'テキストボックスの位置を復元
        objTextbox.Top = .Shapes(1).Top
        objTextbox.Left = .Shapes(1).Left
    End With
    
    If udtUndoInfo(bytNo).ShapeCount > 1 Then
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
 
    With udtUndoInfo(bytNo)
        'Undoシートから変更前のセルに移動
'        Call .Worksheet.Cells.UnMerge
        Call .UndoSheet.Cells.Copy(.Worksheet.Cells) 'CutだとExcel2007で幅が復元されない
        Call .UndoSheet.Cells.Clear 'コメントオブジェクトを削除する必要がある
    End With
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
Exit Sub
ErrHandle:
    Application.CopyObjectsWithCells = blnCopyObjectsWithCells
    Call Err.Raise(Err.NUMBER, Err.Source, Err.Description)
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
    
    Set objWorksheet = udtUndoInfo(bytNo).Worksheet
    
    'コメント以外のオブジェクトで保存対象でなかったものは削除する
    For Each objShape In objWorksheet.Shapes
        If objShape.Type <> msoComment Then
            blnHit = False
            For i = 1 To udtUndoInfo(bytNo).ShapeCount
                If objShape.ID = udtUndoInfo(bytNo).Shapes(i).ID Then
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
    
    Set objWorksheet = udtUndoInfo(bytNo).Worksheet
    
    For i = 1 To udtUndoInfo(bytNo).RangeCount
        With udtUndoInfo(bytNo).RangeSize(i)
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
    
    Set objWorksheet = udtUndoInfo(bytNo).Worksheet
    
    For Each objShape In objWorksheet.Shapes
        For i = 1 To udtUndoInfo(bytNo).ShapeCount
            With udtUndoInfo(bytNo).Shapes(i)
                If .ID = objShape.ID Then
                    objShape.Height = .Height
                    objShape.Width = .Width
                    objShape.Top = .Top
                    objShape.Left = .Left
                    .Recover = True
                    Exit For
                End If
            End With
        Next
    Next

    For i = 1 To udtUndoInfo(bytNo).ShapeCount
        With udtUndoInfo(bytNo).Shapes(i)
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
    
    Set objWorksheet = udtUndoInfo(bytNo).Worksheet
    
    '保存されたコメントの数だけループ
    For i = 1 To udtUndoInfo(bytNo).CommentCount
        For Each objComment In objWorksheet.Comments
            If udtUndoInfo(bytNo).Comments(i).Address = objComment.Parent.Address Then
                udtUndoInfo(bytNo).Comments(i).Recover = True
                Exit For
            End If
        Next objComment
        
        If udtUndoInfo(bytNo).Comments(i).Recover = True Then
            With objComment.Shape
                .Height = udtUndoInfo(bytNo).Comments(i).Height
                .Width = udtUndoInfo(bytNo).Comments(i).Width
                .Top = udtUndoInfo(bytNo).Comments(i).Top
                .Left = udtUndoInfo(bytNo).Comments(i).Left
            End With
        End If
    Next i
    
    '復元もれのチェック(Debug用)
    For i = 1 To udtUndoInfo(bytNo).CommentCount
        With udtUndoInfo(bytNo).Comments(i)
            If .Recover = False Then
                Call MsgBox(.Name & "(" & .Address & ")" & "が復元されませんでした")
            End If
        End With
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　SetOnUndo
'[ 概  要 ]　ApplicationオブジェクトのOnUndoイベントを設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetOnUndo()
    
    Select Case (Me.UndoType)
    Case E_ColSize, E_ColSize2
        Call Application.OnUndo("幅の変更", "ExecUndo")
    Case E_RowSize, E_RowSize2
        Call Application.OnUndo("高さの変更", "ExecUndo")
    Case E_ShapeSize, E_ShapeSize2
        Call Application.OnUndo("図形のサイズの変更", "ExecUndo")
    Case E_CellBorder
        Call Application.OnUndo("セルの境界移動", "ExecUndo")
    Case E_MergeCell
        Call Application.OnUndo("セルの結合", "ExecUndo")
    Case E_MoveCell
        Call Application.OnUndo("領域の移動", "ExecUndo")
    Case E_CopyCell
        Call Application.OnUndo("領域のコピー", "ExecUndo")
    Case E_PasteValue
        Call Application.OnUndo("値の貼り付け", "ExecUndo")
    Case E_SplitCol
        Call Application.OnUndo("列の分割", "ExecUndo")
    Case E_EraseCol
        Call Application.OnUndo("列の消去", "ExecUndo")
    Case E_SplitRow
        Call Application.OnUndo("行の分割", "ExecUndo")
    Case E_EraseRow
        Call Application.OnUndo("行の消去", "ExecUndo")
    Case E_CellToText
        Call Application.OnUndo("セルをテキストボックスに変換", "ExecUndo")
    Case E_TextToCell
        Call Application.OnUndo("テキストボックスをセルに変換", "ExecUndo")
    Case E_CellValue
        Call Application.OnUndo("セルの値を変更", "ExecUndo")
    Case E_FilterERR
        Call Application.OnUndo("フィルタ使用中の操作はUndo出来ません", "ExecUndo")
    Case E_UndoERR
        Call Application.OnUndo("変更前の状態の保存に失敗しました", "ExecUndo")
    End Select
    
    'クリップボードの中身を復元
    Call RecoverClipbordText
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCopyCell 
   Caption         =   "見たまま貼付け"
   ClientHeight    =   2976
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5244
   OleObjectBlob   =   "frmCopyCell.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCopyCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TRect
    Top      As Long
    Height   As Long
    Left     As Long
    Width    As Long
End Type

'Private blnCheck As Boolean

Private FEnableEvents As Boolean
Private FMove As Boolean
Private FBlockCount As Long
Private FWidthRatio() As Double
Private FBackupDstColWidth() As Double
Private FSrcRange  As Range
Private FDstRange  As Range
Private FSrcColAddress() As String
Private FDstColAddress() As String

Private objTextbox()   As Shape  'idx 0:外枠,1～FBlockCount:列見出し,FBlockCount:内枠
Private lngDisplayObjects As Long
Private lngZoom      As Long

'*****************************************************************************
'[概要] 初期設定情報を設定
'[引数] objFromRange:コピー元の領域
'       objToRange:貼付け先のセル
'       blnMove:True=移動する
'[戻値] なし
'*****************************************************************************
Public Sub Initialize(ByRef objFromRange As Range, ByRef objToRange As Range, ByVal blnMove As Boolean)
    chkIgnore.ControlTipText = "通常は､デフォルトのまま実行ください(コピー元をEXCEL方眼と判定した時にチェックされます)"
    chkBlank.ControlTipText = "貼付け後、空白セルの結合を解除します"
    FMove = blnMove
    
    lngDisplayObjects = ActiveWorkbook.DisplayDrawingObjects
    FEnableEvents = False
    chkIgnore.Value = IsGraphpaper(objFromRange)
    FEnableEvents = True
    
    Call Init(objFromRange)
    Call GetDstRange(objToRange)
    Call FDstRange.Select
    Call SetDstColAddress
    
    'コピー先のシートをActivateにする
    Call objToRange.Worksheet.Activate
    
    Call MakeTextBox

'    '選択領域が画面から消えている時
'    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '画面分割のない時
'        If IntersectRange(ActiveWindow.VisibleRange, objToRange) Is Nothing Then
'            With objToRange
'                Call ActiveWindow.ScrollIntoView(.Left / DPIRatio, .Top / DPIRatio, .Width / DPIRatio, .Height / DPIRatio)
'            End With
'        End If
'    End If
End Sub

'*****************************************************************************
'[概要] テキストボックスを作成
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub MakeTextBox()
    ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
    ReDim objTextbox(0 To FBlockCount + 1)
    
    '内枠作成
    If FDstRange.Rows.Count > 1 Then
        With MinusRange(FDstRange, FDstRange.Rows(1))
            Set objTextbox(FBlockCount + 1) = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
        End With
        'エラーの時に赤字でエラー内容を表示させる
        With objTextbox(FBlockCount + 1).TextFrame2.TextRange.Font
            .NameComplexScript = ActiveWorkbook.Styles("Normal").Font.Name
            .NameFarEast = ActiveWorkbook.Styles("Normal").Font.Name
            .Name = ActiveWorkbook.Styles("Normal").Font.Name
            .Size = ActiveWorkbook.Styles("Normal").Font.Size
        End With
        With objTextbox(FBlockCount + 1).TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.Rgb = Rgb(255, 0, 0) '赤
            .Transparency = 0
        End With
        'テキストボックスの罫線なし
        With objTextbox(FBlockCount + 1).Line
            .Visible = msoFalse
        End With
        With objTextbox(FBlockCount + 1).Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.SchemeColor = 65
            .Transparency = 0.12  '背景を透けさせる
        End With
    End If
    
    '列情報を１行目に表示
    Dim i As Long
    For i = 1 To FBlockCount
        'テキストボックス作成
        With Intersect(Range(FDstColAddress(i)), FDstRange.Rows(1))
            Set objTextbox(i) = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
        End With
        With objTextbox(i).TextFrame2.TextRange.Font
            .NameComplexScript = ActiveWorkbook.Styles("Normal").Font.Name
            .NameFarEast = ActiveWorkbook.Styles("Normal").Font.Name
            .Name = ActiveWorkbook.Styles("Normal").Font.Name
            .Size = ActiveWorkbook.Styles("Normal").Font.Size
        End With
        With objTextbox(i).TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.Rgb = Rgb(0, 0, 0)
            .Transparency = 0
        End With
        With objTextbox(i).TextFrame2
            .VerticalAnchor = msoAnchorMiddle '中央寄せ
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter '中央寄せ
            '前後の余白を0に
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
        End With
        With objTextbox(i).Fill
            .ForeColor.Rgb = Rgb(218, 231, 245)
            .Transparency = 0
        End With
        With objTextbox(i).Line
            .Visible = msoTrue
            .ForeColor.Rgb = Rgb(0, 0, 0)
            .Weight = 0.75
            .Transparency = 0
        End With
    Next
    
    '外枠作成
    With FDstRange
        Set objTextbox(0) = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
    End With
    'テキストボックスの背景なし
    With objTextbox(0).Fill
        .Visible = msoFalse
    End With
    'テキストボックスの罫線を変更
    With objTextbox(0).Line
        .Weight = 2#
        .Style = msoLineSingle
        .Transparency = 0#
        .Visible = msoTrue
        .ForeColor.SchemeColor = 64
        .BackColor.Rgb = Rgb(255, 255, 255)
        .Pattern = msoPattern50Percent
    End With
    
'    Call objTextbox(0).ZOrder(msoBringToFront)
    
    '見出し列にコピー元の列アドレスを表示
    Call EditTextbox
End Sub

'*****************************************************************************
'[概要] テキストボックスの横幅も変更して、見出し行にコピー元の列アドレスを表示
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub EditTextbox()
    Dim i As Long
    '見出し行
    For i = 1 To FBlockCount
        '幅を変更
        With Intersect(Range(FDstColAddress(i)), FDstRange.Rows(1))
            objTextbox(i).Left = .Left
            objTextbox(i).Width = .Width
        End With
        With objTextbox(i).TextFrame2.TextRange
            'テキストボックスに列名を表示
            .Characters.Text = GetColAddress(FSrcColAddress(i))
        End With
    Next
    
    '内枠
    If FDstRange.Rows.Count > 1 Then
        With MinusRange(FDstRange, FDstRange.Rows(1))
            objTextbox(FBlockCount + 1).Left = .Left
            objTextbox(FBlockCount + 1).Width = .Width
        End With
    End If
    
    '外枠
    With FDstRange
        objTextbox(0).Left = .Left
        objTextbox(0).Width = .Width
    End With

    lblCellAddress.Caption = " " & FDstRange.Address(0, 0)
    Call CheckPaste
End Sub

'*****************************************************************************
'[概要] 例：A:A → A, A:B → A:B
'[引数] 例：A:A
'[戻値] 例：A
'*****************************************************************************
Private Function GetColAddress(ByVal strAddress As String) As String
    Dim strABC
    GetColAddress = strAddress
    On Error Resume Next
    strABC = Split(strAddress, ":")
    If strABC(0) = strABC(1) Then
        GetColAddress = strABC(0)
    End If
End Function

'*****************************************************************************
'[概要] FSrcColAddressとFSrcWidthを設定
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub Init(ByRef objSelection As Range)
    Dim i As Long
    Set FSrcRange = objSelection
    
    'True:セルの左側が結合セルの境界となる列
    ReDim IsBoderCol(1 To FSrcRange.Columns.Count + 1) As Boolean
    
    Dim objArea As Range
    Dim RightCol As Long 'セルの右側が結合セルの境界となる列番号
    Dim objWkColumns As Range
    Dim Offset As Long
    Offset = objSelection.Column - 1
    
    '選択された列数分LoopしてIsBoderCol()を設定
    IsBoderCol(1) = True
    IsBoderCol(FSrcRange.Columns.Count + 1) = True
    Dim x As Long, y As Long
    For x = 2 To FSrcRange.Columns.Count
        For y = 1 To FSrcRange.Rows.Count
            If chkIgnore.Value Then
                With FSrcRange.Cells(y, x).MergeArea
                    '結合セルの時
                    If .Count > 1 Then
                        '左端の列にフラグをたてる
                        IsBoderCol(.Column - Offset) = True
                        '右端の右隣の列にフラグをたてる
                        IsBoderCol(.Column + .Columns.Count - Offset) = True
                    ElseIf VarType(.Value) <> vbEmpty Then
                        '値の入力されたセルの時、フラグをたてる
                        IsBoderCol(x) = True
                    End If
                End With
            Else
                With FSrcRange.Cells(y, x)
                    If .MergeArea.Column = .Column Then
                        IsBoderCol(x) = True
                    End If
                End With
            End If
        Next
    Next
    
    Dim j As Long
    Dim LeftCol As Long  'セルの左側が結合セルの境界となる列番号
    LeftCol = 1
    ReDim FSrcWidth(1 To FSrcRange.Columns.Count)
    ReDim FSrcColAddress(1 To FSrcRange.Columns.Count)
    '選択された列数分Loop
    For i = 2 To FSrcRange.Columns.Count + 1
        If IsBoderCol(i) Then
            With FSrcRange
                j = j + 1
                '該当列のアドレス(列名)を設定　※例 D:F
                FSrcColAddress(j) = GetColumnsAddress(FSrcRange, LeftCol, i - 1)
                '次の塊の右側が結合セルの境界となる列
                LeftCol = i
            End With
        End If
    Next
    
    '配列サイズを設定
    FBlockCount = j
    
    '配列を前詰めに圧縮
    ReDim Preserve FSrcColAddress(1 To FBlockCount)
    ReDim FWidthRatio(1 To FBlockCount)
    For i = 1 To FBlockCount
        '結合セルの塊ごとの全体幅に対する割合を設定
        FWidthRatio(i) = FSrcRange.Worksheet.Range(FSrcColAddress(i)).Width / FSrcRange.Width
    Next
End Sub

'*****************************************************************************
'[概要] 幅が近いRangeを取得する
'[引数] 対象Range、新しい列数
'[戻値] なし
'*****************************************************************************
Private Function GetDstRange(ByRef objTopLeftCell As Range) As Range
    Dim objWkRange As Range
    Dim DstColCnt  As Long
    
    Set objWkRange = objTopLeftCell.Resize(, Columns.Count - objTopLeftCell.Column + 1)
    DstColCnt = WorksheetFunction.Max(GetColNumber(objWkRange, FSrcRange.Width), FBlockCount)
    Set FDstRange = objTopLeftCell.Resize(FSrcRange.Rows.Count, DstColCnt)
    
    Set GetDstRange = FDstRange
End Function

'*****************************************************************************
'[概要] FDstColAddressを設定
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SetDstColAddress()
    ReDim FDstColAddress(1 To FBlockCount)
    
    ReDim LeftCol(1 To FBlockCount) As Long  '幅を判定する左端の列番号
    Dim RightCol As Long '幅を判定する右端の列番号
    LeftCol(1) = 1
    RightCol = FDstRange.Columns.Count - FBlockCount
    
    Dim dblWidth As Double
    Dim objWkRange As Range
    Dim i As Long
    For i = 1 To FBlockCount - 1
        RightCol = RightCol + 1
        dblWidth = SumRatio(1, i) * FDstRange.Width
        Set objWkRange = FDstRange.Resize(, RightCol)
        LeftCol(i + 1) = GetColNumber(objWkRange, dblWidth) + 1
        LeftCol(i + 1) = WorksheetFunction.Max(LeftCol(i + 1), LeftCol(i) + 1)
    Next
    For i = 1 To FBlockCount - 1
        FDstColAddress(i) = GetColumnsAddress(FDstRange, LeftCol(i), LeftCol(i + 1) - 1)
    Next
    
    '最後の列は残り全て
    FDstColAddress(FBlockCount) = GetColumnsAddress(FDstRange, LeftCol(i), FDstRange.Columns.Count)
End Sub

'*****************************************************************************
'[概要] StartIdxからEndIdxまでの全体幅に対する割合を合計する
'[引数] StartIdx,EndIdx
'[戻値] 幅の合計
'*****************************************************************************
Private Function SumRatio(ByVal StartIdx As Long, Optional ByVal EndIdx As Long) As Double
    Dim i As Long
    For i = StartIdx To EndIdx
        SumRatio = SumRatio + FWidthRatio(i)
    Next
End Function

'*****************************************************************************
'[概要] 対象とするRangeのStartColxからEndColまでの列アドレスを取得する
'[引数] 対象とするRange,Start列,End列
'[戻値] 対象の列アドレス　例：C:E
'*****************************************************************************
Private Function GetColumnsAddress(ByRef objRange As Range, ByVal StartCol As Long, ByVal EndCol As Long) As String
    Dim objWkRange As Range
    With objRange
        Set objWkRange = .Worksheet.Range(.Columns(StartCol), .Columns(EndCol))
    End With
    GetColumnsAddress = objWkRange.EntireColumn.Address(0, 0)
End Function

'*****************************************************************************
'[概要] 与件の列範囲内で該当の幅に近い列番号(objRange内の)を取得する
'[引数] 列範囲、取得幅
'[戻値] 列を表すアドレス　例 C:E
'*****************************************************************************
Private Function GetColNumber(ByRef objRange As Range, ByVal dblWidth As Double) As Long
    If objRange.Columns.Count = 1 Then
        GetColNumber = 1
        Exit Function
    End If
    
    If objRange.Width <= dblWidth Then
        GetColNumber = objRange.Columns.Count
        Exit Function
    End If
    
    '幅を判定する右隣のセルの真ん中
    Dim dblHalf As Double
    
    Dim i As Long
    For i = 1 To objRange.Columns.Count - 1
        dblHalf = objRange.Columns(i + 1).Width / 2
        If dblWidth <= objRange.Resize(, i).Width + dblHalf Then
            Exit For
        End If
    Next
    GetColNumber = i
End Function

'*****************************************************************************
'[概要] 列数を増減させる
'[引数] 対象Range、新しい列数
'[戻値] なし
'*****************************************************************************
Private Sub SplitOrEraseCol(ByRef objRange As Range, ByVal NewColCount As Long)
    Dim OldColCount As Long
    OldColCount = objRange.Columns.Count
    
    If OldColCount = NewColCount Then
        Exit Sub
    End If

    If NewColCount > OldColCount Then
        With objRange
            Call SplitCol(.Columns(OldColCount), NewColCount - OldColCount + 1)
        End With
    Else
        With objRange
            Call EraseCol(.Worksheet.Range(.Columns(NewColCount + 1), .Columns(OldColCount)))
        End With
    End If
End Sub

'*****************************************************************************
'[概要] 列を分割する
'[引数] 対象列、分割数
'[戻値] なし
'*****************************************************************************
Private Sub SplitCol(ByRef objRange As Range, ByVal SplitCount As Long)
    Dim objNewCol As Range
    Dim i As Long
    
    '選択列の右側に1列挿入
    Call objRange.Columns(2).Insert(xlShiftToRight, xlFormatFromLeftOrAbove)
    
    '新しい列
    Set objNewCol = objRange.Columns(2)
    
    '挿入列の１セル毎に罫線をコピーする
    Call CopyBorder("右上下", objRange, objNewCol)
    
    '横方向に結合する
    Dim objMergeRange As Range
'    Dim lngtype As Long
'    If chkIgnore.Value Then
'        lngtype = 2 '結合されていないセルは結合しない
'    Else
'        lngtype = 1 'すべて横方向に結合する
'    End If
    
    '1行毎にLoop
    For i = 1 To objRange.Rows.Count
        If objNewCol.Cells(i, 1).MergeArea.Count = 1 Then
            Set objMergeRange = GetMergeColRange(1, objRange.Cells(i, 1), objNewCol.Cells(i, 1))
            If Not (objMergeRange Is Nothing) Then
                Call objMergeRange.Merge
            End If
        End If
    Next
    
    '分割数だけ、列を挿入する
    For i = 3 To SplitCount
        Call objNewCol.EntireColumn.Insert
    Next
End Sub

'*****************************************************************************
'[概要] 列を消去する
'[引数] 対象列
'[戻値] なし
'*****************************************************************************
Private Sub EraseCol(ByRef objRange As Range)
    With objRange
        '右端の罫線をコピーする
        Call CopyBorder("右", .Columns(.Columns.Count), .Columns(0))
    
        '削除
        Call .Delete(xlShiftToLeft)
    End With
End Sub

'*****************************************************************************
'[概要] 結合されていないセルは無視するチェック時
'*****************************************************************************
Private Sub chkIgnore_Click()
    If Not FEnableEvents Then Exit Sub
    
    If chkWidth.Value Then
        chkWidth.Value = False
    End If
    
    Dim i As Long
    For i = LBound(objTextbox) To UBound(objTextbox)
        If Not (objTextbox(i) Is Nothing) Then
            Call objTextbox(i).Delete
            Set objTextbox(i) = Nothing
        End If
    Next
    
    Call Init(FSrcRange)
    Call GetDstRange(FDstRange)
    Call FDstRange.Select
    Call SetDstColAddress
    Call MakeTextBox
End Sub

'*****************************************************************************
'[概要] コピー元の幅を再現するチェック時
'*****************************************************************************
Private Sub chkWidth_Click()
    Static BackupDstRange As Range
    Application.ScreenUpdating = False
        
    Dim i As Long
    Dim ColWidth As Double
    If chkWidth.Value Then
        ReDim FBackupDstColWidth(1 To FBlockCount)
        Set BackupDstRange = FDstRange
        Set FDstRange = FDstRange.Resize(, FBlockCount)
        For i = 1 To FBlockCount
            ColWidth = FSrcRange.Width * FWidthRatio(i) / DPIRatio
            FBackupDstColWidth(i) = FDstRange.Columns(i).EntireColumn.ColumnWidth
            FDstRange.Columns(i).EntireColumn.ColumnWidth = PixelToWidth(ColWidth)
        Next
        If FSrcRange.Rows.Count > 1 Then
            With objTextbox(FBlockCount + 1).TextFrame2.TextRange
                .Characters.Text = "コピー元の幅を再現する時は、幅を変更できません"
            End With
        End If
    Else
        For i = 1 To FBlockCount
            FDstRange.Columns(i).EntireColumn.ColumnWidth = FBackupDstColWidth(i)
        Next
        Set FDstRange = BackupDstRange
        If FSrcRange.Rows.Count > 1 Then
            With objTextbox(FBlockCount + 1).TextFrame2.TextRange
                .Characters.Text = ""
            End With
        End If
    End If
    
    '対象列アドレスの再設定
    Call SetDstColAddress
    '見出し列にコピー元の列アドレスを表示
    Call EditTextbox
    
    Call FDstRange.Select
    Application.ScreenUpdating = True
End Sub

'*****************************************************************************
'[概要] フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()
    '呼び元に通知する
    blnFormLoad = True
    lngZoom = ActiveWindow.Zoom
End Sub

'*****************************************************************************
'[概要] フォームアンロード時
'*****************************************************************************
Private Sub UserForm_Terminate()
    '呼び元に通知する
    blnFormLoad = False

    Dim i As Long
    For i = LBound(objTextbox) To UBound(objTextbox)
        If Not (objTextbox(i) Is Nothing) Then
            Call objTextbox(i).Delete
        End If
    Next
    
    ActiveWorkbook.DisplayDrawingObjects = lngDisplayObjects
    ActiveWindow.Zoom = lngZoom
End Sub

'*****************************************************************************
'[概要] ＯＫボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrHandle
    Dim blnCopyObjectsWithCells  As Boolean
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells
    Application.CopyObjectsWithCells = False '呼び元で復元するため当モジュールでは復元しない

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False 'コメントがある時警告が出る時がある

    Dim i As Long
    For i = LBound(objTextbox) To UBound(objTextbox)
        If Not (objTextbox(i) Is Nothing) Then
            Call objTextbox(i).Delete
            Set objTextbox(i) = Nothing
        End If
    Next
    
    'アンドゥ用に元の幅を復元する
    If chkWidth.Value Then
        ReDim SaveDstColWidth(1 To FBlockCount) As Double
        For i = 1 To FBlockCount
            SaveDstColWidth(i) = FDstRange.Columns(i).EntireColumn.ColumnWidth
        Next
        Call BackupDstColWidth
    End If
    
    'アンドゥ用に元の状態を保存する
    If FMove Then
        Call SaveUndoInfo(E_CopyCell, FSrcRange)
    Else
        Call SaveUndoInfo(E_CopyCell, FDstRange)
    End If

    '幅を再復元する
    If chkWidth.Value Then
        For i = 1 To FBlockCount
            FDstRange.Columns(i).EntireColumn.ColumnWidth = SaveDstColWidth(i)
        Next
    End If
    
    'セルをコピーする
    Call CopyCell

    Call FDstRange.Select
    
    Application.DisplayAlerts = True
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea2"))
    Call Unload(Me)
    Call SetOnUndo
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Call MsgBox(Err.Description, vbExclamation)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea2"))
    Call Unload(Me)
End Sub

'*****************************************************************************
'[概要] キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call BackupDstColWidth
    Call Unload(Me)
End Sub

'*****************************************************************************
'[概要] ヘルプボタン押下時
'*****************************************************************************
Private Sub cmdHelp_Click()
    Call OpenHelpPage("http://takana.web5.jp/EasyLout/V5/Clipbord.htm#PasteAppearance")
End Sub

'*****************************************************************************
'[概要] コピー元の幅を再現するがチェックされている時、元の幅を再現する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub BackupDstColWidth()
    Dim i As Long
    If chkWidth.Value Then
        For i = 1 To FBlockCount
            FDstRange.Columns(i).EntireColumn.ColumnWidth = FBackupDstColWidth(i)
        Next
    End If
End Sub

'*****************************************************************************
'[概要] 領域をコピーする
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub CopyCell()
    Dim objWkRange As Range
    
    '領域をワークシートにコピーする
    Set objWkRange = CopyToWkSheet()

    '移動時はコピー元をクリアする
    If FMove Then
        With FSrcRange
            .ClearContents
            .UnMerge
            .Borders.LineStyle = xlNone
        End With
    End If
    
    Dim objRange As Range
    Dim lngOffset As Long
    lngOffset = FDstRange.Column - FSrcRange.Column
    'ワークシートで領域のサイズを変更する
    Dim i As Long
    For i = FBlockCount To 1 Step -1
        Set objRange = Intersect(objWkRange, objWkRange.Worksheet.Range(FSrcColAddress(i)).Offset(, lngOffset))
        Call SplitOrEraseCol(objRange, Range(FDstColAddress(i)).Columns.Count)
    Next
    
    Call objWkRange.Resize(, FDstRange.Columns.Count).Copy(FDstRange)
    If chkBlank.Value Then
        Set objWkRange = GetBlankAndMergeRange(FDstRange)
        If Not (objWkRange Is Nothing) Then
            Call objWkRange.UnMerge
            objWkRange.HorizontalAlignment = xlGeneral
        End If
    End If
End Sub

'*****************************************************************************
'[概要] 空白 かつ 結合されたセルを取得する
'[引数] 対象領域
'[戻値] 空白 かつ 結合されたセル
'*****************************************************************************
Private Function GetBlankAndMergeRange(ByRef objSelection As Range) As Range
    Dim objRange   As Range
    Dim objCell    As Range
    
    '結合されたセルはUsedRange以外にはないので
    Set objRange = IntersectRange(objSelection, GetUsedRange())
    If objRange Is Nothing Then
        Exit Function
    End If
    
    'セルの数だけループ
    For Each objCell In objRange
        '空白か？
        If objCell.Value = "" And objCell.Formula = "" Then
            With objCell.MergeArea
                '結合セルか？
                If .Count > 1 Then
                    '左上のセルか
                    If .Row = objCell.Row And .Column = objCell.Column Then
                        Set GetBlankAndMergeRange = UnionRange(GetBlankAndMergeRange, objCell)
                    End If
                End If
            End With
        End If
    Next
End Function

'*****************************************************************************
'[イベント] KeyDown
'[概要] カーソルキーで移動先を変更させる
'*****************************************************************************
Private Sub cmdOK_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub cmdCancel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub cmdHelp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub fraKeyCapture_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub chkIgnore_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub chkWidth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub

'*****************************************************************************
'[イベント] UserForm_KeyDown
'[概要] カーソルキーで移動先を変更させる
'*****************************************************************************
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i         As Long
    Dim lngTop    As Long
    Dim lngLeft   As Long
    Dim lngBottom As Long
    Dim lngRight  As Long

    Select Case (KeyCode)
    Case vbKeyLeft, vbKeyRight, vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyUp, vbKeyDown
        Call fraKeyCapture.SetFocus
    Case Else
        Exit Sub
    End Select

    'Altが押されていればスクロール
    If GetKeyState(vbKeyMenu) < 0 Then
        Select Case (KeyCode)
        Case vbKeyLeft
            Call ActiveWindow.SmallScroll(, , , 1)
        Case vbKeyRight
            Call ActiveWindow.SmallScroll(, , 1)
        Case vbKeyUp
            Call ActiveWindow.SmallScroll(, 1)
        Case vbKeyDown
            Call ActiveWindow.SmallScroll(1)
        End Select
        Exit Sub
    End If

    'Zoom
    Select Case (KeyCode)
    Case vbKeyHome, vbKeyPageUp, vbKeyPageDown
        Select Case (KeyCode)
        Case vbKeyHome
            ActiveWindow.Zoom = lngZoom
        Case vbKeyPageUp
            ActiveWindow.Zoom = WorksheetFunction.Min(ActiveWindow.Zoom + 10, 400)
        Case vbKeyPageDown
            ActiveWindow.Zoom = WorksheetFunction.Max(ActiveWindow.Zoom - 10, 10)
        End Select

        With FDstRange
            lngLeft = WorksheetFunction.Max(.Left / DPIRatio - 1, 0) * ActiveWindow.Zoom / 100
            lngTop = WorksheetFunction.Max(.Top / DPIRatio - 1, 0) * ActiveWindow.Zoom / 100
            Call ActiveWindow.ScrollIntoView(lngLeft, lngTop, 1, 1)
        End With
        Exit Sub
    End Select
    
    'コピー元の幅を再現する時は、変更不可
    If chkWidth.Value Then
        Exit Sub
    End If
    
    '選択領域の四方の位置を待避
    With FDstRange
        lngTop = .Row
        lngBottom = .Row + .Rows.Count - 1
        lngLeft = .Column
        lngRight = .Column + .Columns.Count - 1
    End With

    Select Case (Shift)
    Case 0
        '選択領域を移動
        Select Case (KeyCode)
        Case vbKeyLeft
            lngLeft = lngLeft - 1
            lngRight = lngRight - 1
        Case vbKeyRight
            lngLeft = lngLeft + 1
            lngRight = lngRight + 1
        End Select
    Case 1
        '選択領域の大きさを変更
        If GetKeyState(vbKeyZ) < 0 Then
            Select Case (KeyCode)
            Case vbKeyLeft
                lngLeft = lngLeft - 1
            Case vbKeyRight
                lngLeft = lngLeft + 1
            End Select
        Else
            Select Case (KeyCode)
            Case vbKeyLeft
                lngRight = lngRight - 1
            Case vbKeyRight
                lngRight = lngRight + 1
            End Select
        End If
    Case Else
        Exit Sub
    End Select

    'チェック
    If (FBlockCount <= lngRight - lngLeft + 1) And _
       (1 <= lngLeft And lngRight <= Columns.Count) Then
        Set FDstRange = Range(Cells(lngTop, lngLeft), Cells(lngBottom, lngRight))
        Call SetDstColAddress
    Else
        Exit Sub
    End If
    
    'テキストボックスを編集
    Call EditTextbox

    '選択領域が画面から消えたら画面をスクロール
    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '画面分割のない時
        Select Case (KeyCode)
        Case vbKeyLeft
            i = WorksheetFunction.Max(FDstRange.Column - 1, 1)
            If IntersectRange(ActiveWindow.VisibleRange, Columns(i)) Is Nothing Then
                Call ActiveWindow.SmallScroll(, , , 1)
            End If
        Case vbKeyRight
            i = WorksheetFunction.Min(FDstRange.Column + FDstRange.Columns.Count, Columns.Count)
            If IntersectRange(ActiveWindow.VisibleRange, Columns(i)) Is Nothing Then
                Call ActiveWindow.SmallScroll(, , 1)
            End If
        End Select
    End If
End Sub

'*****************************************************************************
'[概要] 貼付け可能かどうかチェック
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub CheckPaste()
On Error GoTo ErrHandle
    If CheckBorder() = True Then
        Call Err.Raise(C_CheckErrMsg, , "結合されたセルの一部を変更することはできません")
    End If
    If FDstRange.Rows.Count > 1 And chkWidth.Value = False Then
        With objTextbox(FBlockCount + 1).TextFrame.Characters
            .Text = ""
        End With
    End If
    cmdOK.Enabled = True  'OKボタンを使用可にする
Exit Sub
ErrHandle:
    If Err.Number = C_CheckErrMsg Then
        If FSrcRange.Rows.Count > 1 Then
            With objTextbox(FBlockCount + 1).TextFrame.Characters
                .Text = Err.Description
            End With
        End If
        cmdOK.Enabled = False 'OKボタンを使用不可にする
    Else
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If
End Sub

'*****************************************************************************
'[概要] 貼付け先の四方が結合セルをまたいでいるかどうか
'[引数] なし
'[戻値] True:結合セルをまたいでいる、False:いない
'*****************************************************************************
Private Function CheckBorder() As Boolean
    Dim objChkRange As Range
    If FMove Then
        Set objChkRange = MinusRange(ArrangeRange(FDstRange), UnionRange(FSrcRange, FDstRange))
    Else
        Set objChkRange = MinusRange(ArrangeRange(FDstRange), FDstRange)
    End If
    CheckBorder = Not (objChkRange Is Nothing)
End Function

'*****************************************************************************
'[概要] ワークシートに領域をコピーする
'[引数] なし
'[戻値] ワークシートのコピーされたRange(サイズはコピー元のサイズ)
'*****************************************************************************
Private Function CopyToWkSheet() As Range
    Dim objWorksheet As Worksheet

    Set objWorksheet = ThisWorkbook.Worksheets("Workarea1")
    Call DeleteSheet(objWorksheet)

    '「Workarea」シートに選択領域を複写する
    With FSrcRange
        Set CopyToWkSheet = objWorksheet.Range(FDstRange.Address).Resize(.Rows.Count, .Columns.Count)
        Call .Copy(CopyToWkSheet)
    End With
End Function

'*****************************************************************************
'[概要] RangeがExcel方眼紙かどうか判定
'[引数] 判定するRange
'[戻値] True:方眼紙
'*****************************************************************************
Public Function IsGraphpaper(ByRef objRange As Range) As Boolean
    Dim lngColCnt As Long
    Dim lngRowCnt As Long
    Dim i As Long
    
    '可視セルの数
    For i = 1 To objRange.Columns.Count
        If Not objRange.Columns(i).Hidden Then
            lngColCnt = lngColCnt + 1
        End If
    Next
    For i = 1 To objRange.Rows.Count
        If Not objRange.Rows(i).Hidden Then
            lngRowCnt = lngRowCnt + 1
        End If
    Next
    
    '1セルの幅の平均が、1セルの高さの平均の2倍以下の時は方眼紙とする
    If objRange.Width / lngColCnt <= objRange.Height * 2 / lngRowCnt Then
        IsGraphpaper = True
    Else
        IsGraphpaper = False
    End If
End Function





Attribute VB_Name = "Organize"
Option Explicit
Option Private Module

Private FCount As Long

'*****************************************************************************
'[概要] 標準フォントの変更
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ChangeNormalFont()
    ActiveWorkbook.Styles("Normal").Font.Name = GetSetting(REGKEY, "KEY", "FontName", DEFAULTFONT)
    Call ChangeDefaultShape
End Sub

'**********************************************
'[概要] 既定の図形を変更
'[引数] なし
'[戻値] なし
'**********************************************
Private Sub ChangeDefaultShape()
    Dim objShape As Shape
    Set objShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 10, 10, 100, 50)
    Call SetShapeProperties(objShape, RGB(255, 0, 0), 1.5, False)
    
    Set objShape = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 110, 100, 50)
    Call SetShapeProperties(objShape, RGB(0, 0, 0), 0.75, True)
    
    Set objShape = ActiveSheet.Shapes.AddLine(120, 10, 200, 50)
    Call SetShapeProperties(objShape, RGB(0, 0, 0), 0.75, True)
End Sub

'**********************************************
'[概要] 図形のプロパティを設定し、既定値として保存後削除する
'[引数] objShape: 設定対象の図形
'       lngLineColor: 線の色（RGB値）
'       dblLineWeight: 線の太さ
'       blnFillVisible: 塗りつぶしの表示/非表示
'[戻値] なし
'**********************************************
Private Sub SetShapeProperties(ByRef objShape As Shape, ByVal lngLineColor As Long, ByVal dblLineWeight As Double, ByVal blnFillVisible As Boolean)

    ' 図形が四角形などの場合は塗りつぶしを設定
    If objShape.Type <> msoLine Then
        With objShape.Fill
            .ForeColor.RGB = RGB(255, 255, 255) ' 白
            .Transparency = 0 ' 不透明
            .Solid ' 単色
            .Visible = blnFillVisible '塗りつぶしの有無
        End With
    End If
    
    ' テキストフレームがある場合はフォント設定
    If objShape.Type <> msoLine Then
        With objShape.TextFrame2.TextRange.Font
            .NameComplexScript = ActiveWorkbook.Styles("Normal").Font.Name
            .NameFarEast = ActiveWorkbook.Styles("Normal").Font.Name
            .Name = ActiveWorkbook.Styles("Normal").Font.Name
        End With
        With objShape.TextFrame2.TextRange.Font.Fill
            .ForeColor.RGB = RGB(0, 0, 0) ' 黒
            .Transparency = 0 ' 不透明
            .Solid ' 単色
            .Visible = True
        End With
    End If
    
    ' 線の設定
    With objShape.Line
        .ForeColor.RGB = lngLineColor
        .Transparency = 0 ' 不透明
        .Weight = dblLineWeight
        .Visible = True
        If objShape.Type = msoLine Then
            .BeginArrowheadStyle = msoArrowheadNone
            .EndArrowheadStyle = msoArrowheadNone
        End If
    End With
    
    ' 既定値として設定して削除
    Call objShape.SetShapesDefaultProperties
    Call objShape.Delete
End Sub

'*****************************************************************************
'[概要] セルに標準フォントを適用
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ApplyNormalFont()
    If ActiveWindow.SelectedSheets.Count = 1 And CheckSelection() = E_Range Then
        Selection.Font.Name = ActiveWorkbook.Styles("Normal").Font.Name
    Else
        Dim ws As Worksheet
        For Each ws In ActiveWindow.SelectedSheets
            ws.Cells.Font.Name = ActiveWorkbook.Styles("Normal").Font.Name
        Next
    End If
End Sub

'*****************************************************************************
'[概要] 図形に標準フォントを適用
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ApplyNormalFontToShape()
    Application.ScreenUpdating = False
    If CheckSelection() = E_Shape Then
        On Error Resume Next
        With Selection.ShapeRange.TextFrame2.TextRange.Font
            .NameComplexScript = ActiveWorkbook.Styles("Normal").Font.Name
            .NameFarEast = ActiveWorkbook.Styles("Normal").Font.Name
            .Name = ActiveWorkbook.Styles("Normal").Font.Name
        End With
        On Error GoTo 0
    Else
        Dim ws As Worksheet
        Dim shp As Shape
        For Each ws In ActiveWindow.SelectedSheets
            For Each shp In ws.Shapes
                On Error Resume Next
                With shp.TextFrame2.TextRange.Font
                    .NameComplexScript = ActiveWorkbook.Styles("Normal").Font.Name
                    .NameFarEast = ActiveWorkbook.Styles("Normal").Font.Name
                    .Name = ActiveWorkbook.Styles("Normal").Font.Name
                End With
                On Error GoTo 0
            Next
        Next
    End If
End Sub

'*****************************************************************************
'[概要] 全てのシートの倍率を100%にして、A1セルを選択
'       & 改ページプレビューを解除
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SelectHomePosition()
    Dim objActiveSheet As Worksheet
    Set objActiveSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    Dim objSheet As Worksheet
    For Each objSheet In Worksheets
        Call objSheet.Activate
        Call ActiveWindow.SmallScroll(, Rows.Count, , Columns.Count)
        Call objSheet.Range("A1").Select
        
        '改ページプレビュー解除
        ActiveWindow.View = xlNormalView
        
        '倍率100%
        ActiveWindow.Zoom = 100
    Next
    
    Call objActiveSheet.Select
End Sub

'*****************************************************************************
'[概要] 数式がエラーのセルを選択
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SelectErrFormula()
    If CheckSelection <> E_Range Then
        Call Range("A1").Select
    End If
    On Error Resume Next
    Selection.SpecialCells(xlCellTypeFormulas, xlErrors).Select
    If Err.Number <> 0 Then
        Call MsgBox("数式がエラーのセルはありません")
        Exit Sub
    End If
End Sub

'*****************************************************************************
'[概要] 数式がエラーになっている条件付き書式を削除
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub DeleteErrFormatConditions()
    Call MsgBox("DeleteErrFormatConditions 工事中")
End Sub

'*****************************************************************************
'[概要] ユーザ定義スタイルをすべて削除
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub DeleteUserStyles()
    '件数のカウント
    FCount = DeleteStyles(ActiveWorkbook, True)
    
    'ホームタブ(スタイル)を表示させる
    Call GetRibbonUI.ActivateTabMso("TabHome")
    
    'タブを切り替えるため、タイマーを使用
    Call Application.OnTime(Now(), MacroName("DeleteUserStyles2"))
End Sub
Private Sub DeleteUserStyles2()
    DoEvents
    
    If FCount = 0 Then
        Call MsgBox("ユーザ定義のスタイルはありません")
        Exit Sub
    End If
    
    Dim strMsg As String
    strMsg = FCount & " 件 のユーザ定義スタイルが見つかりました" & vbLf
    strMsg = strMsg & "削除しますか？"
    Select Case MsgBox(strMsg, vbYesNo)
    Case vbYes
        Call DeleteStyles(ActiveWorkbook)
    End Select
End Sub

'*****************************************************************************
'[概要] ユーザ定義の名前をすべて削除
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub DeleteNameObjects()
    '件数のカウント
    Dim lngCnt As Long
    Dim objName As Name
    lngCnt = DeleteNames(ActiveWorkbook, True)
    If lngCnt = 0 Then
        Call MsgBox("ユーザ定義の名前はありません")
        Call ActiveWorkbook.Activate
        Call ActiveWorkbook.ActiveSheet.Activate
        Call CommandBars.ExecuteMso("NameManager")
        Exit Sub
    End If
    
    Dim strMsg As String
    strMsg = lngCnt & " 件 の名前が見つかりました" & vbLf & vbLf
    strMsg = strMsg & "削除しますか？" & vbCrLf
    strMsg = strMsg & "　「 はい 」････ 削除を実行する" & vbCrLf
    strMsg = strMsg & "　「いいえ」････ 名前の管理画面を表示する"
    Select Case MsgBox(strMsg, vbYesNoCancel + vbQuestion + vbDefaultButton1)
    Case vbYes
        Call DeleteNames(ActiveWorkbook)
    Case vbNo
        Call ActiveWorkbook.Activate
        Call ActiveWorkbook.ActiveSheet.Activate
        Call CommandBars.ExecuteMso("NameManager")
    End Select
End Sub

'*****************************************************************************
'[概要] ユーザ設定のビューをすべて削除
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub DeleteUserViews()
    '件数のカウント
    Dim lngCnt As Long
    lngCnt = ActiveWorkbook.CustomViews.Count
    If lngCnt = 0 Then
        Call MsgBox("ユーザ設定のビューはありません")
        Call ActiveWorkbook.Activate
        Call ActiveWorkbook.ActiveSheet.Activate
        Call CommandBars.ExecuteMso("ViewCustomViews")
        Exit Sub
    End If
    
    Dim strMsg As String
    strMsg = lngCnt & " 件 のユーザ設定のビューが見つかりました" & vbLf & vbLf
    strMsg = strMsg & "削除しますか？" & vbCrLf
    strMsg = strMsg & "　「 はい 」････ 削除を実行する" & vbCrLf
    strMsg = strMsg & "　「いいえ」････ ユーザ設定のビューを表示する"
    Select Case MsgBox(strMsg, vbYesNoCancel + vbQuestion + vbDefaultButton1)
    Case vbYes
        Call DeleteViews(ActiveWorkbook)
    Case vbNo
        Call ActiveWorkbook.Activate
        Call ActiveWorkbook.ActiveSheet.Activate
        Call CommandBars.ExecuteMso("ViewCustomViews")
    End Select
End Sub

'*****************************************************************************
'[概要] 面積ゼロの図形の選択
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SelectFlatShapes()
On Error GoTo ErrHandle
    If ActiveSheet.Shapes.Count = 0 Then
        Call MsgBox("対象の図形はありません")
        Exit Sub
    End If
    If ActiveWorkbook.DisplayDrawingObjects = xlHide Then
        ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
        Call GetRibbonUI.InvalidateControl("C1")
    End If
    
    ReDim lngArray(1 To ActiveSheet.Shapes.Count)
    Dim objShape As Shape
    Dim i As Long
    Dim j As Long
    For i = 1 To ActiveSheet.Shapes.Count
        Set objShape = ActiveSheet.Shapes(i)
        'コメントの図形は対象外とする
        If ActiveSheet.Shapes(i).Type <> msoComment Then
            '直線かどうか
            If TypeName(objShape.DrawingObject) = "Line" Then
                If objShape.Width = 0 And objShape.Height = 0 Then
                    j = j + 1
                    lngArray(j) = i
                End If
            Else
                If objShape.Width = 0 Or objShape.Height = 0 Then
                    j = j + 1
                    lngArray(j) = i
                End If
            End If
        End If
    Next
    
    If j = 0 Then
        Call MsgBox("対象の図形はありません")
        Call ShowSelectionPane
        Exit Sub
    End If

    On Error Resume Next
    ReDim Preserve lngArray(1 To j)
    Call ActiveSheet.Shapes.Range(lngArray).Select
    
    Dim strMsg As String
    strMsg = j & " 個 の対象の図形を選択しました" & vbLf & vbLf
    strMsg = strMsg & "不要であればDeleteキーで削除してください"
    Call MsgBox(strMsg)
    Call ShowSelectionPane
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 罫線と同化した直線の選択
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SelectFlatArrows()
On Error GoTo ErrHandle
    If ActiveSheet.Shapes.Count = 0 Then
        Call MsgBox("対象の直線はありません")
        Exit Sub
    End If
    If ActiveWorkbook.DisplayDrawingObjects = xlHide Then
        ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
        Call GetRibbonUI.InvalidateControl("C1")
    End If
    
    ReDim lngArray(1 To ActiveSheet.Shapes.Count)
    Dim objShape As Shape
    Dim i As Long
    Dim j As Long
    For i = 1 To ActiveSheet.Shapes.Count
        Set objShape = ActiveSheet.Shapes(i)
        '直線
        If TypeName(objShape.DrawingObject) = "Line" Then
            If objShape.Width = 0 And objShape.Height = 0 Then
                j = j + 1
                lngArray(j) = i
            ElseIf objShape.Width = 0 Then
                If objShape.Left = objShape.TopLeftCell.Left Then
                    j = j + 1
                    lngArray(j) = i
                End If
            ElseIf objShape.Height = 0 Then
                If objShape.Top = objShape.TopLeftCell.Top Then
                    j = j + 1
                    lngArray(j) = i
                End If
            End If
        End If
    Next
    
    If j = 0 Then
        Call MsgBox("対象の直線はありません")
        Call ShowSelectionPane
        Exit Sub
    End If

    On Error Resume Next
    ReDim Preserve lngArray(1 To j)
    Call ActiveSheet.Shapes.Range(lngArray).Select

    Dim strMsg As String
    strMsg = j & " 個 の対象の直線を選択しました" & vbLf & vbLf
    strMsg = strMsg & "不要であればDeleteキーで削除してください"
    Call MsgBox(strMsg)
    Call ShowSelectionPane
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] A1形式とR1C1参照形式を相互に切替える
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ToggleA1R1C1()
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub

'*****************************************************************************
'[概要] 名前オブジェクトを削除する
'[引数] Workbook, blnCountOnly:件数のカウントのみの時True
'[戻値] 削除対象のオブジェクトの件数
'*****************************************************************************
Public Function DeleteNames(ByRef objWorkbook As Workbook, Optional ByVal blnCountOnly As Boolean = False) As Long
On Error Resume Next
    Dim objName     As Name
    For Each objName In objWorkbook.Names
        Select Case objName.MacroType
        'EXCEL2019の謎の事象対応(TEXTJOIN関数等を使えば勝手に名前が定義されるが削除すると例外になるので回避)
        Case xlFunction, xlCommand, xlNotXLM
        Case Else
            If Right(objName.RefersTo, 5) = "#REF!" Then
                DeleteNames = DeleteNames + 1
                If Not blnCountOnly Then
                    Call objName.Delete
                    DoEvents
                End If
            ElseIf (Right(objName.Name, Len("Print_Area")) <> "Print_Area") And _
               (Right(objName.Name, Len("Print_Titles")) <> "Print_Titles") And _
               objName.Visible Then
                DeleteNames = DeleteNames + 1
                If Not blnCountOnly Then
                    Call objName.Delete
                    DoEvents
                End If
            End If
        End Select
    Next
End Function

'*****************************************************************************
'[概要] スタイルを削除する
'[引数] Workbook, blnCountOnly:件数のカウントのみの時True
'[戻値] 削除対象のオブジェクトの件数
'*****************************************************************************
Public Function DeleteStyles(ByRef objWorkbook As Workbook, Optional ByVal blnCountOnly As Boolean = False) As Long
On Error Resume Next
    Dim objStyle  As Style
    For Each objStyle In objWorkbook.Styles
        If objStyle.BuiltIn = False Then
            DeleteStyles = DeleteStyles + 1
            If Not blnCountOnly Then
                Call objStyle.Delete
                DoEvents
            End If
        End If
    Next
End Function

'*****************************************************************************
'[概要] ユーザ設定のビューを削除する
'[引数] Workbook
'[戻値] なし
'*****************************************************************************
Private Sub DeleteViews(ByRef objWorkbook As Workbook)
On Error Resume Next
    Dim objView  As CustomView
    For Each objView In objWorkbook.CustomViews
        Call objView.Delete
        DoEvents
    Next
End Sub

'*****************************************************************************
'[概要] 使用されたセルの範囲を最適化する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub UsedRange()
    ActiveSheet.UsedRange.Select
End Sub

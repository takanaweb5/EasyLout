Attribute VB_Name = "basEasyLout"
Option Explicit
Option Private Module

Private Const C_DEBUG = True '開発時はTrueにする

Private Const C_ICONSIZE = 18 'galleryアイコンのサイズ
Public FCommand  As String
Public FParam    As Variant
Public FPressKey As EPressKey
Public Enum EPressKey
    E_Shift = 1
    E_Ctrl = 2
    E_ShiftAndCtrl = 3
End Enum

Public Enum EValueType
    E_FillColor
    E_BMarkColor
End Enum

Private FValues As Variant
Private FIcons(0 To 47) As IPicture
Private FPickupColors As Object

'*****************************************************************************
'[イベント] onLoad
'*****************************************************************************
Private Sub onLoad(Ribbon As IRibbonUI)
    'リボンUIをテンポラリのコマンドバーに保存する
    '(モジュール変数に保存した場合は、例外やコードの強制停止で値が損なわれるため)
    Call CreateTmpCommandBar(Ribbon)
End Sub

''*****************************************************************************
''[概要] 塗りつぶしのデフォルト色を設定
''[引数] なし
''[戻値] なし
''*****************************************************************************
'Private Sub SetDefaultColor()
'    FillColor = rgbYellow
'    BMarkColor = &HFFFFCC '薄い水色
'End Sub

'*****************************************************************************
'[概要] IRibbonUIを保存するCommandBarを作成する
'       あわせて、リボンコントロールの状態を保存するCommandBarControlを作成する
'       モジュール変数に保存した場合は、リコンパイルやコードの強制停止で値が損なわれるため
'[引数] IRibbonUI
'[戻値] なし
'*****************************************************************************
Private Sub CreateTmpCommandBar(ByRef Ribbon As IRibbonUI)
    On Error Resume Next
    Call Application.CommandBars(ThisWorkbook.Name).Delete
    On Error GoTo 0
    
    Dim i As Long
    Dim objCmdBar As CommandBar
    Set objCmdBar = Application.CommandBars.Add(ThisWorkbook.Name, Position:=msoBarPopup, Temporary:=True)
    With objCmdBar.Controls.Add(msoControlButton)
        .Tag = "RibbonUI" & ThisWorkbook.Name
        .Parameter = ObjPtr(Ribbon)
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Tag = "C2" & ThisWorkbook.Name
        .State = True '初期設定はチェックあり
    End With
    With objCmdBar.Controls.Add(msoControlButton)
        .Tag = "FillColor" & ThisWorkbook.Name
        .Parameter = rgbYellow
    End With
    With objCmdBar.Controls.Add(msoControlButton)
        .Tag = "BMarkColor" & ThisWorkbook.Name
        .Parameter = &HFFFFCC   '薄い水色
    End With
End Sub

'*****************************************************************************
'[概要] CommandBarからIRibbonUIを取得する
'[引数] なし
'[戻値] IRibbonUI
'*****************************************************************************
Public Function GetRibbonUI() As IRibbonUI
    Dim Pointer  As LongPtr
    With CommandBars.FindControl(, , "RibbonUI" & ThisWorkbook.Name)
        Pointer = .Parameter
    End With
    Dim obj As Object
    Call CopyMemory(obj, Pointer, Len(Pointer))
    Set GetRibbonUI = obj
End Function

'*****************************************************************************
'[概要] テンポラリのCommandBarControlを取得する
'[引数] Controlを識別するID（リボンコントロールのID）
'[戻値] CommandBarControl
'*****************************************************************************
Public Function GetTmpControl(ByVal strId As String) As CommandBarControl
    Set GetTmpControl = CommandBars.FindControl(, , strId & ThisWorkbook.Name)
End Function

'*****************************************************************************
'[概要] 保存した値を取得する
'[引数] 値のタイプ
'[戻値] 保存した色
'*****************************************************************************
Public Function FColor(ByVal eType As EValueType) As Long
    Select Case eType
    Case E_FillColor
        FColor = CommandBars.FindControl(, , "FillColor" & ThisWorkbook.Name).Parameter
    Case E_BMarkColor
        FColor = CommandBars.FindControl(, , "BMarkColor" & ThisWorkbook.Name).Parameter
    End Select
End Function

'*****************************************************************************
'[イベント] loadImage
'*****************************************************************************
Private Sub loadImage(imageID As String, ByRef returnedVal)
    returnedVal = imageID
End Sub

'*****************************************************************************
'[イベント] getVisible
'*****************************************************************************
Private Sub getVisible(Control As IRibbonControl, ByRef returnedVal)
'    returnedVal = (GetValue(control.Id, "Visible") = 1)
    Select Case Control.ID
    Case "dmy"
        returnedVal = C_DEBUG
    Case Else
        returnedVal = True
    End Select
End Sub
'
'*****************************************************************************
'[イベント] getEnabled
'*****************************************************************************
Private Sub getEnabled(Control As IRibbonControl, ByRef returnedVal)
    Select Case Control.ID
    Case "B311", "B312", "B313", "B314"
        returnedVal = CommandBars.GetEnabledMso("ObjectsAlignTop")
    Case "B315", "B316"
        returnedVal = (CheckSelection() = E_Shape)
    Case "B321" 'セルをテキストボックスに変換
        returnedVal = (CheckSelection() = E_Range)
    Case "B322" 'テキストボックスをセルに変換
        returnedVal = (CheckSelection() = E_Shape)
    Case "B323" 'コメントをテキストボックスに変換
        returnedVal = IsCommnetSelect
    Case "B324" 'テキストボックスをコメントに戻す
        returnedVal = IsSelectCommentTextbox
    Case "B325" 'コメントを入力規則に変換
        returnedVal = IsCommnetSelect
    Case "B326" '入力規則をコメントに変換
        returnedVal = Not (GetInputRules(True) Is Nothing)
    Case "B632" '選択セルの色を取得
        Select Case CheckSelection()
        Case E_Range
            returnedVal = IsOnlyColor()
        Case E_Shape
            returnedVal = IsOnlyColor()
        Case Else
            returnedVal = False
        End Select
    Case Else
        returnedVal = True
    End Select
End Sub

'*****************************************************************************
'[概要] 選択されているセル(図形)が同一色で、塗りつぶしありかどうか
'[引数] なし
'[戻値] True:条件を満たすとき
'*****************************************************************************
Private Function IsOnlyColor() As Boolean
On Error GoTo ErrHandle
    IsOnlyColor = (GetColor(Selection.Interior) <> xlNone)
    Exit Function
ErrHandle:
    IsOnlyColor = False
End Function

'*****************************************************************************
'[概要] Interiorオブジェクトの色を取得
'[引数] Interiorオブジェクト
'[戻値] 色、xlNoneの時は色なし
'*****************************************************************************
Private Function GetColor(ByRef objInterior As Interior) As Long
    If IsNull(Selection.Interior.ColorIndex) Then
        GetColor = xlNone
        Exit Function
    End If
    Select Case objInterior.ColorIndex
    Case xlNone, xlAutomatic
        GetColor = xlNone
    Case Else
        GetColor = objInterior.Color
    End Select
End Function

'*****************************************************************************
'[概要] コメントが選択されているかどうか
'[引数] なし
'[戻値] True:コメントが選択されている
'*****************************************************************************
Private Function IsCommnetSelect() As Boolean
    Select Case CheckSelection()
    Case E_Range
        IsCommnetSelect = Not (GetComments Is Nothing)
    Case E_Shape
        IsCommnetSelect = (Selection.ShapeRange.Type = msoComment)
    End Select
End Function

'*****************************************************************************
'[イベント] getShowLabel
'*****************************************************************************
Private Sub getShowLabel(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = (GetValue(Control.ID, "ShowLabel") = 1)
End Sub

'*****************************************************************************
'[イベント] getLabel
'*****************************************************************************
Private Sub getLabel(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetValue(Control.ID, "Label")
    Select Case Control.ID
    Case "B5311"
        returnedVal = Replace(returnedVal, "{FONTNAME}", GetSetting(REGKEY, "KEY", "FontName", DEFAULTFONT))
    Case "B632"
'        returnedVal = "選択セル(図形)の色を取得"
        If CheckSelection() = E_Shape Then
            returnedVal = "選択図形の色を取得"
        Else
            returnedVal = "選択セルの色を取得"
        End If
    End Select
End Sub

'*****************************************************************************
'[イベント] getScreentip
'*****************************************************************************
Private Sub getScreentip(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetValue(Control.ID, "Screentip")
    Select Case Control.ID
    Case "B5311"
        returnedVal = Replace(returnedVal, "{FONTNAME}", GetSetting(REGKEY, "KEY", "FontName", DEFAULTFONT))
    Case "B621"
        returnedVal = Replace(returnedVal, "{COLOR}", GetColorStr(FColor(E_BMarkColor)))
    Case "B631"
        returnedVal = Replace(returnedVal, "{COLOR}", Trim(GetColorStr(FColor(E_FillColor)) & " " & GetColorHex(FColor(E_FillColor))))
    End Select
End Sub

'*****************************************************************************
'[イベント] getSupertip
'*****************************************************************************
Private Sub getSupertip(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetValue(Control.ID, "Supertip")
    Select Case Control.ID
    Case "B5311"
        returnedVal = Replace(returnedVal, "{FONTNAME}", GetSetting(REGKEY, "KEY", "FontName", DEFAULTFONT))
    End Select
End Sub

'*****************************************************************************
'[概要] RGB値を色を表現する文字に変換
'[引数] RGB値
'[戻値] 例：赤、ColorIndexにない値の時は16進数「#FFCC00」
'*****************************************************************************
Private Function GetColorStr(ByVal lngRGB As Long) As String
    Dim lngColor As Long
    Dim i As Long
    For i = 1 To 46
        lngColor = ThisWorkbook.Worksheets("Color").Range("D2:D47").Cells(i, 1).Value
        If lngRGB = lngColor Then
            GetColorStr = ThisWorkbook.Worksheets("Color").Range("C2:C47").Cells(i, 1).Value
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[概要] RGB値を色を表現する16進表現を取得
'[引数] RGB値
'[戻値] 例：#FFCC00
'*****************************************************************************
Private Function GetColorHex(ByVal lngRGB As Long) As String
    GetColorHex = "#" & WorksheetFunction.Dec2Hex(BGR2RGB(lngRGB), 6)
End Function

'*****************************************************************************
'[イベント] getShowImage
'*****************************************************************************
Private Sub getShowImage(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = (GetValue(Control.ID, "ShowImage") = 1)
End Sub

'*****************************************************************************
'[イベント] getImage
'*****************************************************************************
Private Sub getImage(Control As IRibbonControl, ByRef returnedVal)
    Select Case Control.ID
    Case "B631"
        Set returnedVal = CreateColorImage()
        Exit Sub
    End Select
    
    Dim Str As String
    Str = GetValue(Control.ID, "ImageMso")
    If Str <> "" Then
        returnedVal = Str
        Exit Sub
    End If
    
    Str = GetValue(Control.ID, "ImageFile")
    If Str <> "" Then
        Set returnedVal = GetImageFromResource(Str)
    End If
End Sub

'*****************************************************************************
'[概要] 塗りつぶしボタンのアイコンを作成する
'[引数] なし
'[戻値] IPicture
'*****************************************************************************
Private Function CreateColorImage() As IPicture
    Dim ICONSIZE As Long
    
    'DPIから 16Pixel or 20Pixelのアイコンを使用するか判定 ※Windowsの標準は96DPI
    Select Case GetDPI()
    Case Is < 120
        ICONSIZE = 16
    Case 120
        ICONSIZE = 20
    Case Else
        ICONSIZE = 32
    End Select
        
    Dim lngColor As Long
    ReDim Pixels(1 To ICONSIZE, 1 To ICONSIZE) As Long
    Dim x As Long, y As Long
    lngColor = BGR2RGB(FColor(E_FillColor)) + &HFF000000  '該当色 + α(不透明)
    For y = 2 To ICONSIZE - 1
        For x = 2 To ICONSIZE - 1
            If (x = 2) Or (x = ICONSIZE - 1) Or (y = 2) Or (y = ICONSIZE - 1) Then
                Pixels(y, x) = &HFFC0C0C0 '囲い(25%灰色)
            Else
                Pixels(y, x) = lngColor
            End If
        Next
    Next

    Dim objGdip As New CGdiplus
    Call objGdip.CreateFromPixels(Pixels())
    Set CreateColorImage = objGdip.ToIPicture
ErrHandle:
End Function

'*****************************************************************************
'[イベント] getSize
'*****************************************************************************
Private Sub getSize(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetValue(Control.ID, "ButtonSize")
End Sub

'*****************************************************************************
'[イベント] getContent 動的にメニューを作成する
'*****************************************************************************
Private Sub getContent(Control As IRibbonControl, ByRef returnedVal)
On Error Resume Next
    Select Case Control.ID
    Case "M31"
        returnedVal = GetRangeText(ThisWorkbook.Worksheets("dynamicMenu").Range("A1:A21"))
    Case "M32"
        returnedVal = GetRangeText(ThisWorkbook.Worksheets("dynamicMenu").Range("A23:A32"))
    Case "M51"
        returnedVal = GetRangeText(ThisWorkbook.Worksheets("dynamicMenu").Range("A34:A49"))
    Case "M52"
        returnedVal = GetRangeText(ThisWorkbook.Worksheets("dynamicMenu").Range("A51:A56"))
    Case "M53"
        returnedVal = GetRangeText(ThisWorkbook.Worksheets("dynamicMenu").Range("A58:A80"))
    Case "M71"
        returnedVal = GetRangeText(ThisWorkbook.Worksheets("dynamicMenu").Range("A82:A86"))
    End Select
End Sub

'*****************************************************************************
'[イベント] getPressed
'*****************************************************************************
Private Sub getPressed(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = False
    Select Case Control.ID
    Case "C1"
        On Error Resume Next
        returnedVal = (ActiveWorkbook.DisplayDrawingObjects = xlHide)
    Case "C2"
        returnedVal = GetTmpControl("C2").State
    End Select
End Sub

'*****************************************************************************
'[イベント] onCheckAction
'*****************************************************************************
Private Sub onCheckAction(Control As IRibbonControl, pressed As Boolean)
    Select Case Control.ID
    Case "C1"
        Call HideShapes(pressed)
    Case "C2"
        'チェック状態を保存
        GetTmpControl("C2").State = pressed
    End Select
End Sub

'*****************************************************************************
'[概要] 図形を表示・非表示をトグルさせる
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub ToggleHideShapes()
On Error GoTo ErrHandle
    Call HideShapes(ActiveWorkbook.DisplayDrawingObjects <> xlHide)
    Call GetRibbonUI.InvalidateControl("C1")
ErrHandle:
End Sub

'*****************************************************************************
'[イベント] onAction
'*****************************************************************************
Private Sub onAction(Control As IRibbonControl)
    'キーの再定義 Excelのバグ?でキーが無効になることがあるため
    Call SetKeys
    Call SetChkBox
    
    FCommand = GetValue(Control.ID, "Action")
    FParam = GetValue(Control.ID, "Parameter")
    FPressKey = IIf(GetKeyState(vbKeyShift) < 0, 1, 0) + IIf(GetKeyState(vbKeyControl) < 0, 2, 0)
    
On Error GoTo ErrHandle
    If FParam <> "" Then
        Call Application.Run(FCommand, FParam)
    Else
        Call Application.Run(FCommand)
    End If
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] チェックボックスのチェックを設定する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SetChkBox()
    Call GetRibbonUI.InvalidateControl("C1")
End Sub

'*****************************************************************************
'[概要] Commandsシートから該当の値を取得する
'[引数] コントロールId、項目名
'[戻値] IPicture
'*****************************************************************************
Private Function GetValue(ByVal ID As String, ByVal strCol As String) As Variant
    Dim x As Long
    Dim y As Long
    
    GetValue = ""
    
    If VarType(FValues) = vbEmpty Then
        FValues = ThisWorkbook.Worksheets("Commands").UsedRange
    End If
        
    '列数LOOP
    For x = 1 To UBound(FValues, 2)
        If FValues(1, x) = strCol Then
            '行数LOOP
            For y = 2 To UBound(FValues, 1)
                If FValues(y, 3) = ID Then
                    GetValue = FValues(y, x)
                    Exit Function
                End If
            Next
        End If
    Next

'    Dim objRange As Range
'    Set objRange = ThisWorkbook.Worksheets("Commands").UsedRange
'
'    '列数LOOP
'    For x = 1 To objRange.Columns.Count
'        If objRange.Cells(1, x).Value = strCol Then
'            Exit For
'        End If
'    Next
'
'    '行数LOOP
'    For y = 2 To objRange.Rows.Count
'        If objRange.Cells(y, "C").Value = Id Then
'            Exit For
'        End If
'    Next

'    GetValue = objRange.Cells(y, x).Value
End Function

'*****************************************************************************
'[概要] セルのデータからアイコンファイルを読込む
'[引数] Resourceシートののデータのファイル名(A列の値)
'[戻値] IPicture
'*****************************************************************************
Private Function GetImageFromResource(ByVal strImageFile As String) As IPicture
    Dim y As Long
    Dim i As Long
    Dim objRange As Range
    Dim objRow As Range
    Set objRange = ThisWorkbook.Worksheets("Resource").UsedRange
    
    Dim strPixel As String
    Dim strImageFile2 As String
    'DPIから 16Pixel or 20Pixelのアイコンを使用するか判定 ※Windowsの標準は96DPI
    Select Case GetDPI()
    Case Is < 120
        '例：Sample.png → Sample16.png
        strImageFile2 = Replace(strImageFile, ".png", 16 & ".png")
    Case 120
        '例：Sample.png → Sample20.png
        strImageFile2 = Replace(strImageFile, ".png", 20 & ".png")
    Case Else
        strImageFile2 = strImageFile
    End Select
        
    '行数LOOP
    For y = 1 To objRange.Rows.Count
        If objRange.Cells(y, "A").Value = strImageFile Then
            For i = 0 To 2
                If objRange.Cells(y + i, "A").Value = strImageFile2 Then
                    Set objRow = objRange.Rows(y + i)
                    Exit For
                End If
            Next
            If objRow Is Nothing Then
                Set objRow = objRange.Rows(y)
            End If
            Exit For
        End If
    Next
    
    If objRow Is Nothing Then Exit Function
On Error GoTo ErrHandle
    Set GetImageFromResource = LoadImageFromResource(objRow)
ErrHandle:
End Function

'*****************************************************************************
'[概要] セルのデータからアイコンファイルを読込む
'[引数] データを取得する行(Rangeオブジェクト)
'[戻値] なし
'*****************************************************************************
Private Function LoadImageFromResource(ByRef objRow As Range) As IPicture
    'ファイルサイズの配列を作成
    ReDim Data(1 To objRow.Cells(1, 1).End(xlToRight).Column - 1) As Byte
    Dim x As Long
    For x = 1 To UBound(Data)
         Data(x) = objRow.Cells(1, x + 1)
    Next
    
    Dim Gdip As New CGdiplus
    Call Gdip.CreateFromHGlobal(VarPtr(Data(1)))
    Set LoadImageFromResource = Gdip.ToIPicture
End Function

'*****************************************************************************
'[概要] アクティブセルにファイルを読込む(開発用)
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Sub LoadBinaryFile(ByVal strFilename As String, ByRef objCell As Range)
On Error GoTo ErrHandle
    ReDim Data(1 To FileLen(strFilename)) As Byte

    Open strFilename For Binary Access Read As #1
    Get #1, , Data
    Close #1

    Dim x As Long
    For x = 1 To UBound(Data)
        objCell.Cells(1, x + 1) = Data(x)
    Next
    objCell = strFilename
ErrHandle:
End Sub

'*****************************************************************************
'[イベント] getItemWidth
'*****************************************************************************
Sub getItemWidth(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = C_ICONSIZE
End Sub

'*****************************************************************************
'[イベント] getItemHeight
'*****************************************************************************
Sub getItemHeight(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = C_ICONSIZE
End Sub

'*****************************************************************************
'[イベント] getSelectedItemID
'*****************************************************************************
'Sub getSelectedItemID(Control As IRibbonControl, ByRef returnedVal)
'End Sub

'*****************************************************************************
'[イベント] getSelectedItemIndex
'*****************************************************************************
Sub getSelectedItemIndex(Control As IRibbonControl, ByRef returnedVal)
    Select Case Control.ID
    Case "G621"
        Dim lngColor As Long
        Dim i As Long
        For i = 0 To 39
            lngColor = ThisWorkbook.Worksheets("Color").Range("D2:D47").Cells(i + 1, 1).Value
            If lngColor = FColor(E_BMarkColor) Then
                returnedVal = i
                Exit Sub
            End If
        Next
    Case "G62"
'        Dim lngColorIndex As Long
'        Dim objInterior As Interior
'        Set objInterior = Selection.Interior
'
'        If VarType(objInterior.ColorIndex) <> vbNull Then
'            Dim i As Long
'            For i = 1 To 40
'                If objInterior.ColorIndex = ThisWorkbook.Worksheets("Color").Range("B2:B47").Cells(i, 1).Value Then
'                    If objInterior.Color - ActiveWorkbook.Colors(objInterior.ColorIndex) = 0 Then
'                        index = i - 1
'                        Exit Sub
'                    End If
'                End If
'            Next
'        End If
'        index = -1
'        Application.WarnOnFunctionNameConflict = False
'        Application.GenerateTableRefs = xlGenerateTableRefStruct
    End Select
'    Debug.Print index
End Sub

'*****************************************************************************
'[イベント] getItemCount
'*****************************************************************************
Sub getItemCount(Control As IRibbonControl, ByRef returnedVal)
    If FIcons(0) Is Nothing Then
        Dim lngColorIndex As Long
        Dim i As Long
        For i = 0 To 45
            lngColorIndex = ThisWorkbook.Worksheets("Color").Cells(i + 2, 2)
            Set FIcons(i) = GetColorPicture(ActiveWorkbook.Colors(lngColorIndex))
        Next
        '無効用の透明アイコンの作成
        Dim Pixels(1 To C_ICONSIZE, 1 To C_ICONSIZE) As Long
        Dim objGdip As New CGdiplus
        Call objGdip.CreateFromPixels(Pixels())
        Set FIcons(46) = objGdip.ToIPicture
        Set FIcons(47) = objGdip.ToIPicture
    End If
    
    Select Case Control.ID
    Case "G621"
        returnedVal = 46
    Case "G631"
        If FPickupColors Is Nothing Then
            returnedVal = 46
        Else
            returnedVal = 48 + FPickupColors.Count
        End If
        Call GetRibbonUI.InvalidateControl("B632")
    End Select
End Sub

'*****************************************************************************
'[イベント] getItemID
'*****************************************************************************
Sub getItemID(Control As IRibbonControl, Index As Integer, ByRef returnedVal)
    Dim lngColorIndex As Long
    lngColorIndex = ThisWorkbook.Worksheets("Color").Cells(Index + 2, 2)
    returnedVal = Control.ID & "_I" & Format(lngColorIndex, "00") 'IDは重複してはいけない
End Sub

'*****************************************************************************
'[イベント] getItemSupertip
'*****************************************************************************
Sub getItemSupertip(Control As IRibbonControl, Index As Integer, ByRef returnedVal)
    Select Case Control.ID
    Case "G621"
        With ThisWorkbook.Worksheets("Color").Rows(Index + 2)
            returnedVal = .Columns(3)
        End With
    Case "G631"
        Select Case Index
        Case 0 To 45
            With ThisWorkbook.Worksheets("Color").Rows(Index + 2)
                returnedVal = .Columns(3) & " #" & Mid(.Columns(4), 7, 2) _
                                                 & Mid(.Columns(4), 5, 2) _
                                                 & Mid(.Columns(4), 3, 2)
            End With
        Case 46, 47
            returnedVal = "無効"
        Case Else
            Dim Keys() As Variant
            Keys = FPickupColors.Keys
            returnedVal = GetColorHex(Keys(Index - 48))
        End Select
    End Select
End Sub

'*****************************************************************************
'[イベント] getItemImage
'*****************************************************************************
Sub getItemImage(Control As IRibbonControl, Index As Integer, ByRef returnedVal)
    Select Case Index
    Case 0 To 47
        Set returnedVal = FIcons(Index)
    Case Else
        Dim Items() As Variant
        Items = FPickupColors.Items
        Set returnedVal = Items(Index - 48)
    End Select
End Sub

'*****************************************************************************
'[概要] galleryアイテム用のカラーのイメージを動的に作成する
'[引数] RGBカラー(例：&HFF0000)
'[戻値] IPicture
'*****************************************************************************
Public Function GetColorPicture(ByVal lngColor As Long) As IPicture
    Dim Pixels(1 To C_ICONSIZE, 1 To C_ICONSIZE) As Long
    Dim x As Long, y As Long

    lngColor = BGR2RGB(lngColor) + &HFF000000 '該当色 + α(不透明)
    For y = 2 To C_ICONSIZE - 1
        For x = 2 To C_ICONSIZE - 1
            If (x = 2) Or (x = C_ICONSIZE - 1) Or (y = 2) Or (y = C_ICONSIZE - 1) Then
                Pixels(y, x) = &HFF808080 '囲い(50%灰色)
            Else
                Pixels(y, x) = lngColor
            End If
        Next
    Next

    Dim objGdip As New CGdiplus
    Call objGdip.CreateFromPixels(Pixels())
    Set GetColorPicture = objGdip.ToIPicture
End Function

'*****************************************************************************
'[概要] BGR -> RGB に変換
'[引数] BGRカラー(例：&HFF0000)
'[戻値] RGBカラー(例：&H0000FF)
'*****************************************************************************
Private Function BGR2RGB(ByVal lngColor As Long) As Long
    Dim strBGR As String
    strBGR = WorksheetFunction.Dec2Hex(lngColor, 6)
    Dim R As String, G As String, B As String
    B = Mid(strBGR, 1, 2)
    G = Mid(strBGR, 3, 2)
    R = Mid(strBGR, 5, 2)
    BGR2RGB = "&H" & R & G & B
End Function

''*****************************************************************************
''[イベント] getItemLabel
''*****************************************************************************
'Sub getItemLabel(Control As IRibbonControl, index As Integer, ByRef returnedVal)
'    returnedVal = ""
'End Sub
'
''*****************************************************************************
''[イベント] getItemScreentip
''*****************************************************************************
'Sub getItemScreentip(Control As IRibbonControl, index As Integer, ByRef returnedVal)
'    returnedVal = ""
'End Sub

'*****************************************************************************
'[イベント] galleryのアイテムをクリックした時
'*****************************************************************************
Sub gallery_onAction(Control As IRibbonControl, itemID As String, Index As Integer)
    Select Case Control.ID
    Case "G621"
        GetTmpControl("BMarkColor").Parameter = ThisWorkbook.Worksheets("Color").Range("D2:D47").Cells(Index + 1, 1).Value
        If TypeOf Selection Is Range Then
            With Selection.Interior
                .Color = FColor(E_BMarkColor)
                .Pattern = xlSolid
                .PatternColor = C_PatternColor
            End With
        End If
        Call GetRibbonUI.InvalidateControl("B621")
        Call GetRibbonUI.InvalidateControl("C2")
    Case "G631"
        Select Case Index
        Case 0 To 45
            GetTmpControl("FillColor").Parameter = ThisWorkbook.Worksheets("Color").Range("D2:D47").Cells(Index + 1, 1).Value
            Call FillColor
        Case 46, 47
        Case Else
            Dim Keys() As Variant
            Keys = FPickupColors.Keys
            GetTmpControl("FillColor").Parameter = Keys(Index - 48)
            Call FillColor
        End Select
    End Select
End Sub

'*****************************************************************************
'[概要] 選択セル(または図形)の色を取得
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub PickupColor()
    GetTmpControl("FillColor").Parameter = Selection.Interior.Color
    Call GetRibbonUI.InvalidateControl("B631")
    
    Dim i As Long
    Dim lngColor As Long
    Dim lngFillColor As Long
    lngFillColor = FColor(E_FillColor)
    For i = 1 To 46
        lngColor = ThisWorkbook.Worksheets("Color").Range("D2:D47").Cells(i, 1).Value
        If lngFillColor = lngColor Then
            Exit Sub
        End If
    Next
    
    If FPickupColors Is Nothing Then
        Set FPickupColors = CreateObject("Scripting.Dictionary")
    End If
    
    If FPickupColors.Exists(lngFillColor) Then
        Exit Sub
    End If
    
    Call FPickupColors.Add(lngFillColor, GetColorPicture(lngFillColor))
End Sub

'*****************************************************************************
'[概要] リボンのコールバック関数を実行する(開発用)
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub onAction2(Control As IRibbonControl)
    If C_DEBUG Then
        Select Case Control.ID
        Case "Bdmy1"
            FValues = Empty
            Call GetRibbonUI.Invalidate
        Case "Bdmy2"
            ThisWorkbook.IsAddin = Not ThisWorkbook.IsAddin
        Case "Bdmy3"
            Call ThisWorkbook.Save
        End Select
    End If
End Sub

'*****************************************************************************
'[概要] ResourceシートのPngファイルをフォルダに書き込む(開発用)
'[引数] 書込むフォルダ
'[戻値] なし
'*****************************************************************************
Public Sub SavePngFiles(ByVal strFolder As String)
    Dim y As Long
    Dim objRange As Range
    Dim objRow As Range
    Set objRange = ThisWorkbook.Worksheets("Resource").UsedRange
    
    Dim strFilename As String
    
    '行数LOOP
    For y = 1 To objRange.Rows.Count
        'ファイルサイズの配列を作成
        ReDim Data(1 To objRange.Cells(y, 1).End(xlToRight).Column - 1) As Byte
        Dim x As Long
        For x = 1 To UBound(Data)
            Data(x) = objRange.Cells(y, x + 1)
        Next
        
        strFilename = strFolder & "\" & objRange.Cells(y, "A").Value
        Open strFilename For Binary Access Write As #1
        Put #1, , Data
        Close #1
    Next
End Sub

'Sub LoadIcon()
'    Call LoadBinaryFile("FindNext.png", ThisWorkbook.Worksheets("Resource").Range("A70"))
'End Sub


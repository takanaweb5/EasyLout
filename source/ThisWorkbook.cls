VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim blnIsAddin As Boolean

'*****************************************************************************
'[イベント]　AddinInstall
'[ 概  要 ]　アドインのインストール時に実行される
'*****************************************************************************
Private Sub Workbook_AddinInstall()
On Error GoTo ErrHandle

    'ＰＣによってはアイコンが表示されないことがあるため
    Call Sleep(1000)
    MsgBox "Workbook_AddinInstall"
    
    'かんたんレイアウトのインストール
    Call InstToolBar
    'かんたんレイアウト2のインストール
    Call InstTool2Bar
    
    blnIsAddin = True
    Application.CommandBars("かんたんレイアウト").Enabled = True
    Application.CommandBars("かんたんレイアウト2").Enabled = True
    Application.CommandBars("かんたんレイアウトアイコン").Enabled = False
Exit Sub
ErrHandle:
'    Call MsgBox("インストールに失敗しました", vbCritical)
End Sub

'*****************************************************************************
'[ 関数名 ]　InstToolBar
'[ 概  要 ]　かんたんレイアウトバーのインストール
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub InstToolBar()
    Dim objCmdBar    As CommandBar 'かんたんレイアウト2バー
    Dim i As Long
    
    On Error Resume Next
    Call Application.CommandBars("かんたんレイアウト").Delete
    On Error GoTo ErrHandle
    
    Set objCmdBar = Application.CommandBars.Add("かんたんレイアウト")
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "列"
        .OnAction = "ColRowChange"
        .TooltipText = "行と列を切替える"
        .Style = msoButtonCaption
        .Tag = "{S+F2}{C+S+C}{C+S+V}"
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "縮小"
        .BeginGroup = True
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "拡大"
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "前移動"
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "後移動"
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "サイズ一覧"
        .BeginGroup = True
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "等間隔"
    End With
        
    With objCmdBar.Controls.Add(msoControlButton, 882)
        .Caption = "自動調整"
        .Style = msoButtonIcon
    End With
        
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "分割"
        .BeginGroup = True
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "消去"
    End With
    
    With objCmdBar.Controls.Add(msoControlButton, 1742)
        .Caption = "結合"
        .BeginGroup = True
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "セル境界前移動"
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "セル境界後移動"
    End With
    
    'アイコンやTooltipTextの設定
    On Error Resume Next
    With objCmdBar.Controls
        For i = 1 To .Count
            Call SetCommand("列", .Item(i))
        Next i
    End With
    On Error GoTo ErrHandle
    
    objCmdBar.Visible = True
Exit Sub
ErrHandle:
    Call MsgBox("インストールに失敗しました", vbCritical)
End Sub

'*****************************************************************************
'[ 関数名 ]　InstElseBar
'[ 概  要 ]　かんたんレイアウト2バーのインストール
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub InstTool2Bar()
    Dim objCmdBar    As CommandBar 'かんたんレイアウト2バー
    Dim i As Long
    
    On Error Resume Next
    Call Application.CommandBars("かんたんレイアウト2").Delete
    On Error GoTo ErrHandle
    
    Set objCmdBar = Application.CommandBars.Add("かんたんレイアウト2")
    
    With objCmdBar.Controls.Add(msoControlButton, 849)
        .Caption = "数式バー"
    End With
    
    With objCmdBar.Controls.Add(msoControlButton, 485)
        .Caption = "グリッドの表示"
    End With
    
    With objCmdBar.Controls.Add(msoControlButton, 549)
        .Caption = "グリッドに合せる"
    End With
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "領域移動"
        .BeginGroup = True
    End With
        
    On Error Resume Next
    With objCmdBar.Controls.Add(msoControlButton, 5837)
        .Caption = "罫線を除くすべてを貼り付け"
        .Style = msoButtonIcon
    End With
    On Error GoTo ErrHandle
        
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "値の貼り付け"
    End With
        
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "値を保持して結合"
        .BeginGroup = True
    End With
        
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "1セル単位に分割"
    End With
        
    With objCmdBar.Controls.Add(msoControlButton, 800)
        .Caption = "セル結合の解除"
    End With
        
    With objCmdBar.Controls.Add(msoControlPopup)
        .Caption = "その他"
        .BeginGroup = True
        .OnAction = "OnElseMenuClick"
    End With
        
    With objCmdBar.Controls("その他").Controls.Add(msoControlButton)
        .Caption = "｢かんたんレイアウト｣ヘルプ"
    End With
        
    With objCmdBar.Controls("その他").Controls.Add(msoControlButton)
        .Caption = "テキストボックスをセルに変換"
    End With
        
    With objCmdBar.Controls("その他").Controls.Add(msoControlButton)
        .Caption = "セルをテキストボックスに変換"
    End With
    
    With objCmdBar.Controls("その他").Controls.Add(msoControlButton)
        .Caption = "図形をグリッドに合せる"
    End With
        
    With objCmdBar.Controls("その他").Controls.Add(msoControlPopup)
        .Caption = "文字種の変換"
        .BeginGroup = True
        
        With .Controls.Add(msoControlButton)
            .Caption = "大文字に変換"
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "小文字に変換"
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "先頭のみ大文字に変換"
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "ひらがなに変換"
            .BeginGroup = True
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "カタカナに変換"
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "全角に変換"
            .BeginGroup = True
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "半角に変換"
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "カタカナのみ全角に変換"
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "カタカナ以外半角に変換"
        End With
        With .Controls.Add(msoControlButton)
            .Caption = "先頭と末尾の空白を削除"
            .BeginGroup = True
        End With
    End With
    
    With objCmdBar.Controls("その他").Controls.Add(msoControlButton)
        .Caption = "選択領域の一部取消し"
        .BeginGroup = True
    End With
        
    With objCmdBar.Controls("その他").Controls.Add(msoControlButton)
        .Caption = "入力補助画面"
        If Val(Application.Version) < 12 Then
            .ShortcutText = "Shift+F2"
        End If
    End With
        
    With objCmdBar.Controls("その他").Controls.Add(msoControlButton)
        .Caption = "図形非表示"
        If Val(Application.Version) < 12 Then
            .ShortcutText = "Ctrl+6"
        End If
    End With
        
    'アイコンやTooltipTextの設定
    On Error Resume Next
    With objCmdBar.Controls
        For i = 1 To .Count
            Call SetCommand("ツール", .Item(i))
        Next i
    End With
    With objCmdBar.Controls("その他").Controls
        For i = 1 To .Count
            Call SetCommand("その他", .Item(i))
        Next i
    End With
    With objCmdBar.Controls("その他").Controls("文字種の変換").Controls
        For i = 1 To .Count
            Call SetCommand("文字種変換", .Item(i))
        Next i
    End With
    On Error GoTo ErrHandle
    
    objCmdBar.Visible = True
Exit Sub
ErrHandle:
    Call MsgBox("インストールに失敗しました", vbCritical)
End Sub

'*****************************************************************************
'[イベント]　AddinUninstall
'[ 概  要 ]　アドインのアンインストール時に実行される
'*****************************************************************************
Private Sub Workbook_AddinUninstall()
    On Error Resume Next
    Call Application.CommandBars("かんたんレイアウト").Delete
    Call Application.CommandBars("かんたんレイアウト2").Delete
    Call Application.CommandBars("かんたんレイアウトアイコン").Delete
End Sub

Private Sub Auto_Remove()
End Sub

'*****************************************************************************
'[イベント]　Workbook_Open
'[ 概  要 ]　開く時
'*****************************************************************************
Private Sub Workbook_Open()
    Dim strMsg As String
    Dim i As Long
    
    Application.CommandBars("かんたんレイアウトアイコン").Enabled = False
    Call SetDPIRatio
    
    If blnIsAddin = True Then
        Call Application.OnKey("+{F2}", "OpenEdit")
        Call Application.OnKey("+^{c}", "CopyText")
        Call Application.OnKey("+^{v}", "PasteText")
        Exit Sub
    End If
    
'    'ＰＣによってはアイコンが表示されないことがあるため
'    Call Sleep(1000)
'
'    'かんたんレイアウトのインストール
'    Call InstToolBar
'    'かんたんレイアウト2のインストール
'    Call InstTool2Bar
'
'    Application.CommandBars("かんたんレイアウト").Enabled = True
'    Application.CommandBars("かんたんレイアウト2").Enabled = True
'    Application.CommandBars("かんたんレイアウトアイコン").Enabled = False
    
    For i = 1 To AddIns.Count
        If UCase(AddIns(i).Name) = UCase("EasyLout.xla") Then
            If AddIns(i).Installed = True Then
                Call SetKeys
'                Call Application.OnTime(Now(), "LoadThisWorkbook")
                Exit Sub
            End If
        End If
    Next i
    
    strMsg = "このファイルを開くことは出来ません。" & vbCrLf
    strMsg = strMsg & "アドインに登録して使用して下さい。" & vbCrLf & vbCrLf
    strMsg = strMsg & "登録方法はヘルプファイルをご覧下さい。"
    Call MsgBox(strMsg, vbExclamation, ThisWorkbook.Name)
    Call ThisWorkbook.Close
End Sub

'*****************************************************************************
'[イベント]　Workbook_BeforeClose
'[ 概  要 ]　閉じる時
'*****************************************************************************
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Call Application.OnKey("+{F2}")
    Call Application.OnKey("+^{c}")
    Call Application.OnKey("+^{v}")
    Call Application.OnKey("{BS}")
End Sub


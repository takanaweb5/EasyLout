VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEdit 
   Caption         =   "かんたんレイアウト"
   ClientHeight    =   3825
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7368
   OleObjectBlob   =   "frmEdit.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hWnd       As Long
Private OrgWndProc As Long
Private blnZoomed  As Boolean
Private objTmpBar  As CommandBar
Private dblAnchor  As Double

'*****************************************************************************
'[イベント]　UserForm_Initialize
'[ 概  要 ]　フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()
    Dim lngStyle As Long
    Dim i        As Long
    
    dblAnchor = Me.Height - cmdCancel.Top
    
    '********************************************
    'ウィンドウのサイズを変更出来るように変更
    '********************************************
    hWnd = FindWindow("ThunderDFrame", Me.Caption)
    lngStyle = GetWindowLong(hWnd, GWL_STYLE)
    If ActiveCell.HasFormula = True Then
        Call SetWindowLong(hWnd, GWL_STYLE, lngStyle Or WS_THICKFRAME)
    Else
        Call SetWindowLong(hWnd, GWL_STYLE, lngStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX)
    End If
    
    '********************************************
    'サブクラス化してマウスホイールを有効にする
    '********************************************
    OrgWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubClassProc)
        
    '********************************************
    'フォームの初期状態を設定
    '********************************************
    With txtEdit
        If ActiveCell.HasFormula = True Then
            .MultiLine = False
            .WordWrap = False
            .ScrollBars = fmScrollBarsNone
            .SelectionMargin = False
            .Text = ActiveCell.Formula
            chkWordWrap.Enabled = False
        Else
            .MultiLine = True
            .WordWrap = False
            .ScrollBars = fmScrollBarsBoth
            .SelectionMargin = True
            If IsOnlyCell(Selection) Then
                .Text = ActiveCell.Value
            Else
                .Text = GetRangeText(Selection)
            End If
            chkWordWrap = .WordWrap
        End If
    End With
    
    '********************************************
    '右クリックメニュー作成
    '********************************************
    Set objTmpBar = CommandBars.Add(Position:=msoBarPopup, Temporary:=True)
    With objTmpBar.Controls
        With .Add()
            .Caption = "元に戻す(&U)　　　Ctrl+Z"
        End With
        With .Add()
            .Caption = "やり直し(&F)　Ctrl+Shift+Z"
        End With
        With .Add(, 21)
            .BeginGroup = True
        End With
        With .Add(, 19)
        End With
        With .Add(, 22)
        End With
        With .Add()
            .Caption = "削除(&D)"
        End With
        With .Add()
            .BeginGroup = True
            .Caption = "すべて選択(&A)　　Ctrl+A"
        End With
        With .Add()
            .BeginGroup = True
            .Caption = "大文字に変換"
        End With
        With .Add()
            .Caption = "小文字に変換"
        End With
        With .Add()
            .Caption = "先頭のみ大文字に変換"
        End With
        With .Add()
            .Caption = "ひらがなに変換"
        End With
        With .Add()
            .Caption = "カタカナに変換"
        End With
        With .Add()
            .Caption = "全角に変換"
        End With
        With .Add()
            .Caption = "半角に変換"
        End With
        With .Add()
            .Caption = "カタカナ以外半角に変換"
        End With
    End With

    For i = 1 To objTmpBar.Controls.Count
        objTmpBar.Controls(i).OnAction = "OnPopupClick2"
        objTmpBar.Controls(i).Tag = i
    Next i
End Sub

'*****************************************************************************
'[イベント]　UserForm_Terminate
'[ 概  要 ]　デストラクタ
'*****************************************************************************
Private Sub UserForm_Terminate()
    'ウィンドウプロシジャーを元にもどす
    If OrgWndProc <> 0 Then
        Call SetWindowLong(hWnd, GWL_WNDPROC, OrgWndProc)
    End If
    
    '右クリックメニュー削除
    Call objTmpBar.Delete
End Sub

'*****************************************************************************
'[イベント]　UserForm_QueryClose
'[ 概  要 ]　×ボタンでフォームを閉じる時、フォームを破棄させない
'*****************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '×ボタンでフォームを閉じる時、フォームを破棄させない
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        blnZoomed = IsZoomed(hWnd)
        Me.Hide
    End If
    
    'ウィンドウのサイズを元に戻す
    Call ShowWindow(hWnd, SW_SHOWNORMAL)
End Sub
'*****************************************************************************
'[イベント]　cmdOK_Click
'[ 概  要 ]　ＯＫボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrHandle
    Dim strText         As String
    Dim objOldSelection As Range
    Dim objNewSelection As Range
    
    '改行のCRLF→LF
    strText = Replace$(txtEdit.Text, vbCr, "")
    
    If ActiveCell.HasFormula Or IsOnlyCell(Selection) Then
        Call SaveUndoInfo(E_CellValue, ActiveCell.MergeArea)
        ActiveCell.Value = Replace$(strText, vbTab, "")
        Call SetOnUndo
    Else
        Set objOldSelection = Selection
        Set objNewSelection = GetPasteRange(strText, Selection)
        Call SaveUndoInfo(E_CellValue, objOldSelection)
        Call objOldSelection.ClearContents
        Call objNewSelection.Select
        Call PasteTabText(strText, objNewSelection)
        Call SetOnUndo
    End If
ErrHandle:
    blnZoomed = IsZoomed(hWnd)
    Me.Hide
    'ウィンドウのサイズを元に戻す
    Call ShowWindow(hWnd, SW_SHOWNORMAL)
End Sub

'*****************************************************************************
'[イベント]　cmdCancel_Click
'[ 概  要 ]　キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    blnZoomed = IsZoomed(hWnd)
    Me.Hide
    'ウィンドウのサイズを元に戻す
    Call ShowWindow(hWnd, SW_SHOWNORMAL)
End Sub

'*****************************************************************************
'[イベント]　SpbSize_Change
'[ 概  要 ]　フォントサイズ変更時
'*****************************************************************************
Private Sub SpbSize_Change()
    txtSize.Text = CStr(SpbSize.Value)
    txtEdit.Font.Size = SpbSize.Value
End Sub

'*****************************************************************************
'[イベント]　SpbSize_KeyDown
'[ 概  要 ]　ESCキーでフォントサイズの変更を取消しさせないようにする
'*****************************************************************************
Private Sub SpbSize_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Call cmdCancel_Click
        Exit Sub
    End If

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        KeyCode = 0
        Call txtEdit.SetFocus
        Exit Sub
    End If
End Sub

'*****************************************************************************
'[イベント]　SpbSize_Exit
'[ 概  要 ]　フォントサイズ変更後の水平スクロールバーを表示するためのおまじない
'*****************************************************************************
Private Sub SpbSize_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call txtEdit.SetFocus
    Call Me.Repaint
End Sub

'*****************************************************************************
'[イベント]　chkWordWrap_Click
'[ 概  要 ]　右端の折り返しの有無を変更します
'*****************************************************************************
Private Sub chkWordWrap_Change()
    txtEdit.WordWrap = chkWordWrap
    If chkWordWrap Then
        txtEdit.ScrollBars = fmScrollBarsVertical
    Else
        txtEdit.ScrollBars = fmScrollBarsBoth
    End If
    Call txtEdit.SetFocus
    Call Me.Repaint
End Sub

'*****************************************************************************
'[イベント]　UserForm_Resize
'[ 概  要 ]　フォームのサイズ変更時
'*****************************************************************************
Private Sub UserForm_Resize()
    If Me.Width < 360 Then
        Me.Width = 360
    End If
    If Me.Height < 65 Or ActiveCell.HasFormula = True Then
        Me.Height = 65
    End If
    
    cmdCancel.Top = Me.Height - dblAnchor
    cmdCancel.Left = Me.Width - cmdCancel.Width - 10
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - 10 - cmdOK.Width
    txtEdit.Width = Me.Width - 6
    txtEdit.Height = cmdCancel.Top - 6
    frmFontSize.Top = cmdCancel.Top
    SpbSize.Top = cmdCancel.Top
    chkWordWrap.Top = cmdCancel.Top + 1
End Sub

'*****************************************************************************
'[イベント]　txtEdit_KeyDown
'[ 概  要 ]　Ctrl+Return または Alt+Return でＯＫボタン押下する
'*****************************************************************************
Private Sub txtEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '右クリックメニューを表示する
    If KeyCode = 93 Then
        Call txtEdit_MouseUp(2, 0, 0, 0)
        KeyCode = 0
        Exit Sub
    End If
    
    If Shift = 2 Or Shift = 4 Then
        If KeyCode = vbKeyReturn Then
            Call cmdOK_Click
            Exit Sub
        End If
    End If

    'Ctrl+Shift+Z でRedo
    If Shift = 3 And KeyCode = vbKeyZ Then
        Call Me.RedoAction
        KeyCode = 0
        Exit Sub
    End If
End Sub

'*****************************************************************************
'[イベント]　txtEdit_MouseUp
'[ 概  要 ]　右クリックメニューを表示する
'*****************************************************************************
Private Sub txtEdit_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then '右ボタン
        objTmpBar.Controls(1).Enabled = Me.CanUndo
        objTmpBar.Controls(2).Enabled = Me.CanRedo
        objTmpBar.Controls(3).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(4).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(5).Enabled = txtEdit.CanPaste
        objTmpBar.Controls(6).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(7).Enabled = (txtEdit.Text <> "")
        objTmpBar.Controls(8).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(9).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(10).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(11).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(12).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(13).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(14).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.Controls(15).Enabled = (txtEdit.SelLength > 0)
        objTmpBar.ShowPopup
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　OnPopupClick
'[ 概  要 ]　ポップアップメニュークリック時
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub OnPopupClick()
    Dim strNewSelText As String
On Error GoTo ErrHandle
    Select Case CommandBars.ActionControl.Tag
    Case 1 '元に戻す
        Call Me.UndoAction
    Case 2 'やり直し
        Call Me.RedoAction
    Case 3 '切り取り
        Call SendKeys("^x", True)
    Case 4 'コピー
        Call SendKeys("^c", True)
    Case 5 '貼り付け
        Call SendKeys("^v", True)
    Case 6 '削除
        Call SendKeys("{DEL}", True)
    Case 7 'すべて選択
        Call SendKeys("^a", True)
    Case 8  '大文字に変換
        strNewSelText = StrConvert(txtEdit.SelText, "UpperCase")
    Case 9  '小文字に変換
        strNewSelText = StrConvert(txtEdit.SelText, "LowerCase")
    Case 10 '先頭のみ大文字に変換
        strNewSelText = StrConvert(txtEdit.SelText, "ProperCase")
    Case 11 'ひらがなに変換
        strNewSelText = StrConvert(txtEdit.SelText, "Hiragana")
    Case 12 'カタカナに変換
        strNewSelText = StrConvert(txtEdit.SelText, "Katakana")
    Case 13 '全角に変換
        strNewSelText = StrConvert(txtEdit.SelText, "Wide")
    Case 14 '半角に変換
        strNewSelText = StrConvert(txtEdit.SelText, "Narrow")
    Case 15 'カタカナ以外半角に変換
        strNewSelText = StrConvert(txtEdit.SelText, "WideExceptKana")
    End Select
    
    'Undoができるようにクリップボードを使用する
    If strNewSelText <> "" Then
        Dim lngSelL As Long
        Dim lngSelR  As Long
        
        '半角カタカナの「ｶﾞ」などは文字数が変わるので注意
        lngSelL = txtEdit.SelStart - 1
        lngSelR = Len(txtEdit.Text) - (lngSelL + txtEdit.SelLength)
        Call SetClipbordText(strNewSelText)
        Call SendKeys("^v", True)
        txtEdit.SelStart = lngSelL + 1
        txtEdit.SelLength = Len(txtEdit.Text) - (lngSelL + lngSelR)
        
        'クリップボードのクリア
        Call ClearClipbord
    End If
ErrHandle:
End Sub

'*****************************************************************************
'[プロパティ]　Zoomed
'[ 概  要 ]　ウィンドウサイズが最大化されているか？
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get Zoomed() As Boolean
    Zoomed = blnZoomed
End Property
Public Property Let Zoomed(ByVal Value As Boolean)
    'ウィンドウサイズを最大化する
    If ActiveCell.HasFormula = False And Value = True Then
        Call ShowWindow(hWnd, SW_MAXIMIZE)
        Me.Hide
    End If
End Property

'*****************************************************************************
'[プロパティ]　WndProc
'[ 概  要 ]　ウィンドウプロシジャーのハンドル
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get WndProc() As Long
    WndProc = OrgWndProc
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSizeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private frmParent              As MSForms.Frame
Private lblNo                  As MSForms.Label
Attribute lblNo.VB_VarHelpID = -1
Private lblArea                As MSForms.Label
Private lblmm                  As MSForms.Label
Private lblPer                 As MSForms.Label
Private WithEvents chkSelect   As MSForms.CheckBox
Attribute chkSelect.VB_VarHelpID = -1
Private WithEvents txtSize     As MSForms.TextBox
Attribute txtSize.VB_VarHelpID = -1
Private WithEvents spbChgSize  As MSForms.SpinButton
Attribute spbChgSize.VB_VarHelpID = -1
Private WithEvents lblOther    As MSForms.Label
Attribute lblOther.VB_VarHelpID = -1

Private objColRows    As Object  '該当の列(行)
Private blnHidden()   As Boolean 'Trueの時、サイズ変更の対象外
Private dblRatio()    As Double

'*****************************************************************************
'[ 関数名 ]　CreateSizeList
'[ 概  要 ]  i:項番、objParent:フレーム
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub CreateSizeList(ByVal i As Long, ByRef objParent As MSForms.Frame)
    Dim lngTop As Long
    
    Set frmParent = objParent
    lngTop = (i - 1) * 21 + 2
    
    'No
    Set lblNo = objParent.Controls.Add("Forms.Label.1")
    With lblNo
        .Top = lngTop
        .Left = 6
        .Width = 20
        .Height = 18
        .TextAlign = fmTextAlignRight
        .Caption = CStr(i)
    End With
    
    '対象エリア
    Set lblArea = objParent.Controls.Add("Forms.Label.1")
    With lblArea
        .Top = lngTop
        .Left = 43
        .Width = 50
        .Height = 21
        .TextAlign = fmTextAlignLeft
    End With
    
    'テキストボックス
    Set txtSize = objParent.Controls.Add("Forms.TextBox.1")
    With txtSize
        .Top = lngTop - 2
        .Left = 92
        .Width = 40
        .Height = 18
        .TextAlign = fmTextAlignRight
        .SelectionMargin = False
        .Locked = True
        .TabStop = False
        .Name = "txtSize" & CStr(i)
    End With
    
    'スピンボタン
    Set spbChgSize = objParent.Controls.Add("Forms.SpinButton.1")
    With spbChgSize
        .Top = lngTop - 2
        .Left = 133
        .Width = 30
        .Height = 18
        .Min = 0
        .Max = 32767
        .Name = "spbChgSize" & CStr(i)
    End With
    
    'サイズ(mm)
    Set lblmm = objParent.Controls.Add("Forms.Label.1")
    With lblmm
        .Top = lngTop
        .Left = 164
        .Width = 49
        .Height = 18
        .TextAlign = fmTextAlignRight
    End With
    
    '割合(％)
    Set lblPer = objParent.Controls.Add("Forms.Label.1")
    With lblPer
        .Top = lngTop
        .Left = 213
        .Width = 43
        .Height = 18
        .TextAlign = fmTextAlignRight
    End With
    
    '選択
    Set chkSelect = objParent.Controls.Add("Forms.CheckBox.1")
    With chkSelect
        .Top = lngTop - 2
        .Left = 268
        .Width = 13
        .Height = 18
        .TabStop = False
        .Name = "chkSelect" & CStr(i)
        .Caption = ""
        .ControlTipText = frmSizeList.Controls("lblSelect").ControlTipText
        If objParent.Name = "frmPage1" Then
            .Enabled = False
        End If
    End With
    
    '備考
    Set lblOther = objParent.Controls.Add("Forms.Label.1")
    With lblOther
        .Top = lngTop
        .Left = 288
        .Width = 40
        .Height = 18
        .TextAlign = fmTextAlignRight
    End With
    
    objParent.ScrollHeight = lngTop + 16
End Sub

'*****************************************************************************
'[イベント]　Class_Terminate
'[ 概  要 ]　デストラクタ
'*****************************************************************************
Private Sub Class_Terminate()
    'コントロールをすべて削除する
    Call frmParent.Controls.Clear
End Sub

'*****************************************************************************
'[イベント]　lblOther_MouseMove
'[ 概  要 ]　詳細情報上でマウスカーソルを手の形にする
'*****************************************************************************
Private Sub lblOther_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If lblOther.Caption = "詳細情報" Then
        Call SetCursor(LoadCursor(0, IDC_HAND))
    End If
End Sub

'*****************************************************************************
'[イベント]　lblOther_Click
'[ 概  要 ]　詳細情報クリック時
'*****************************************************************************
Private Sub lblOther_Click()
    If lblOther.Caption <> "詳細情報" Then
        Exit Sub
    End If
    frmSizeList.Controls("tbsTab").Tag = lblNo
    frmSizeList.Controls("tbsTab").Value = 1
End Sub

'*****************************************************************************
'[イベント]　spbChgSize_KeyDown
'[ 概  要 ]　｢↓｣，｢↑｣キーで上下に移動するようにする
'*****************************************************************************
Private Sub spbChgSize_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
On Error GoTo ErrHandle
    Select Case (KeyCode)
    Case vbKeyLeft, vbKeyRight
        Exit Sub
    Case vbKeyUp
        Call frmParent.Controls("txtSize" & lblNo - 1).SetFocus
    Case vbKeyDown
        Call frmParent.Controls("txtSize" & lblNo + 1).SetFocus
    Case vbKeyEscape
        KeyCode = 0
    End Select
Exit Sub
ErrHandle:
    KeyCode = 0
End Sub

'*****************************************************************************
'[イベント]　txtSize_KeyDown
'[ 概  要 ]　｢→｣，｢←｣キーでサイズの変更を出来るようにする
'            ｢↓｣，｢↑｣キーで上下に移動するようにする
'*****************************************************************************
Private Sub txtSize_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    
    Select Case (KeyCode)
    Case vbKeyLeft, vbKeyRight
        Call spbChgSize.SetFocus
    Case vbKeyUp
        Call frmParent.Controls("txtSize" & lblNo - 1).SetFocus
    Case vbKeyDown
        Call frmParent.Controls("txtSize" & lblNo + 1).SetFocus
    Case vbKeyReturn
    Case Else
        If Shift = 0 Then
            Call MsgBox("カーソルキーで操作して下さい")
            Call spbChgSize.SetFocus
        End If
    End Select
End Sub

'*****************************************************************************
'[イベント]　chkSelect_KeyDown
'[ 概  要 ]　｢↓｣，｢↑｣キーで上下に移動するようにする
'*****************************************************************************
Private Sub chkSelect_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    
    Select Case (KeyCode)
    Case vbKeyLeft, vbKeyRight
        Call spbChgSize.SetFocus
    Case vbKeyUp
        Call frmParent.Controls("txtSize" & lblNo - 1).SetFocus
    Case vbKeyDown
        Call frmParent.Controls("txtSize" & lblNo + 1).SetFocus
    End Select
End Sub

'*****************************************************************************
'[イベント]　chkSelect_Change
'[ 概  要 ]　選択をチェックした時
'*****************************************************************************
Private Sub chkSelect_Change()
    Dim blnValue As Boolean
    
    blnValue = frmParent.Controls("chkSelect" & lblNo).Value
    
    '現在選択されている番号のチェックをはずす
    If blnValue = True Then
        If frmSizeList.SelectNo <> 0 Then
            frmParent.Controls("spbChgSize" & frmSizeList.SelectNo).Enabled = True
            frmParent.Controls("chkSelect" & frmSizeList.SelectNo).Value = False
        End If
    End If
    
    'チェックされた番号のスピンボタンを操作出来なくする
    frmParent.Controls("spbChgSize" & lblNo).Enabled = Not blnValue
    
    '選択されている番号をフォームに通知する
    If blnValue = True Then
        frmSizeList.SelectNo = lblNo
    Else
        frmSizeList.SelectNo = 0
    End If
End Sub

'*****************************************************************************
'[イベント]　spbChgSize_Change
'[ 概  要 ]　サイズ変更時
'*****************************************************************************
Private Sub spbChgSize_Change()
    Dim lngIndex As Long
    Dim lngNo    As Long
    Dim lngSgn   As Long
    
    If txtSize.Text = "" Then
        Exit Sub
    End If
    
    If spbChgSize.Value = CLng(txtSize.Text) Then
        Exit Sub
    End If
    
    If frmSizeList.Caption = "サイズ一覧" Then
        'アンドゥ用に元のサイズを保存する
        If frmSizeList.Tag = E_Col Then
            Call SaveUndoInfo(E_ColSize, Selection, GetSameWidthAddresses(Selection))
        Else
            Call SaveUndoInfo(E_RowSize, Selection, GetSameHeightAddresses(Selection))
        End If
    End If
    
    If frmSizeList.CheckRatio(lblNo) = False Then
        If objColRows.Count > 1 Then
            Call SaveRatio
        End If
    End If
    
    '増加なのか減少なのかを判定
    lngSgn = Sgn(spbChgSize.Value - CLng(txtSize.Text))
    
    'サイズ調整用の番号を取得
    lngNo = frmSizeList.SelectNo
    If lngNo = 0 Then
        '選択列(行)のサイズを変更
        If frmSizeList.CheckRatio(lblNo) = True Then
            lngIndex = GetIndex(lngSgn)
        Else
            lngIndex = 1
        End If
        objColRows.size(lngIndex) = objColRows.PixelToSize(objColRows.Pixel(lngIndex) + lngSgn)     '@
    Else
        '調整列(行)のサイズを変更
        If lngNo <> CLng(lblNo) Then
            With frmParent.Controls("spbChgSize" & lngNo)
                If lngSgn > 0 Then
                    If .Value = 0 Then
                        spbChgSize.Value = CLng(txtSize.Text)
                        Exit Sub
                    End If
                End If
                .Value = .Value - lngSgn
            End With
        End If
        
        '選択列(行)のサイズを変更
        objColRows.size = objColRows.PixelToSize(objColRows.Pixel + lngSgn)     '@
    End If
    
    Call frmSizeList.ReCalc
    
    With frmSizeList
        If .Caption = "サイズ一覧" Then
            .Caption = "サイズ一覧（変更あり）"
            
            If ActiveSheet.AutoFilter Is Nothing Then
                .Controls("cmdSave").Enabled = True
                .Controls("cmdUndo").Enabled = True
            End If
            
            '閉じるボタンを無効にする
            Call EnableMenuItem(GetSystemMenu(FindWindow("ThunderDFrame", .Caption), False), SC_CLOSE, (MF_BYCOMMAND Or MF_GRAYED))
        End If
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　SaveRatio
'[ 概  要 ]　エリア内の、各列(行)の比率を保存する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SaveRatio()
    Dim i As Long
    
    Call frmSizeList.ClearRatio
    Call frmSizeList.SaveRatio(lblNo)
    ReDim dblRatio(1 To objColRows.Count)
    For i = 1 To objColRows.Count
        If Me.Pixel = 0 Then
            dblRatio(i) = 1 / objColRows.Count
        Else
            dblRatio(i) = objColRows.Pixel(i) / Me.Pixel
        End If
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　ClearRatio
'[ 概  要 ]　比率を保存した情報をクリアする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ClearRatio()
    Erase dblRatio()
End Sub

'*****************************************************************************
'[プロパティ]　Pixel
'[ 概  要 ]　サイズをピクセルで取得する
'[ 戻り値 ]　単位ピクセル
'*****************************************************************************
Public Property Get Pixel() As Long
    Pixel = objColRows.Pixel
End Property

'*****************************************************************************
'[ 関数名 ]　SetSize
'[ 概  要 ]　テキストボックスとスピンボタンに現在のサイズを再設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetSize()
    If objColRows.Pixel >= 32767 Then
        Call Err.Raise(513, , "値が32767を超えるため計算できません") '@
    End If
    
    txtSize.Text = CStr(objColRows.Pixel)
    lblmm.Caption = Format$(objColRows.Millimeters, "##0.0") & "mm"
    spbChgSize.Value = objColRows.Pixel
    
    If objColRows.Count = 1 Then '@
        lblOther.Caption = Format$(objColRows.size, "##0.00")  '@
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　SetPercent
'[ 概  要 ]　「％」を再設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetPercent(ByVal lngSumPixel As Long)
    If lngSumPixel = 0 Then
        lblPer.Caption = "％"
    ElseIf objColRows.Pixel = lngSumPixel Then
        lblPer.Caption = "100％"
    Else
        lblPer.Caption = Format$(objColRows.Pixel / lngSumPixel * 100, "#0.0") & "％"
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　SetValues
'[ 概  要 ]　状態を設定する
'[ 引  数 ]　対象のエリア
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetValues(ByRef objRange As Range)
    Dim lngCount As Long
    Dim i        As Long
    Select Case frmSizeList.Tag
    Case E_Col
        Set objColRows = New CSizeCol
        ReDim blnHidden(1 To objRange.Columns.Count)
        lblArea.Caption = GetColAddress(objRange) '例 「A:C」
        lngCount = objRange.Columns.Count
    Case E_ROW
        Set objColRows = New CSizeRow
        ReDim blnHidden(1 To objRange.Rows.Count)
        lblArea.Caption = GetRowAddress(objRange) '例 「1:10」
        lngCount = objRange.Rows.Count
    End Select
    
    Call objColRows.Initialize(objRange)
    
    '設定値の表示
    If lngCount > 1 Then
        lblOther.Caption = "詳細情報"
        lblOther.Font.Underline = True
        lblOther.Font.size = 10
        lblOther.Width = 45
        lblOther.Top = lblOther.Top + 2
        lblOther.ForeColor = &HFF0000
    End If
        
    If objColRows.Hidden = False Then
        For i = 1 To UBound(blnHidden)
            blnHidden(i) = objColRows.Hidden(i)
        Next i
    End If
    
    'テキストボックスとスピンボタンに現在のサイズを設定
    Call SetSize
End Sub

'*****************************************************************************
'[ 関数名 ]　GetColAddress
'[ 概  要 ]　エリアの名前を取得する「例　A:C」
'[ 引  数 ]　対象のエリア
'[ 戻り値 ]　なし
'*****************************************************************************
Private Function GetColAddress(ByRef objRange As Range) As String
    Dim strAddress As String
    Dim strStart   As String '先頭の列名
    Dim strLast    As String '最後の列名
    Dim lngCount   As Long   '列数
    
    lngCount = objRange.Columns.Count    '@
    Select Case Application.ReferenceStyle
    Case xlA1
        strAddress = objRange.Columns(1).Address(False, False)        '@
        strStart = Left$(strAddress, InStr(1, strAddress, ":") - 1)
        strAddress = objRange.Columns(lngCount).Address(False, False) '@
        strLast = Left$(strAddress, InStr(1, strAddress, ":") - 1)
    Case xlR1C1
        strStart = objRange.Columns(1).Column        '@
        strLast = objRange.Columns(lngCount).Column  '@
    End Select
    
    If lngCount > 1 Then
        GetColAddress = strStart & ":" & strLast
    Else
        GetColAddress = strStart
    End If
End Function

'*****************************************************************************
'[ 関数名 ]　GetRowAddress
'[ 概  要 ]　エリアの名前を取得する「例　1:10」
'[ 引  数 ]　対象のエリア
'[ 戻り値 ]　なし
'*****************************************************************************
Private Function GetRowAddress(ByRef objRange As Range) As String
    Dim strAddress As String
    Dim strStart   As String '先頭の行名
    Dim strLast    As String '最後の行名
    Dim lngCount   As Long   '行数
    
    lngCount = objRange.Rows.Count    '@
    strStart = objRange.Rows(1).Row        '@
    strLast = objRange.Rows(lngCount).Row  '@
    
    If lngCount > 1 Then
        If Len(strStart) + Len(strLast) <= 7 Then
            GetRowAddress = strStart & ":" & strLast
        Else
            GetRowAddress = strStart & ":" & vbLf & strLast
        End If
    Else
        GetRowAddress = strStart
    End If
End Function

'*****************************************************************************
'[ 関数名 ]　GetIndex
'[ 概  要 ]　今回変更対象の列の番号を取得する
'[ 引  数 ]　lngSign  縮小時:-1  拡大時:+1
'[ 戻り値 ]　列の番号
'*****************************************************************************
Private Function GetIndex(ByVal lngSign As Long) As Long
    Dim i As Long
    Dim lngNewPixel As Long
    Dim dblDiff     As Double
    Dim dblMinDiff  As Double
    
    lngNewPixel = Me.Pixel + lngSign
    
    If lngNewPixel > 0 Then
        dblMinDiff = 99999999
        For i = 1 To objColRows.Count
            If objColRows.Pixel(i) + lngSign >= 0 Then
                '元の割合との差を計算する
                dblDiff = Abs(dblRatio(i) - (objColRows.Pixel(i) + lngSign) / lngNewPixel)
                '元の割合との差が最小となる列を選択する
                If dblDiff <= dblMinDiff Then
                    If dblDiff < dblMinDiff Then
                        dblMinDiff = dblDiff
                        GetIndex = i
                    ElseIf lngSign = -1 Then
                        dblMinDiff = dblDiff
                        GetIndex = i
                    End If
                End If
            End If
        Next i
    Else
        For i = 1 To objColRows.Count
            'サイズが1ピクセル以上の列を選択する
            If objColRows.Pixel(i) > 0 Then
                GetIndex = i
            End If
        Next i
    End If
End Function

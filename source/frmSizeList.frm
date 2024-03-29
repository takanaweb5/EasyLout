VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSizeList 
   Caption         =   "サイズ一覧"
   ClientHeight    =   4560
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7392
   OleObjectBlob   =   "frmSizeList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSizeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'タイプ
Enum EColRowType
    E_Col = 1 '列変更時
    E_ROW = 2 '行変更時
End Enum

'サイズ一覧のカスタマイズ情報「@@@mm は @@@ピクセル」
Private Type TSizeInfo
    Millimeters As Long
    Pixel       As Long
End Type

Private udtSizeInfo As TSizeInfo
Private colSizeList(1 To 3) As New Collection
Private lngSelNo   As Long '調整用の選択されたエリアNo
Private lngRatioNo As Long '割合を保存しているエリアNo
Private colAddress As Collection '同じ幅の塊ごとのアドレスの配列(Undo用)
Private Const C_MAX = 256

'*****************************************************************************
'[イベント]　UserForm_Initialize
'[ 概  要 ]　フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()
End Sub

'*****************************************************************************
'[イベント]　UserForm_Terminate
'[ 概  要 ]　フォーム解放時
'*****************************************************************************
Private Sub UserForm_Terminate()
End Sub

'*****************************************************************************
'[ 関数名 ]　Initialize
'[ 概  要 ]  フォームの初期設定を行う
'[ 引  数 ]　列・行　いずれを対象とするか
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub Initialize(ByVal Value As EColRowType)
    Me.Tag = Value
    
    Application.ScreenUpdating = False
    
    '例「100mm は 378ピクセル」
    Call SetSizeInfoLabel
    
    Select Case Value
    Case E_Col
        Call InitCol
        tbsTab(1).Caption = "1列単位の情報(全体)"
    Case E_ROW
        Call InitRow
        tbsTab(1).Caption = "1行単位の情報(全体)"
    End Select
    
    Call ReCalc
    
    Application.ScreenUpdating = True
End Sub

'*****************************************************************************
'[ 関数名 ]　SetSizeInfoLabel
'[ 概  要 ]  「@@@mm は @@@ピクセル」の設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub SetSizeInfoLabel()
    Select Case Me.Tag
    Case E_Col
        udtSizeInfo.Millimeters = GetSetting(REGKEY, "SizeInfo", "Width_mm", 0)
        udtSizeInfo.Pixel = GetSetting(REGKEY, "SizeInfo", "Width_Pixel", 0)
    Case E_ROW
        udtSizeInfo.Millimeters = GetSetting(REGKEY, "SizeInfo", "Height_mm", 0)
        udtSizeInfo.Pixel = GetSetting(REGKEY, "SizeInfo", "Height_Pixel", 0)
    End Select
    
    If udtSizeInfo.Millimeters = 0 Then
        udtSizeInfo.Millimeters = 100
        udtSizeInfo.Pixel = Application.CentimetersToPoints(10) / DPIRatio
    End If
    
    With udtSizeInfo
        '例「@@@mm は @@@ピクセル」
        lblSizeInfo.Caption = .Millimeters & "mm は " & .Pixel & "ピクセル"
    End With
End Sub

'*****************************************************************************
'[ 関数名 ]　InitCol
'[ 概  要 ]  列幅サイズ変更時のフォーム初期設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub InitCol()
    Dim objSizeList  As CSizeList
    Dim i            As Long
    
    '***************************************
    '１列単位の情報を作成
    '***************************************
    Dim objSelection As Range
    Dim j            As Long
    Dim k            As Long
    
    '選択範囲のColumnsの和集合を取り重複列を排除する
    Set objSelection = Union(Selection.EntireColumn, Selection.EntireColumn)
        
    'トータル列数のカウント
    For i = 1 To objSelection.Areas.Count
        j = j + objSelection.Areas(i).Columns.Count
        
        '256を最大とする
        If j > C_MAX Then
            Application.ScreenUpdating = True
            Call Err.Raise(513, , C_MAX & "以上の列に対して実行できません。")
        End If
    Next
    
    '配列に列番号を格納する
    ReDim udtSortArray(1 To j) As TSortArray
    For i = 1 To objSelection.Areas.Count
        For j = 1 To objSelection.Areas(i).Columns.Count
            k = k + 1
            udtSortArray(k).Key1 = objSelection.Areas(i).Columns(j).Column
        Next
    Next

    '列番号でソートする
    Call SortArray(udtSortArray())
    
    '列の数だけループ
    For i = 1 To UBound(udtSortArray)
        k = udtSortArray(i).Key1
        Set objSizeList = New CSizeList
        Call objSizeList.CreateSizeList(i, frmPage2)
        Call objSizeList.SetValues(Columns(k).EntireColumn)
        Call colSizeList(2).Add(objSizeList)
    Next
    
    '***************************************
    'エリア単位の情報を作成
    '***************************************
    'エリアの数だけループ
    For i = 1 To Selection.Areas.Count
        Set objSizeList = New CSizeList
        Call objSizeList.CreateSizeList(i, frmPage1)
        Call objSizeList.SetValues(Selection.Areas(i).Columns.EntireColumn)
        Call colSizeList(1).Add(objSizeList)
    Next
End Sub

'*****************************************************************************
'[ 関数名 ]　InitRow
'[ 概  要 ]  行高サイズ変更時のフォーム初期設定
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub InitRow()
    Dim objSizeList  As CSizeList
    Dim i            As Long
    
    '***************************************
    '１列単位の情報を作成
    '***************************************
    Dim objSelection As Range
    Dim j            As Long
    Dim k            As Long
    
    '選択範囲のColumnsの和集合を取り重複列を排除する
    Set objSelection = Union(Selection.EntireRow, Selection.EntireRow)
        
    'トータル列数のカウント
    For i = 1 To objSelection.Areas.Count
        j = j + objSelection.Areas(i).Rows.Count
        
        '256を最大とする
        If j > C_MAX Then
            Application.ScreenUpdating = True
            Call Err.Raise(513, , C_MAX & "以上の行に対して実行できません。")
        End If
    Next
    
    '配列に行番号を格納する
    ReDim udtSortArray(1 To j) As TSortArray
    For i = 1 To objSelection.Areas.Count
        For j = 1 To objSelection.Areas(i).Rows.Count
            k = k + 1
            udtSortArray(k).Key1 = objSelection.Areas(i).Rows(j).Row
        Next
    Next

    '列番号でソートする
    Call SortArray(udtSortArray())
    
    '列の数だけループ
    For i = 1 To UBound(udtSortArray)
        k = udtSortArray(i).Key1
        Set objSizeList = New CSizeList
        Call objSizeList.CreateSizeList(i, frmPage2)
        Call objSizeList.SetValues(Rows(k).EntireRow)
        Call colSizeList(2).Add(objSizeList)
    Next

    '***************************************
    'エリア単位の情報を作成
    '***************************************
    'エリアの数だけループ
    For i = 1 To Selection.Areas.Count
        Set objSizeList = New CSizeList
        Call objSizeList.CreateSizeList(i, frmPage1)
        Call objSizeList.SetValues(Selection.Areas(i).Rows.EntireRow)
        Call colSizeList(1).Add(objSizeList)
    Next
End Sub

'*****************************************************************************
'[イベント]　frmPage1_Exit
'[ 概  要 ]　保存した割合をクリアする
'*****************************************************************************
Private Sub frmPage1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call ClearRatio
End Sub

'*****************************************************************************
'[イベント]　lblSizeInfo_Click
'[ 概  要 ]　「@@@mm は @@@ピクセル」の値をカスタマイズさせる
'*****************************************************************************
Public Sub lblSizeInfo_Click()
    With frmSizeInfo
        'フォームを表示
        Call .Show
        
        'キャンセル時
        If blnFormLoad = False Then
            Exit Sub
        End If
        
        Call Unload(frmSizeInfo)
    End With
    
    Call SetSizeInfoLabel
    Call ReCalc
End Sub

'*****************************************************************************
'[イベント]　UserForm_QueryClose
'[ 概  要 ]　×ボタンでフォームを閉じる時、変更を元に戻す
'*****************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    '変更がなければフォームを閉じる
    If Me.Caption <> "サイズ一覧（変更あり）" Then
        Exit Sub
    End If
        
    '×ボタンでフォームを閉じる時、フォームを閉じない
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Exit Sub
    End If
    
    Call SetOnUndo
End Sub

'*****************************************************************************
'[イベント]　cmdSave_Click
'[ 概  要 ]　確定ボタン押下時
'*****************************************************************************
Private Sub cmdSave_Click()
    Me.Caption = "サイズ一覧"
    cmdSave.Enabled = False
    cmdUndo.Enabled = False
    '閉じるボタンを有効にする
    Call EnableMenuItem(GetSystemMenu(FindWindow("ThunderDFrame", Me.Caption), False), SC_CLOSE, MF_BYCOMMAND)
End Sub

'*****************************************************************************
'[イベント]　cmdUndo_Click
'[ 概  要 ]　元に戻すボタン押下時
'*****************************************************************************
Private Sub cmdUndo_Click()
    Call ExecUndo
    Application.ScreenUpdating = True
    Call ReCalc
    
    Call cmdSave_Click
End Sub

'*****************************************************************************
'[イベント]　cmdOK_Click
'[ 概  要 ]　ＯＫボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[イベント]　tbsTab_Change
'[ 概  要 ]　タブ変更時
'*****************************************************************************
Private Sub tbsTab_Change()
    Dim lngNo   As Long '「詳細情報」をクリックしたNo
    Dim frmPage As MSForms.Frame
    
    Select Case Me.Tag
    Case E_Col
        tbsTab(1).Caption = "1列単位の情報"
    Case E_ROW
        tbsTab(1).Caption = "1行単位の情報"
    End Select
    
    Select Case tbsTab.Value
    Case 0 '左側のタブ
        frmPage1.Visible = True
        frmPage2.Visible = False
        frmPage3.Visible = False
        Me.Controls("tbsTab").Tag = 0
    Case 1 '右側のタブ
        lngNo = Me.Controls("tbsTab").Tag
        If lngNo = 0 Then
            frmPage1.Visible = False
            frmPage2.Visible = True
            frmPage3.Visible = False
        Else '「詳細情報」をクリックしてタブが変更された時
            frmPage1.Visible = False
            frmPage2.Visible = False
            frmPage3.Visible = True
        End If
    End Select
        
    '「選択」チェックのクリア
    lngSelNo = 0
    If tbsTab.Value = 1 Then
        If lngNo = 0 Then
            Call ClearSelChk(frmPage2)
        Else
            Call ClearSelChk(frmPage3)
        End If
    End If
    
    '「詳細情報」をクリックしてタブが変更された時
    If lngNo <> 0 Then
        tbsTab(1).Caption = tbsTab(1).Caption & "(詳細)"
        Select Case Me.Tag
        Case E_Col
            Call TabChangeCol(lngNo)
        Case E_ROW
            Call TabChangeRow(lngNo)
        End Select
    Else
        tbsTab(1).Caption = tbsTab(1).Caption & "(全体)"
    End If
    
    Call ReCalc
End Sub

'*****************************************************************************
'[ 関数名 ]　ClearSelChk
'[ 概  要 ]  「選択」チェックのクリア
'[ 引  数 ]　クリアするフレームオブジェクト
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ClearSelChk(ByRef frmPage As MSForms.Frame)
On Error GoTo ErrHandle
    Dim i As Long
    For i = 1 To C_MAX
        frmPage.Controls("chkSelect" & i).Value = False
    Next i
ErrHandle:
End Sub
'*****************************************************************************
'[ 関数名 ]　TabChangeCol
'[ 概  要 ]  「詳細情報」をクリックしてタブが変更された時(列変更時)
'[ 引  数 ]　エリアのNo
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub TabChangeCol(ByVal lngNo As Long)
    Dim objSizeList As CSizeList
    Dim objRange    As Range
    Dim i           As Long

    Set colSizeList(3) = New Collection
    Set objRange = Selection.Areas(lngNo)
    
    '列の数だけループ
    For i = 1 To objRange.Columns.Count  '@
        Set objSizeList = New CSizeList
        Call objSizeList.CreateSizeList(i, frmPage3)
        Call objSizeList.SetValues(objRange.Columns(i).EntireColumn) '@
        Call colSizeList(3).Add(objSizeList)

        '100を最大とする
        If i = C_MAX Then
            Exit For
        End If
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　TabChangeRow
'[ 概  要 ]  「詳細情報」をクリックしてタブが変更された時(行変更時)
'[ 引  数 ]　エリアのNo
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub TabChangeRow(ByVal lngNo As Long)
    Dim objSizeList As CSizeList
    Dim objRange    As Range
    Dim i           As Long

    Set colSizeList(3) = New Collection
    Set objRange = Selection.Areas(lngNo)
    
    '列の数だけループ
    For i = 1 To objRange.Rows.Count  '@
        Set objSizeList = New CSizeList
        Call objSizeList.CreateSizeList(i, frmPage3)
        Call objSizeList.SetValues(objRange.Rows(i).EntireRow) '@
        Call colSizeList(3).Add(objSizeList)

        '100を最大とする
        If i = C_MAX Then
            Exit For
        End If
    Next i
End Sub

'*****************************************************************************
'[ 関数名 ]　ReCalc
'[ 概  要 ]  ｢合計｣,｢平均｣に最新の値を設定する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ReCalc()
    Dim objSizeList As CSizeList
    Dim lngSumPixel As Long
    Dim i           As Long
    Dim strAvg      As String
    
    If tbsTab.Value = 0 Then
        i = 1
    ElseIf Me.Controls("tbsTab").Tag = 0 Then
        i = 2
    Else
        i = 3
    End If
    
    'エリアの数だけループ
    For Each objSizeList In colSizeList(i)
        Call objSizeList.SetSize
        lngSumPixel = lngSumPixel + objSizeList.Pixel
    Next objSizeList

    'エリアの数だけループ
    For Each objSizeList In colSizeList(i)
        Call objSizeList.SetPercent(lngSumPixel)
    Next objSizeList

    txtSum.Text = CStr(lngSumPixel)
    lblSum.Caption = Format$(PixelToMillimeter(lngSumPixel), "##0.0") & "mm"
    
    If colSizeList(i).Count <> 0 Then
        txtAvg.Text = Format$(CStr(lngSumPixel / colSizeList(i).Count), "0.0")
        strAvg = Format$(PixelToMillimeter(lngSumPixel / colSizeList(i).Count), "##0.0") & "mm"
        If colSizeList(i).Count = 1 Then
            lblAvg.Caption = strAvg
        Else
            lblAvg.Caption = strAvg & Format$(1 / colSizeList(i).Count * 100, " #0.0") & "％"
        End If
    End If

    Call Me.Controls("frmPage" & CStr(i)).SetFocus
End Sub

'*****************************************************************************
'[ 関数名 ]　PixelToMillimeter
'[ 概  要 ]  単位の変換 Pixel → mm
'[ 引  数 ]　Pixel
'[ 戻り値 ]　mm
'*****************************************************************************
Public Function PixelToMillimeter(ByVal lngPixel As Long) As Double
    With udtSizeInfo
        PixelToMillimeter = lngPixel * .Millimeters / .Pixel
    End With
End Function

'*****************************************************************************
'[プロパティ]　SelectNo
'[ 概  要 ]　選択でチェックされている、エリアNo
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get SelectNo() As Long
    SelectNo = lngSelNo
End Property
Public Property Let SelectNo(ByVal Value As Long)
    lngSelNo = Value
End Property

'*****************************************************************************
'[ 関数名 ]　SaveRatio
'[ 概  要 ]　比率を保存するエリアNoを保存
'[ 引  数 ]　エリアNo
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SaveRatio(ByVal lngNo As Long)
    If tbsTab.Value = 0 Then
        lngRatioNo = lngNo
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　CheckRatio
'[ 概  要 ]　比率が保存されているかチェック
'[ 引  数 ]　エリアNo
'[ 戻り値 ]　True:保存されている
'*****************************************************************************
Public Function CheckRatio(ByVal lngNo As Long) As Boolean
    If tbsTab.Value = 0 And lngRatioNo = lngNo Then
        CheckRatio = True
    Else
        CheckRatio = False
    End If
End Function

'*****************************************************************************
'[ 関数名 ]　ClearRatio
'[ 概  要 ]　比率を保存した情報をクリアする
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ClearRatio()
    If lngRatioNo <> 0 Then
        Call colSizeList(1).Item(lngRatioNo).ClearRatio
        lngRatioNo = 0
    End If
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUnSelect 
   Caption         =   "選択してください"
   ClientHeight    =   2328
   ClientLeft      =   108
   ClientTop       =   336
   ClientWidth     =   4668
   OleObjectBlob   =   "frmUnSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmUnSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'領域の取消し画面のモード
Public Enum EUnselectMode
    E_Unselect  '取消し
    E_Reverse   '反転
    E_Union     '追加
    E_Intersect '絞り込み
    E_Merge    '結合セルのみ選択
End Enum

Private lngReferenceStyle As Long
Private strSelectBefore As String
Private blnCheck As Boolean

Private strLastSheet   As String '前回の領域の復元用
Private strLastAddress As String '前回の領域の復元用

'*****************************************************************************
'[イベント]　各種マウス操作時
'[ 概  要 ]　RefEditを有効にさせる
'*****************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub Frame_Click()
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub lblTitle_Click()
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub lblTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub cmdOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub cmdCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    RefEdit.SetFocus
End Sub

'*****************************************************************************
'[イベント]　RefEditで領域選択時
'[ 概  要 ]　アドレスが255文字を超えた時の対応
'*****************************************************************************
Private Sub RefEdit_Change()
    '[Ctrl]Keyが押下されていれば、選択領域を次々と追加している時
    If GetKeyState(vbKeyControl) < 0 Then
        'アドレスが255文字を超えてリセットされてしまった時
        If strSelectBefore <> "" Then
            If Range(RefEdit.Value).Areas.Count <= 1 And _
               Range(strSelectBefore).Areas.Count > 1 Then
                Call MsgBox("これ以上は選択出来ません", vbExclamation)
                RefEdit.Value = strSelectBefore
                Call cmdOK.SetFocus
                Call RefEdit.SetFocus
                Exit Sub
            End If
        End If
    End If
    strSelectBefore = RefEdit.Value
End Sub

''*****************************************************************************
''[イベント]　RefEditで領域選択中にCtrlキーを離した時
''[ 概  要 ]　アドレスが255文字を超えた時の対応
''*****************************************************************************
'Private Sub SetEnterEnabled()
'    If RefEdit.Value = "" And Me.Mode <> E_Merge Then
'        cmdOK.Enabled = False
'    Else
'        cmdOK.Enabled = True
'    End If
'End Sub

'*****************************************************************************
'[イベント]　RefEditで領域選択中にCtrlキーを離した時
'[ 概  要 ]　アドレスが255文字を超えた時の対応
'*****************************************************************************
Private Sub RefEdit_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyControl Then
        RefEdit.Value = strSelectBefore
        Call cmdOK.SetFocus
        Call RefEdit.SetFocus
    End If
End Sub

'*****************************************************************************
'[イベント]　chkReverse_Change
'[ 概  要 ]　反転チェック時
'*****************************************************************************
Private Sub chkReverse_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkReverse.Value = True Then
        Call ChangeMode(E_Reverse)
    Else
        Call ChangeMode(E_Unselect)
    End If
End Sub

'*****************************************************************************
'[イベント]　chkIntersect_Change
'[ 概  要 ]　絞り込みチェック時
'*****************************************************************************
Private Sub chkIntersect_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkIntersect.Value = True Then
        Call ChangeMode(E_Intersect)
    Else
        Call ChangeMode(E_Unselect)
    End If
End Sub

'*****************************************************************************
'[イベント]　chkUnion_Change
'[ 概  要 ]　追加チェック時
'*****************************************************************************
Private Sub chkUnion_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkUnion.Value = True Then
        Call ChangeMode(E_Union)
    Else
        Call ChangeMode(E_Unselect)
    End If
End Sub

'*****************************************************************************
'[イベント]　chkMerge_Click
'[ 概  要 ]　結合セルのみチェック時
'*****************************************************************************
Private Sub chkMerge_Click()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkMerge.Value = True Then
        Call ChangeMode(E_Merge)
    Else
        Call ChangeMode(E_Unselect)
    End If
End Sub

'*****************************************************************************
'[ 関数名 ]　ChangeMode
'[ 概  要 ]  ｢反転｣･｢絞り込み｣･｢取消し｣のモードを変更する
'[ 引  数 ]　モードタイプ
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub ChangeMode(ByVal enmModeType As EUnselectMode)
    Select Case enmModeType
    Case E_Unselect
        lblTitle.Caption = "マウスで選択を取消させたい領域を選択してください"
    Case E_Reverse
        lblTitle.Caption = "マウスで選択を反転させたい領域を選択してください"
    Case E_Intersect
        lblTitle.Caption = "マウスで選択を絞り込みたい領域を選択してください"
    Case E_Union
        lblTitle.Caption = "マウスで選択を追加したい領域を選択してください"
    Case E_Merge
        lblTitle.Caption = "結合セルのみを選択します"
    End Select
    
    blnCheck = True
    chkReverse.Value = (enmModeType = E_Reverse)
    chkIntersect.Value = (enmModeType = E_Intersect)
    chkUnion.Value = (enmModeType = E_Union)
    chkMerge.Value = (enmModeType = E_Merge)
 
    blnCheck = False
    
    If enmModeType = E_Merge Then
        RefEdit.Enabled = False
    Else
        RefEdit.Enabled = True
        Call RefEdit.SetFocus
    End If
End Sub
    
'*****************************************************************************
'[イベント]　UserForm_Initialize
'[ 概  要 ]　フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()
    lngReferenceStyle = Application.ReferenceStyle
    Application.ReferenceStyle = xlA1

    'RefEditを隠す
    RefEdit.Top = RefEdit.Top + 100
    
    '呼び元に通知する
    blnFormLoad = True
End Sub

'*****************************************************************************
'[イベント]　UserForm_Terminate
'[ 概  要 ]　フォームアンロード時
'*****************************************************************************
Private Sub UserForm_Terminate()
    Application.ReferenceStyle = lngReferenceStyle
    '呼び元に通知する
    blnFormLoad = False
End Sub

'*****************************************************************************
'[イベント]　cmdLastSelect_Click
'[ 概  要 ]　前回の領域の復元ボタン押下時
'*****************************************************************************
Private Sub cmdLastSelect_Click()
    Call Range(strLastAddress).Select
End Sub

'*****************************************************************************
'[イベント]　cmdOK_Click
'[ 概  要 ]　ＯＫボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
    Call cmdOK.SetFocus
    Me.Hide
End Sub

'*****************************************************************************
'[イベント]　cmdCancel_Click
'[ 概  要 ]　キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[ 関数名 ]　SetLastSelect
'[ 概  要 ]　直前のコマンド実行時に選択された領域のアドレスを保存する
'[ 引  数 ]  直前の領域の情報
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub SetLastSelect(ByVal strSheetName As String, ByVal strAddress As String)
    strLastSheet = strSheetName
    strLastAddress = strAddress
    
    If strLastAddress = "" Or ActiveSheet.Name <> strLastSheet Then
        cmdLastSelect.Enabled = False
    End If
End Sub

'*****************************************************************************
'[プロパティ]　SelectRange
'[ 概  要 ]　選択された領域
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get SelectRange() As Range
    Dim objRange  As Range
    Dim strAddress As String
    
    For Each objRange In Range(RefEdit.Value).Areas
        strAddress = strAddress & Range(GetMergeAddress(objRange.Address)).Address(False, False) & ","
    Next
    
    '末尾のカンマを削除
    Set SelectRange = Range(Left$(strAddress, Len(strAddress) - 1))
End Property

'*****************************************************************************
'[プロパティ]　Mode
'[ 概  要 ]　選択モード
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get Mode() As EUnselectMode
    Select Case (True)
    Case chkReverse.Value
        Mode = E_Reverse
    Case chkIntersect.Value
        Mode = E_Intersect
    Case chkUnion.Value
        Mode = E_Union
    Case chkMerge.Value
        Mode = E_Merge
    Case Else
        Mode = E_Unselect
    End Select
End Property

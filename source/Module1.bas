Attribute VB_Name = "Module1"
Option Explicit

Sub ClearSheets()
Attribute ClearSheets.VB_Description = "マクロ記録日 : 2012/7/4  ユーザー名 : naoki-takahashi"
Attribute ClearSheets.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim objSheet As Worksheet
    Dim objActiveSheet As Worksheet
    
    Set objActiveSheet = ActiveSheet
    
'    Application.ScreenUpdating = False
    For Each objSheet In Worksheets
        Call ClearSheet(objSheet)
    Next
    
'    Application.ScreenUpdating = True
'    For Each objSheet In Worksheets
'        'A1の位置を表示させる
'        Call objSheet.Activate
'        Call objSheet.Range("A1").Select
'    Next
    
    'カラーパレットの初期化
    Call ActiveWorkbook.ResetColors
    
    'スタイルの削除
    Call DeleteStyles(ActiveWorkbook)
    
    '名前オブジェクトを削除する
    Call DeleteNames(ActiveWorkbook)
    
    Call objActiveSheet.Select
End Sub

'*****************************************************************************
'[ 関数名 ]　DeleteNames
'[ 概  要 ]　名前オブジェクトを削除する
'[ 引  数 ]　Workbook
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub DeleteNames(ByRef objWorkbook As Workbook)
    Dim objName     As Name
    For Each objName In objWorkbook.Names
        If (Right$(objName.Name, Len("Print_Area")) <> "Print_Area") And _
           (Right$(objName.Name, Len("Print_Titles")) <> "Print_Titles") And _
           (Right$(objName.Name, Len("Database")) <> "Database") Then
'            Debug.Print objName.Name
            Call objName.Delete
        End If
    Next objName
End Sub

Private Sub ClearSheet(ByRef objSheet As Worksheet)
    Call objSheet.Activate
    
'    Call ActiveWindow.ScrollIntoView(0, 0, 1, 1)
    Call ActiveWindow.SmallScroll(, Rows.Count, , Columns.Count)
    Call objSheet.Range("A1").Select
    
    '枠線の表示
'    ActiveWindow.DisplayGridlines = False
    
    '分割を解除
    If ActiveWindow.Split = True And ActiveWindow.FreezePanes = False Then
        ActiveWindow.Split = False
    End If
    '改ページプレビュー解除
    ActiveWindow.View = xlNormalView
    
    '改ページ表示を解除
    objSheet.DisplayAutomaticPageBreaks = False
    
    '倍率
    ActiveWindow.Zoom = 100
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSizeCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objColumn    As Range   '該当の列

'*****************************************************************************
'[ 関数名 ]　Initialize
'[ 概  要 ]　対象の列(Range)を設定する
'[ 引  数 ]　対象の列 Rangeオブジェクト
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub Initialize(ByRef objRange As Range)
    Set objColumn = objRange
End Sub

'*****************************************************************************
'[プロパティ]　Size
'[ 概  要 ]　列の幅
'[ 引  数 ]　列の番号
'*****************************************************************************
Public Property Get Size(Optional ByVal i As Long = 1) As Double
    Size = objColumn.Columns(i).ColumnWidth
End Property
Public Property Let Size(Optional ByVal i As Long = 1, ByVal Value As Double)
    objColumn.Columns(i).ColumnWidth = Value
End Property

'*****************************************************************************
'[プロパティ]　Count
'[ 概  要 ]　列の数
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get Count() As Long
    Count = objColumn.Columns.Count
End Property

'*****************************************************************************
'[プロパティ]　Hidden
'[ 概  要 ]　非表示かどうか
'[ 引  数 ]　列の番号(0:全体が非表示かどうか)
'*****************************************************************************
Public Property Get Hidden(Optional ByVal i As Long = 1) As Long
    If i = 0 Then
        Hidden = objColumn.Columns.Hidden
    Else
        Hidden = objColumn.Columns(i).Hidden
    End If
End Property

'*****************************************************************************
'[プロパティ]　Pixel
'[ 概  要 ]　列の幅をピクセルで取得する
'[ 戻り値 ]　列の幅(単位ピクセル)
'*****************************************************************************
Public Property Get Pixel(Optional ByVal i As Long = 0) As Long
    If i = 0 Then
        Pixel = objColumn.Width / DPIRatio
    Else
        Pixel = objColumn.Columns(i).Width / DPIRatio
    End If
End Property

'*****************************************************************************
'[プロパティ]　Millimeters
'[ 概  要 ]　列の幅をmmで取得する
'[ 戻り値 ]　列の幅(単位mm)
'*****************************************************************************
Public Property Get Millimeters(Optional ByVal i As Long = 0) As Double
    Millimeters = frmSizeList.PixelToMillimeter(Me.Pixel(i))
End Property

'*****************************************************************************
'[ 関数名 ]　PixelToSize
'[ 概  要 ]　幅の単位を変換
'[ 引  数 ]　lngPixel : 幅(単位:ピクセル)
'[ 戻り値 ]　ColumnWidth
'*****************************************************************************
Public Function PixelToSize(ByVal lngPixel As Long) As Double
    PixelToSize = PixelToWidth(lngPixel)
End Function

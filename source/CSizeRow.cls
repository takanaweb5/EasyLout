VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSizeRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private objRow    As Range   '該当の行

'*****************************************************************************
'[ 関数名 ]　Initialize
'[ 概  要 ]　対象の行(Range)を設定する
'[ 引  数 ]　対象の行 Rangeオブジェクト
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub Initialize(ByRef objRange As Range)
    Set objRow = objRange
End Sub

'*****************************************************************************
'[プロパティ]　Size
'[ 概  要 ]　行の高さ
'[ 引  数 ]　行の番号
'*****************************************************************************
Public Property Get Size(Optional ByVal i As Long = 1) As Double
    Size = objRow.Rows(i).RowHeight
End Property
Public Property Let Size(Optional ByVal i As Long = 1, ByVal Value As Double)
    objRow.Rows(i).RowHeight = Value
End Property

'*****************************************************************************
'[プロパティ]　Count
'[ 概  要 ]　行の数
'[ 引  数 ]　なし
'*****************************************************************************
Public Property Get Count() As Long
    Count = objRow.Rows.Count
End Property

'*****************************************************************************
'[プロパティ]　Hidden
'[ 概  要 ]　非表示かどうか
'[ 引  数 ]　行の番号(0:全体が非表示かどうか)
'*****************************************************************************
Public Property Get Hidden(Optional ByVal i As Long = 1) As Long
    If i = 0 Then
        Hidden = objRow.Rows.Hidden
    Else
        Hidden = objRow.Rows(i).Hidden
    End If
End Property

'*****************************************************************************
'[プロパティ]　Pixel
'[ 概  要 ]　行の高さをピクセルで取得する
'[ 戻り値 ]　行の高さ(単位ピクセル)
'*****************************************************************************
Public Property Get Pixel(Optional ByVal i As Long = 0) As Long
    If i = 0 Then
        Pixel = objRow.Height / 0.75
    Else
        Pixel = objRow.Rows(i).Height / 0.75
    End If
End Property

'*****************************************************************************
'[プロパティ]　Millimeters
'[ 概  要 ]　行の高さをmmで取得する
'[ 戻り値 ]　行の高さ(単位mm)
'*****************************************************************************
Public Property Get Millimeters(Optional ByVal i As Long = 0) As Double
    Millimeters = frmSizeList.PixelToMillimeter(Me.Pixel(i))
End Property

'*****************************************************************************
'[ 関数名 ]　PixelToSize
'[ 概  要 ]　高さの単位を変換
'[ 引  数 ]　lngPixel : 幅(単位:ピクセル)
'[ 戻り値 ]　RowHeight
'*****************************************************************************
Public Function PixelToSize(ByVal lngPixel As Long) As Double
    PixelToSize = PixelToHeight(lngPixel)
End Function

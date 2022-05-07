VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCopyCell 
   Caption         =   "�����܂ܓ\�t��"
   ClientHeight    =   2976
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5244
   OleObjectBlob   =   "frmCopyCell.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmCopyCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TRect
    Top      As Long
    Height   As Long
    Left     As Long
    Width    As Long
End Type

'Private blnCheck As Boolean

Private FEnableEvents As Boolean
Private FMove As Boolean
Private FBlockCount As Long
Private FWidthRatio() As Double
Private FBackupDstColWidth() As Double
Private FSrcRange  As Range
Private FDstRange  As Range
Private FSrcColAddress() As String
Private FDstColAddress() As String

Private objTextbox()   As Shape  'idx 0:�O�g,1�`FBlockCount:�񌩏o��,FBlockCount:���g
Private lngDisplayObjects As Long
Private lngZoom      As Long

'*****************************************************************************
'[�T�v] �����ݒ����ݒ�
'[����] objFromRange:�R�s�[���̗̈�
'       objToRange:�\�t����̃Z��
'       blnMove:True=�ړ�����
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub Initialize(ByRef objFromRange As Range, ByRef objToRange As Range, ByVal blnMove As Boolean)
    chkIgnore.ControlTipText = "�ʏ�ͤ�f�t�H���g�̂܂܎��s��������(�R�s�[����EXCEL����Ɣ��肵�����Ƀ`�F�b�N����܂�)"
    chkBlank.ControlTipText = "�\�t����A�󔒃Z���̌������������܂�"
    FMove = blnMove
    
    lngDisplayObjects = ActiveWorkbook.DisplayDrawingObjects
    FEnableEvents = False
    chkIgnore.Value = IsGraphpaper(objFromRange)
    FEnableEvents = True
    
    Call Init(objFromRange)
    Call GetDstRange(objToRange)
    Call FDstRange.Select
    Call SetDstColAddress
    
    '�R�s�[��̃V�[�g��Activate�ɂ���
    Call objToRange.Worksheet.Activate
    
    Call MakeTextBox

'    '�I��̈悪��ʂ�������Ă��鎞
'    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '��ʕ����̂Ȃ���
'        If IntersectRange(ActiveWindow.VisibleRange, objToRange) Is Nothing Then
'            With objToRange
'                Call ActiveWindow.ScrollIntoView(.Left / DPIRatio, .Top / DPIRatio, .Width / DPIRatio, .Height / DPIRatio)
'            End With
'        End If
'    End If
End Sub

'*****************************************************************************
'[�T�v] �e�L�X�g�{�b�N�X���쐬
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub MakeTextBox()
    ActiveWorkbook.DisplayDrawingObjects = xlDisplayShapes
    ReDim objTextbox(0 To FBlockCount + 1)
    
    '���g�쐬
    If FDstRange.Rows.Count > 1 Then
        With MinusRange(FDstRange, FDstRange.Rows(1))
            Set objTextbox(FBlockCount + 1) = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
        End With
        '�G���[�̎��ɐԎ��ŃG���[���e��\��������
        With objTextbox(FBlockCount + 1).TextFrame2.TextRange.Font
            .NameComplexScript = ActiveWorkbook.Styles("Normal").Font.Name
            .NameFarEast = ActiveWorkbook.Styles("Normal").Font.Name
            .Name = ActiveWorkbook.Styles("Normal").Font.Name
            .Size = ActiveWorkbook.Styles("Normal").Font.Size
        End With
        With objTextbox(FBlockCount + 1).TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.Rgb = Rgb(255, 0, 0) '��
            .Transparency = 0
        End With
        '�e�L�X�g�{�b�N�X�̌r���Ȃ�
        With objTextbox(FBlockCount + 1).Line
            .Visible = msoFalse
        End With
        With objTextbox(FBlockCount + 1).Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.SchemeColor = 65
            .Transparency = 0.12  '�w�i�𓧂�������
        End With
    End If
    
    '������P�s�ڂɕ\��
    Dim i As Long
    For i = 1 To FBlockCount
        '�e�L�X�g�{�b�N�X�쐬
        With Intersect(Range(FDstColAddress(i)), FDstRange.Rows(1))
            Set objTextbox(i) = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
        End With
        With objTextbox(i).TextFrame2.TextRange.Font
            .NameComplexScript = ActiveWorkbook.Styles("Normal").Font.Name
            .NameFarEast = ActiveWorkbook.Styles("Normal").Font.Name
            .Name = ActiveWorkbook.Styles("Normal").Font.Name
            .Size = ActiveWorkbook.Styles("Normal").Font.Size
        End With
        With objTextbox(i).TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.Rgb = Rgb(0, 0, 0)
            .Transparency = 0
        End With
        With objTextbox(i).TextFrame2
            .VerticalAnchor = msoAnchorMiddle '������
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter '������
            '�O��̗]����0��
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
        End With
        With objTextbox(i).Fill
            .ForeColor.Rgb = Rgb(218, 231, 245)
            .Transparency = 0
        End With
        With objTextbox(i).Line
            .Visible = msoTrue
            .ForeColor.Rgb = Rgb(0, 0, 0)
            .Weight = 0.75
            .Transparency = 0
        End With
    Next
    
    '�O�g�쐬
    With FDstRange
        Set objTextbox(0) = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
    End With
    '�e�L�X�g�{�b�N�X�̔w�i�Ȃ�
    With objTextbox(0).Fill
        .Visible = msoFalse
    End With
    '�e�L�X�g�{�b�N�X�̌r����ύX
    With objTextbox(0).Line
        .Weight = 2#
        .Style = msoLineSingle
        .Transparency = 0#
        .Visible = msoTrue
        .ForeColor.SchemeColor = 64
        .BackColor.Rgb = Rgb(255, 255, 255)
        .Pattern = msoPattern50Percent
    End With
    
'    Call objTextbox(0).ZOrder(msoBringToFront)
    
    '���o����ɃR�s�[���̗�A�h���X��\��
    Call EditTextbox
End Sub

'*****************************************************************************
'[�T�v] �e�L�X�g�{�b�N�X�̉������ύX���āA���o���s�ɃR�s�[���̗�A�h���X��\��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub EditTextbox()
    Dim i As Long
    '���o���s
    For i = 1 To FBlockCount
        '����ύX
        With Intersect(Range(FDstColAddress(i)), FDstRange.Rows(1))
            objTextbox(i).Left = .Left
            objTextbox(i).Width = .Width
        End With
        With objTextbox(i).TextFrame2.TextRange
            '�e�L�X�g�{�b�N�X�ɗ񖼂�\��
            .Characters.Text = GetColAddress(FSrcColAddress(i))
        End With
    Next
    
    '���g
    If FDstRange.Rows.Count > 1 Then
        With MinusRange(FDstRange, FDstRange.Rows(1))
            objTextbox(FBlockCount + 1).Left = .Left
            objTextbox(FBlockCount + 1).Width = .Width
        End With
    End If
    
    '�O�g
    With FDstRange
        objTextbox(0).Left = .Left
        objTextbox(0).Width = .Width
    End With

    lblCellAddress.Caption = " " & FDstRange.Address(0, 0)
    Call CheckPaste
End Sub

'*****************************************************************************
'[�T�v] ��FA:A �� A, A:B �� A:B
'[����] ��FA:A
'[�ߒl] ��FA
'*****************************************************************************
Private Function GetColAddress(ByVal strAddress As String) As String
    Dim strABC
    GetColAddress = strAddress
    On Error Resume Next
    strABC = Split(strAddress, ":")
    If strABC(0) = strABC(1) Then
        GetColAddress = strABC(0)
    End If
End Function

'*****************************************************************************
'[�T�v] FSrcColAddress��FSrcWidth��ݒ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub Init(ByRef objSelection As Range)
    Dim i As Long
    Set FSrcRange = objSelection
    
    'True:�Z���̍����������Z���̋��E�ƂȂ��
    ReDim IsBoderCol(1 To FSrcRange.Columns.Count + 1) As Boolean
    
    Dim objArea As Range
    Dim RightCol As Long '�Z���̉E���������Z���̋��E�ƂȂ��ԍ�
    Dim objWkColumns As Range
    Dim Offset As Long
    Offset = objSelection.Column - 1
    
    '�I�����ꂽ�񐔕�Loop����IsBoderCol()��ݒ�
    IsBoderCol(1) = True
    IsBoderCol(FSrcRange.Columns.Count + 1) = True
    Dim x As Long, y As Long
    For x = 2 To FSrcRange.Columns.Count
        For y = 1 To FSrcRange.Rows.Count
            If chkIgnore.Value Then
                With FSrcRange.Cells(y, x).MergeArea
                    '�����Z���̎�
                    If .Count > 1 Then
                        '���[�̗�Ƀt���O�����Ă�
                        IsBoderCol(.Column - Offset) = True
                        '�E�[�̉E�ׂ̗�Ƀt���O�����Ă�
                        IsBoderCol(.Column + .Columns.Count - Offset) = True
                    ElseIf VarType(.Value) <> vbEmpty Then
                        '�l�̓��͂��ꂽ�Z���̎��A�t���O�����Ă�
                        IsBoderCol(x) = True
                    End If
                End With
            Else
                With FSrcRange.Cells(y, x)
                    If .MergeArea.Column = .Column Then
                        IsBoderCol(x) = True
                    End If
                End With
            End If
        Next
    Next
    
    Dim j As Long
    Dim LeftCol As Long  '�Z���̍����������Z���̋��E�ƂȂ��ԍ�
    LeftCol = 1
    ReDim FSrcWidth(1 To FSrcRange.Columns.Count)
    ReDim FSrcColAddress(1 To FSrcRange.Columns.Count)
    '�I�����ꂽ�񐔕�Loop
    For i = 2 To FSrcRange.Columns.Count + 1
        If IsBoderCol(i) Then
            With FSrcRange
                j = j + 1
                '�Y����̃A�h���X(��)��ݒ�@���� D:F
                FSrcColAddress(j) = GetColumnsAddress(FSrcRange, LeftCol, i - 1)
                '���̉�̉E���������Z���̋��E�ƂȂ��
                LeftCol = i
            End With
        End If
    Next
    
    '�z��T�C�Y��ݒ�
    FBlockCount = j
    
    '�z���O�l�߂Ɉ��k
    ReDim Preserve FSrcColAddress(1 To FBlockCount)
    ReDim FWidthRatio(1 To FBlockCount)
    For i = 1 To FBlockCount
        '�����Z���̉򂲂Ƃ̑S�̕��ɑ΂��銄����ݒ�
        FWidthRatio(i) = FSrcRange.Worksheet.Range(FSrcColAddress(i)).Width / FSrcRange.Width
    Next
End Sub

'*****************************************************************************
'[�T�v] �����߂�Range���擾����
'[����] �Ώ�Range�A�V������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Function GetDstRange(ByRef objTopLeftCell As Range) As Range
    Dim objWkRange As Range
    Dim DstColCnt  As Long
    
    Set objWkRange = objTopLeftCell.Resize(, Columns.Count - objTopLeftCell.Column + 1)
    DstColCnt = WorksheetFunction.Max(GetColNumber(objWkRange, FSrcRange.Width), FBlockCount)
    Set FDstRange = objTopLeftCell.Resize(FSrcRange.Rows.Count, DstColCnt)
    
    Set GetDstRange = FDstRange
End Function

'*****************************************************************************
'[�T�v] FDstColAddress��ݒ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SetDstColAddress()
    ReDim FDstColAddress(1 To FBlockCount)
    
    ReDim LeftCol(1 To FBlockCount) As Long  '���𔻒肷�鍶�[�̗�ԍ�
    Dim RightCol As Long '���𔻒肷��E�[�̗�ԍ�
    LeftCol(1) = 1
    RightCol = FDstRange.Columns.Count - FBlockCount
    
    Dim dblWidth As Double
    Dim objWkRange As Range
    Dim i As Long
    For i = 1 To FBlockCount - 1
        RightCol = RightCol + 1
        dblWidth = SumRatio(1, i) * FDstRange.Width
        Set objWkRange = FDstRange.Resize(, RightCol)
        LeftCol(i + 1) = GetColNumber(objWkRange, dblWidth) + 1
        LeftCol(i + 1) = WorksheetFunction.Max(LeftCol(i + 1), LeftCol(i) + 1)
    Next
    For i = 1 To FBlockCount - 1
        FDstColAddress(i) = GetColumnsAddress(FDstRange, LeftCol(i), LeftCol(i + 1) - 1)
    Next
    
    '�Ō�̗�͎c��S��
    FDstColAddress(FBlockCount) = GetColumnsAddress(FDstRange, LeftCol(i), FDstRange.Columns.Count)
End Sub

'*****************************************************************************
'[�T�v] StartIdx����EndIdx�܂ł̑S�̕��ɑ΂��銄�������v����
'[����] StartIdx,EndIdx
'[�ߒl] ���̍��v
'*****************************************************************************
Private Function SumRatio(ByVal StartIdx As Long, Optional ByVal EndIdx As Long) As Double
    Dim i As Long
    For i = StartIdx To EndIdx
        SumRatio = SumRatio + FWidthRatio(i)
    Next
End Function

'*****************************************************************************
'[�T�v] �ΏۂƂ���Range��StartColx����EndCol�܂ł̗�A�h���X���擾����
'[����] �ΏۂƂ���Range,Start��,End��
'[�ߒl] �Ώۂ̗�A�h���X�@��FC:E
'*****************************************************************************
Private Function GetColumnsAddress(ByRef objRange As Range, ByVal StartCol As Long, ByVal EndCol As Long) As String
    Dim objWkRange As Range
    With objRange
        Set objWkRange = .Worksheet.Range(.Columns(StartCol), .Columns(EndCol))
    End With
    GetColumnsAddress = objWkRange.EntireColumn.Address(0, 0)
End Function

'*****************************************************************************
'[�T�v] �^���̗�͈͓��ŊY���̕��ɋ߂���ԍ�(objRange����)���擾����
'[����] ��͈́A�擾��
'[�ߒl] ���\���A�h���X�@�� C:E
'*****************************************************************************
Private Function GetColNumber(ByRef objRange As Range, ByVal dblWidth As Double) As Long
    If objRange.Columns.Count = 1 Then
        GetColNumber = 1
        Exit Function
    End If
    
    If objRange.Width <= dblWidth Then
        GetColNumber = objRange.Columns.Count
        Exit Function
    End If
    
    '���𔻒肷��E�ׂ̃Z���̐^��
    Dim dblHalf As Double
    
    Dim i As Long
    For i = 1 To objRange.Columns.Count - 1
        dblHalf = objRange.Columns(i + 1).Width / 2
        If dblWidth <= objRange.Resize(, i).Width + dblHalf Then
            Exit For
        End If
    Next
    GetColNumber = i
End Function

'*****************************************************************************
'[�T�v] �񐔂𑝌�������
'[����] �Ώ�Range�A�V������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SplitOrEraseCol(ByRef objRange As Range, ByVal NewColCount As Long)
    Dim OldColCount As Long
    OldColCount = objRange.Columns.Count
    
    If OldColCount = NewColCount Then
        Exit Sub
    End If

    If NewColCount > OldColCount Then
        With objRange
            Call SplitCol(.Columns(OldColCount), NewColCount - OldColCount + 1)
        End With
    Else
        With objRange
            Call EraseCol(.Worksheet.Range(.Columns(NewColCount + 1), .Columns(OldColCount)))
        End With
    End If
End Sub

'*****************************************************************************
'[�T�v] ��𕪊�����
'[����] �Ώۗ�A������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SplitCol(ByRef objRange As Range, ByVal SplitCount As Long)
    Dim objNewCol As Range
    Dim i As Long
    
    '�I���̉E����1��}��
    Call objRange.Columns(2).Insert(xlShiftToRight, xlFormatFromLeftOrAbove)
    
    '�V������
    Set objNewCol = objRange.Columns(2)
    
    '�}����̂P�Z�����Ɍr�����R�s�[����
    Call CopyBorder("�E�㉺", objRange, objNewCol)
    
    '�������Ɍ�������
    Dim objMergeRange As Range
'    Dim lngtype As Long
'    If chkIgnore.Value Then
'        lngtype = 2 '��������Ă��Ȃ��Z���͌������Ȃ�
'    Else
'        lngtype = 1 '���ׂĉ������Ɍ�������
'    End If
    
    '1�s����Loop
    For i = 1 To objRange.Rows.Count
        If objNewCol.Cells(i, 1).MergeArea.Count = 1 Then
            Set objMergeRange = GetMergeColRange(1, objRange.Cells(i, 1), objNewCol.Cells(i, 1))
            If Not (objMergeRange Is Nothing) Then
                Call objMergeRange.Merge
            End If
        End If
    Next
    
    '�����������A���}������
    For i = 3 To SplitCount
        Call objNewCol.EntireColumn.Insert
    Next
End Sub

'*****************************************************************************
'[�T�v] �����������
'[����] �Ώۗ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub EraseCol(ByRef objRange As Range)
    With objRange
        '�E�[�̌r�����R�s�[����
        Call CopyBorder("�E", .Columns(.Columns.Count), .Columns(0))
    
        '�폜
        Call .Delete(xlShiftToLeft)
    End With
End Sub

'*****************************************************************************
'[�T�v] ��������Ă��Ȃ��Z���͖�������`�F�b�N��
'*****************************************************************************
Private Sub chkIgnore_Click()
    If Not FEnableEvents Then Exit Sub
    
    If chkWidth.Value Then
        chkWidth.Value = False
    End If
    
    Dim i As Long
    For i = LBound(objTextbox) To UBound(objTextbox)
        If Not (objTextbox(i) Is Nothing) Then
            Call objTextbox(i).Delete
            Set objTextbox(i) = Nothing
        End If
    Next
    
    Call Init(FSrcRange)
    Call GetDstRange(FDstRange)
    Call FDstRange.Select
    Call SetDstColAddress
    Call MakeTextBox
End Sub

'*****************************************************************************
'[�T�v] �R�s�[���̕����Č�����`�F�b�N��
'*****************************************************************************
Private Sub chkWidth_Click()
    Static BackupDstRange As Range
    Application.ScreenUpdating = False
        
    Dim i As Long
    Dim ColWidth As Double
    If chkWidth.Value Then
        ReDim FBackupDstColWidth(1 To FBlockCount)
        Set BackupDstRange = FDstRange
        Set FDstRange = FDstRange.Resize(, FBlockCount)
        For i = 1 To FBlockCount
            ColWidth = FSrcRange.Width * FWidthRatio(i) / DPIRatio
            FBackupDstColWidth(i) = FDstRange.Columns(i).EntireColumn.ColumnWidth
            FDstRange.Columns(i).EntireColumn.ColumnWidth = PixelToWidth(ColWidth)
        Next
        If FSrcRange.Rows.Count > 1 Then
            With objTextbox(FBlockCount + 1).TextFrame2.TextRange
                .Characters.Text = "�R�s�[���̕����Č����鎞�́A����ύX�ł��܂���"
            End With
        End If
    Else
        For i = 1 To FBlockCount
            FDstRange.Columns(i).EntireColumn.ColumnWidth = FBackupDstColWidth(i)
        Next
        Set FDstRange = BackupDstRange
        If FSrcRange.Rows.Count > 1 Then
            With objTextbox(FBlockCount + 1).TextFrame2.TextRange
                .Characters.Text = ""
            End With
        End If
    End If
    
    '�Ώۗ�A�h���X�̍Đݒ�
    Call SetDstColAddress
    '���o����ɃR�s�[���̗�A�h���X��\��
    Call EditTextbox
    
    Call FDstRange.Select
    Application.ScreenUpdating = True
End Sub

'*****************************************************************************
'[�T�v] �t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    '�Ăь��ɒʒm����
    blnFormLoad = True
    lngZoom = ActiveWindow.Zoom
End Sub

'*****************************************************************************
'[�T�v] �t�H�[���A�����[�h��
'*****************************************************************************
Private Sub UserForm_Terminate()
    '�Ăь��ɒʒm����
    blnFormLoad = False

    Dim i As Long
    For i = LBound(objTextbox) To UBound(objTextbox)
        If Not (objTextbox(i) Is Nothing) Then
            Call objTextbox(i).Delete
        End If
    Next
    
    ActiveWorkbook.DisplayDrawingObjects = lngDisplayObjects
    ActiveWindow.Zoom = lngZoom
End Sub

'*****************************************************************************
'[�T�v] �n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrHandle
    Dim blnCopyObjectsWithCells  As Boolean
    blnCopyObjectsWithCells = Application.CopyObjectsWithCells
    Application.CopyObjectsWithCells = False '�Ăь��ŕ������邽�ߓ����W���[���ł͕������Ȃ�

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False '�R�����g�����鎞�x�����o�鎞������

    Dim i As Long
    For i = LBound(objTextbox) To UBound(objTextbox)
        If Not (objTextbox(i) Is Nothing) Then
            Call objTextbox(i).Delete
            Set objTextbox(i) = Nothing
        End If
    Next
    
    '�A���h�D�p�Ɍ��̕��𕜌�����
    If chkWidth.Value Then
        ReDim SaveDstColWidth(1 To FBlockCount) As Double
        For i = 1 To FBlockCount
            SaveDstColWidth(i) = FDstRange.Columns(i).EntireColumn.ColumnWidth
        Next
        Call BackupDstColWidth
    End If
    
    '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
    If FMove Then
        Call SaveUndoInfo(E_CopyCell, FSrcRange)
    Else
        Call SaveUndoInfo(E_CopyCell, FDstRange)
    End If

    '�����ĕ�������
    If chkWidth.Value Then
        For i = 1 To FBlockCount
            FDstRange.Columns(i).EntireColumn.ColumnWidth = SaveDstColWidth(i)
        Next
    End If
    
    '�Z�����R�s�[����
    Call CopyCell

    Call FDstRange.Select
    
    Application.DisplayAlerts = True
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea2"))
    Call Unload(Me)
    Call SetOnUndo
Exit Sub
ErrHandle:
    Application.DisplayAlerts = True
    Call MsgBox(Err.Description, vbExclamation)
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea1"))
    Call DeleteSheet(ThisWorkbook.Worksheets("Workarea2"))
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�T�v] �L�����Z���{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call BackupDstColWidth
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�T�v] �w���v�{�^��������
'*****************************************************************************
Private Sub cmdHelp_Click()
    Call OpenHelpPage("http://takana.web5.jp/EasyLout/V5/Clipbord.htm#PasteAppearance")
End Sub

'*****************************************************************************
'[�T�v] �R�s�[���̕����Č����邪�`�F�b�N����Ă��鎞�A���̕����Č�����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub BackupDstColWidth()
    Dim i As Long
    If chkWidth.Value Then
        For i = 1 To FBlockCount
            FDstRange.Columns(i).EntireColumn.ColumnWidth = FBackupDstColWidth(i)
        Next
    End If
End Sub

'*****************************************************************************
'[�T�v] �̈���R�s�[����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub CopyCell()
    Dim objWkRange As Range
    
    '�̈�����[�N�V�[�g�ɃR�s�[����
    Set objWkRange = CopyToWkSheet()

    '�ړ����̓R�s�[�����N���A����
    If FMove Then
        With FSrcRange
            .ClearContents
            .UnMerge
            .Borders.LineStyle = xlNone
        End With
    End If
    
    Dim objRange As Range
    Dim lngOffset As Long
    lngOffset = FDstRange.Column - FSrcRange.Column
    '���[�N�V�[�g�ŗ̈�̃T�C�Y��ύX����
    Dim i As Long
    For i = FBlockCount To 1 Step -1
        Set objRange = Intersect(objWkRange, objWkRange.Worksheet.Range(FSrcColAddress(i)).Offset(, lngOffset))
        Call SplitOrEraseCol(objRange, Range(FDstColAddress(i)).Columns.Count)
    Next
    
    Call objWkRange.Resize(, FDstRange.Columns.Count).Copy(FDstRange)
    If chkBlank.Value Then
        Set objWkRange = GetBlankAndMergeRange(FDstRange)
        If Not (objWkRange Is Nothing) Then
            Call objWkRange.UnMerge
            objWkRange.HorizontalAlignment = xlGeneral
        End If
    End If
End Sub

'*****************************************************************************
'[�T�v] �� ���� �������ꂽ�Z�����擾����
'[����] �Ώۗ̈�
'[�ߒl] �� ���� �������ꂽ�Z��
'*****************************************************************************
Private Function GetBlankAndMergeRange(ByRef objSelection As Range) As Range
    Dim objRange   As Range
    Dim objCell    As Range
    
    '�������ꂽ�Z����UsedRange�ȊO�ɂ͂Ȃ��̂�
    Set objRange = IntersectRange(objSelection, GetUsedRange())
    If objRange Is Nothing Then
        Exit Function
    End If
    
    '�Z���̐��������[�v
    For Each objCell In objRange
        '�󔒂��H
        If objCell.Value = "" And objCell.Formula = "" Then
            With objCell.MergeArea
                '�����Z�����H
                If .Count > 1 Then
                    '����̃Z����
                    If .Row = objCell.Row And .Column = objCell.Column Then
                        Set GetBlankAndMergeRange = UnionRange(GetBlankAndMergeRange, objCell)
                    End If
                End If
            End With
        End If
    Next
End Function

'*****************************************************************************
'[�C�x���g] KeyDown
'[�T�v] �J�[�\���L�[�ňړ����ύX������
'*****************************************************************************
Private Sub cmdOK_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub cmdCancel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub cmdHelp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub fraKeyCapture_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub chkIgnore_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub chkWidth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub

'*****************************************************************************
'[�C�x���g] UserForm_KeyDown
'[�T�v] �J�[�\���L�[�ňړ����ύX������
'*****************************************************************************
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i         As Long
    Dim lngTop    As Long
    Dim lngLeft   As Long
    Dim lngBottom As Long
    Dim lngRight  As Long

    Select Case (KeyCode)
    Case vbKeyLeft, vbKeyRight, vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyUp, vbKeyDown
        Call fraKeyCapture.SetFocus
    Case Else
        Exit Sub
    End Select

    'Alt��������Ă���΃X�N���[��
    If GetKeyState(vbKeyMenu) < 0 Then
        Select Case (KeyCode)
        Case vbKeyLeft
            Call ActiveWindow.SmallScroll(, , , 1)
        Case vbKeyRight
            Call ActiveWindow.SmallScroll(, , 1)
        Case vbKeyUp
            Call ActiveWindow.SmallScroll(, 1)
        Case vbKeyDown
            Call ActiveWindow.SmallScroll(1)
        End Select
        Exit Sub
    End If

    'Zoom
    Select Case (KeyCode)
    Case vbKeyHome, vbKeyPageUp, vbKeyPageDown
        Select Case (KeyCode)
        Case vbKeyHome
            ActiveWindow.Zoom = lngZoom
        Case vbKeyPageUp
            ActiveWindow.Zoom = WorksheetFunction.Min(ActiveWindow.Zoom + 10, 400)
        Case vbKeyPageDown
            ActiveWindow.Zoom = WorksheetFunction.Max(ActiveWindow.Zoom - 10, 10)
        End Select

        With FDstRange
            lngLeft = WorksheetFunction.Max(.Left / DPIRatio - 1, 0) * ActiveWindow.Zoom / 100
            lngTop = WorksheetFunction.Max(.Top / DPIRatio - 1, 0) * ActiveWindow.Zoom / 100
            Call ActiveWindow.ScrollIntoView(lngLeft, lngTop, 1, 1)
        End With
        Exit Sub
    End Select
    
    '�R�s�[���̕����Č����鎞�́A�ύX�s��
    If chkWidth.Value Then
        Exit Sub
    End If
    
    '�I��̈�̎l���̈ʒu��Ҕ�
    With FDstRange
        lngTop = .Row
        lngBottom = .Row + .Rows.Count - 1
        lngLeft = .Column
        lngRight = .Column + .Columns.Count - 1
    End With

    Select Case (Shift)
    Case 0
        '�I��̈���ړ�
        Select Case (KeyCode)
        Case vbKeyLeft
            lngLeft = lngLeft - 1
            lngRight = lngRight - 1
        Case vbKeyRight
            lngLeft = lngLeft + 1
            lngRight = lngRight + 1
        End Select
    Case 1
        '�I��̈�̑傫����ύX
        If GetKeyState(vbKeyZ) < 0 Then
            Select Case (KeyCode)
            Case vbKeyLeft
                lngLeft = lngLeft - 1
            Case vbKeyRight
                lngLeft = lngLeft + 1
            End Select
        Else
            Select Case (KeyCode)
            Case vbKeyLeft
                lngRight = lngRight - 1
            Case vbKeyRight
                lngRight = lngRight + 1
            End Select
        End If
    Case Else
        Exit Sub
    End Select

    '�`�F�b�N
    If (FBlockCount <= lngRight - lngLeft + 1) And _
       (1 <= lngLeft And lngRight <= Columns.Count) Then
        Set FDstRange = Range(Cells(lngTop, lngLeft), Cells(lngBottom, lngRight))
        Call SetDstColAddress
    Else
        Exit Sub
    End If
    
    '�e�L�X�g�{�b�N�X��ҏW
    Call EditTextbox

    '�I��̈悪��ʂ�����������ʂ��X�N���[��
    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '��ʕ����̂Ȃ���
        Select Case (KeyCode)
        Case vbKeyLeft
            i = WorksheetFunction.Max(FDstRange.Column - 1, 1)
            If IntersectRange(ActiveWindow.VisibleRange, Columns(i)) Is Nothing Then
                Call ActiveWindow.SmallScroll(, , , 1)
            End If
        Case vbKeyRight
            i = WorksheetFunction.Min(FDstRange.Column + FDstRange.Columns.Count, Columns.Count)
            If IntersectRange(ActiveWindow.VisibleRange, Columns(i)) Is Nothing Then
                Call ActiveWindow.SmallScroll(, , 1)
            End If
        End Select
    End If
End Sub

'*****************************************************************************
'[�T�v] �\�t���\���ǂ����`�F�b�N
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub CheckPaste()
On Error GoTo ErrHandle
    If CheckBorder() = True Then
        Call Err.Raise(C_CheckErrMsg, , "�������ꂽ�Z���̈ꕔ��ύX���邱�Ƃ͂ł��܂���")
    End If
    If FDstRange.Rows.Count > 1 And chkWidth.Value = False Then
        With objTextbox(FBlockCount + 1).TextFrame.Characters
            .Text = ""
        End With
    End If
    cmdOK.Enabled = True  'OK�{�^�����g�p�ɂ���
Exit Sub
ErrHandle:
    If Err.Number = C_CheckErrMsg Then
        If FSrcRange.Rows.Count > 1 Then
            With objTextbox(FBlockCount + 1).TextFrame.Characters
                .Text = Err.Description
            End With
        End If
        cmdOK.Enabled = False 'OK�{�^�����g�p�s�ɂ���
    Else
        Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If
End Sub

'*****************************************************************************
'[�T�v] �\�t����̎l���������Z�����܂����ł��邩�ǂ���
'[����] �Ȃ�
'[�ߒl] True:�����Z�����܂����ł���AFalse:���Ȃ�
'*****************************************************************************
Private Function CheckBorder() As Boolean
    Dim objChkRange As Range
    If FMove Then
        Set objChkRange = MinusRange(ArrangeRange(FDstRange), UnionRange(FSrcRange, FDstRange))
    Else
        Set objChkRange = MinusRange(ArrangeRange(FDstRange), FDstRange)
    End If
    CheckBorder = Not (objChkRange Is Nothing)
End Function

'*****************************************************************************
'[�T�v] ���[�N�V�[�g�ɗ̈���R�s�[����
'[����] �Ȃ�
'[�ߒl] ���[�N�V�[�g�̃R�s�[���ꂽRange(�T�C�Y�̓R�s�[���̃T�C�Y)
'*****************************************************************************
Private Function CopyToWkSheet() As Range
    Dim objWorksheet As Worksheet

    Set objWorksheet = ThisWorkbook.Worksheets("Workarea1")
    Call DeleteSheet(objWorksheet)

    '�uWorkarea�v�V�[�g�ɑI��̈�𕡎ʂ���
    With FSrcRange
        Set CopyToWkSheet = objWorksheet.Range(FDstRange.Address).Resize(.Rows.Count, .Columns.Count)
        Call .Copy(CopyToWkSheet)
    End With
End Function

'*****************************************************************************
'[�T�v] Range��Excel���ᎆ���ǂ�������
'[����] ���肷��Range
'[�ߒl] True:���ᎆ
'*****************************************************************************
Public Function IsGraphpaper(ByRef objRange As Range) As Boolean
    Dim lngColCnt As Long
    Dim lngRowCnt As Long
    Dim i As Long
    
    '���Z���̐�
    For i = 1 To objRange.Columns.Count
        If Not objRange.Columns(i).Hidden Then
            lngColCnt = lngColCnt + 1
        End If
    Next
    For i = 1 To objRange.Rows.Count
        If Not objRange.Rows(i).Hidden Then
            lngRowCnt = lngRowCnt + 1
        End If
    Next
    
    '1�Z���̕��̕��ς��A1�Z���̍����̕��ς�2�{�ȉ��̎��͕��ᎆ�Ƃ���
    If objRange.Width / lngColCnt <= objRange.Height * 2 / lngRowCnt Then
        IsGraphpaper = True
    Else
        IsGraphpaper = False
    End If
End Function





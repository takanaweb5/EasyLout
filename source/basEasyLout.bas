Attribute VB_Name = "basEasyLout"
Option Explicit
Option Private Module

Private Type PICTDESC_BMP
    Size    As Long
    Type    As Long
    hBitmap As LongPtr
    hPal    As LongPtr
End Type

Private Type Guid
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpszCLSID As LongPtr, ByRef pclsid As Guid) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef PicDesc As Any, ByRef RefIID As Guid, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long

Private Const PICTYPE_BITMAP = 1
Private Const IID_IPictureDisp As String = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"

Public FCommand As String

'*****************************************************************************
'[�T�v] IRibbonUI��ۑ�����CommandBar���쐬����
'       ���킹�āA���{���R���g���[���̏�Ԃ�ۑ�����CommandBarControl���쐬����
'       ���W���[���ϐ��ɕۑ������ꍇ�́A���R���p�C����R�[�h�̋�����~�Œl�����Ȃ��邽��
'[����] IRibbonUI
'[�ߒl] �Ȃ�
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
    
    '�`�F�b�N�{�b�N�X�̃N���[�����e���|�����ɍ쐬
    For i = 1 To 1
        With objCmdBar.Controls.Add(msoControlButton)
            .Tag = "C" & i & ThisWorkbook.Name
            .State = False '�����ݒ�̓`�F�b�N�Ȃ�
        End With
    Next
End Sub

'*****************************************************************************
'[�T�v] CommandBar����IRibbonUI���擾����
'[����] �Ȃ�
'[�ߒl] IRibbonUI
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
'[�C�x���g] onLoad
'*****************************************************************************
Sub onLoad(Ribbon As IRibbonUI)
    '���{��UI���e���|�����̃R�}���h�o�[�ɕۑ�����
    '(���W���[���ϐ��ɕۑ������ꍇ�́A��O��R�[�h�̋�����~�Œl�����Ȃ��邽��)
    Call CreateTmpCommandBar(Ribbon)
End Sub

'*****************************************************************************
'[�C�x���g] loadImage
'*****************************************************************************
Sub loadImage(imageID As String, ByRef returnedVal)
  returnedVal = imageID
End Sub

'*****************************************************************************
'[�C�x���g] getVisible
'*****************************************************************************
Sub getVisible(Control As IRibbonControl, ByRef returnedVal)
'    returnedVal = (GetValue(control.Id, "Visible") = 1)
    returnedVal = True
End Sub
'
'*****************************************************************************
'[�C�x���g] getEnabled
'*****************************************************************************
Sub getEnabled(Control As IRibbonControl, ByRef returnedVal)
    Select Case Control.ID
    Case "B311", "B312"
        returnedVal = (CheckSelection() = E_Shape)
    Case "B313", "B314", "B315", "B316"
        returnedVal = CommandBars.GetEnabledMso("ObjectsAlignTop")
    Case "B321" '�Z�����e�L�X�g�{�b�N�X�ɕϊ�
        returnedVal = (CheckSelection() = E_Range)
    Case "B322" '�e�L�X�g�{�b�N�X���Z���ɕϊ�
        returnedVal = (CheckSelection() = E_Shape)
    Case "B323" '�R�����g���e�L�X�g�{�b�N�X�ɕϊ�
        returnedVal = IsCommnetSelect
    Case "B324" '�e�L�X�g�{�b�N�X���R�����g�ɖ߂�
        returnedVal = IsSelectCommentTextbox
    Case "B325" '�R�����g����͋K���ɕϊ�
        returnedVal = IsCommnetSelect
    Case "B326" '���͋K�����R�����g�ɕϊ�
        returnedVal = Not (GetInputRules(True) Is Nothing)
    Case Else
        returnedVal = True
    End Select
End Sub

'*****************************************************************************
'[�T�v] �R�����g���I������Ă��邩�ǂ���
'[����] �Ȃ�
'[�ߒl] True:�R�����g���I������Ă���
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
'[�C�x���g] getShowLabel
'*****************************************************************************
Sub getShowLabel(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = (GetValue(Control.ID, "ShowLabel") = 1)
End Sub

'*****************************************************************************
'[�C�x���g] getLabel
'*****************************************************************************
Sub getLabel(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetValue(Control.ID, "Label")
    
    Select Case Control.ID
    Case "B5311"
        returnedVal = Replace(returnedVal, "{FONTNAME}", GetSetting(REGKEY, "KEY", "FontName", DEFAULTFONT))
    End Select
End Sub

'*****************************************************************************
'[�C�x���g] getScreentip
'*****************************************************************************
Sub getScreentip(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetValue(Control.ID, "Screentip")
    
    Select Case Control.ID
    Case "B5311"
        returnedVal = Replace(returnedVal, "{FONTNAME}", GetSetting(REGKEY, "KEY", "FontName", DEFAULTFONT))
    End Select
End Sub
'*****************************************************************************
'[�C�x���g] getSupertip
'*****************************************************************************
Sub getSupertip(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetValue(Control.ID, "Supertip")

    Select Case Control.ID
    Case "B5311"
        returnedVal = Replace(returnedVal, "{FONTNAME}", GetSetting(REGKEY, "KEY", "FontName", DEFAULTFONT))
    End Select
End Sub

'*****************************************************************************
'[�C�x���g] getShowImage
'*****************************************************************************
Sub getShowImage(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = (GetValue(Control.ID, "ShowImage") = 1)
End Sub

'*****************************************************************************
'[�C�x���g] getImage
'*****************************************************************************
Sub getImage(Control As IRibbonControl, ByRef returnedVal)
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
'[�C�x���g] getSize
'*****************************************************************************
Sub getSize(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetValue(Control.ID, "ButtonSize")
End Sub

'*****************************************************************************
'[�C�x���g] getContent ���I�Ƀ��j���[���쐬����
'*****************************************************************************
Sub getContent(Control As IRibbonControl, ByRef returnedVal) '
On Error Resume Next
    Select Case Control.ID
    Case "M31"
        returnedVal = GetRangeText(ThisWorkbook.Worksheets("dynamicMenu").Range("A1:A21"))
    Case "M32"
        returnedVal = GetRangeText(ThisWorkbook.Worksheets("dynamicMenu").Range("A23:A32"))
    Case "M531"
        returnedVal = GetRangeText(ThisWorkbook.Worksheets("dynamicMenu").Range("A34:A38"))
    End Select
End Sub

'*****************************************************************************
'[�C�x���g] getPressed
'*****************************************************************************
Sub getPressed(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = False
    Select Case Control.ID
    Case "C1"
        On Error Resume Next
        returnedVal = (ActiveWorkbook.DisplayDrawingObjects = xlHide)
    End Select
End Sub

'*****************************************************************************
'[�C�x���g] onCheckAction
'*****************************************************************************
Sub onCheckAction(Control As IRibbonControl, pressed As Boolean)
    '�`�F�b�N��Ԃ�ۑ�
    GetTmpControl(Control.ID).State = pressed
    
    Select Case Control.ID
    Case "C1"
        Call HideShapes(pressed)
    End Select
End Sub

'*****************************************************************************
'[�T�v] �}�`��\���E��\�����g�O��������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ToggleHideShapes()
On Error GoTo ErrHandle
    Dim blnHide As Boolean
    blnHide = Not (ActiveWorkbook.DisplayDrawingObjects = xlHide)
    Call HideShapes(blnHide)
    GetTmpControl("C1").State = blnHide
    Call GetRibbonUI.InvalidateControl("C1")
ErrHandle:
End Sub

'*****************************************************************************
'[�C�x���g] onAction
'*****************************************************************************
Sub onAction(Control As IRibbonControl)
    Call SetChkBox
    
    Dim Param
    Param = GetValue(Control.ID, "Parameter")
    On Error Resume Next
    If Param <> "" Then
        Call Application.Run(GetValue(Control.ID, "Action"), Param)
    Else
        Call Application.Run(GetValue(Control.ID, "Action"))
    End If
End Sub

'*****************************************************************************
'[�T�v] �`�F�b�N�{�b�N�X�̃`�F�b�N��ݒ肷��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SetChkBox()
    If ActiveWorkbook Is Nothing Then
        GetTmpControl("C1").State = False
    Else
        GetTmpControl("C1").State = (ActiveWorkbook.DisplayDrawingObjects = xlHide)
    End If
    Call GetRibbonUI.InvalidateControl("C1")
End Sub

'*****************************************************************************
'[�T�v] Commands�V�[�g����Y���̒l���擾����
'[����] �R���g���[��Id�A���ږ�
'[�ߒl] IPicture
'*****************************************************************************
Private Function GetValue(ByVal ID As String, ByVal strCol As String) As Variant
    Dim x As Long
    Dim y As Long
    
    GetValue = ""
    
    Static vValues As Variant
    If VarType(vValues) = vbEmpty Then
        vValues = ThisWorkbook.Worksheets("Commands").UsedRange
    End If
        
    '��LOOP
    For x = 1 To UBound(vValues, 2)
        If vValues(1, x) = strCol Then
            '�s��LOOP
            For y = 2 To UBound(vValues, 1)
                If vValues(y, 3) = ID Then
                    GetValue = vValues(y, x)
                    Exit Function
                End If
            Next
        End If
    Next

'    Dim objRange As Range
'    Set objRange = ThisWorkbook.Worksheets("Commands").UsedRange
'
'    '��LOOP
'    For x = 1 To objRange.Columns.Count
'        If objRange.Cells(1, x).Value = strCol Then
'            Exit For
'        End If
'    Next
'
'    '�s��LOOP
'    For y = 2 To objRange.Rows.Count
'        If objRange.Cells(y, "C").Value = Id Then
'            Exit For
'        End If
'    Next

'    GetValue = objRange.Cells(y, x).Value
End Function

'*****************************************************************************
'[�T�v] �Z���̃f�[�^����A�C�R���t�@�C����Ǎ���
'[����] Resource�V�[�g�̂̃f�[�^�̃t�@�C����(A��̒l)
'[�ߒl] IPicture
'*****************************************************************************
Private Function GetImageFromResource(ByVal strImageFile As String) As IPicture
    Dim y As Long
    Dim objRange As Range
    Dim objRow As Range
    Set objRange = ThisWorkbook.Worksheets("Resource").UsedRange
    
    Dim strPixel As String
    'DPI���� 16Pixel or 20Pixel�̃A�C�R�����g�p���邩���� ��Windows�̕W����96DPI
    Select Case GetDPI()
    Case Is < 120
        '��FSample.png �� Sample16.png
        strImageFile = Replace(strImageFile, ".png", 16 & ".png")
    Case 120
        '��FSample.png �� Sample20.png
        strImageFile = Replace(strImageFile, ".png", 20 & ".png")
    End Select
        
    '�s��LOOP
    For y = 1 To objRange.Rows.Count
        If objRange.Cells(y, "A").Value = strImageFile Then
            Set objRow = objRange.Rows(y)
            Exit For
        End If
    Next
    
    If objRow Is Nothing Then Exit Function
On Error GoTo ErrHandle
    Set GetImageFromResource = LoadImageFromResource(objRow)
ErrHandle:
End Function

'*****************************************************************************
'[�T�v] �Z���̃f�[�^����A�C�R���t�@�C����Ǎ���
'[����] �f�[�^���擾����s(Range�I�u�W�F�N�g)
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Function LoadImageFromResource(ByRef objRow As Range) As IPicture
    '�t�@�C���T�C�Y�̔z����쐬
    ReDim data(1 To objRow.Cells(1, 1).End(xlToRight).Column - 1) As Byte
    Dim x As Long
    For x = 1 To UBound(data)
         data(x) = objRow.Cells(1, x + 1)
    Next
    
    Dim Stream As IUnknown
    If CreateStreamOnHGlobal(VarPtr(data(1)), 0, Stream) <> 0 Then
        Call Err.Raise(513, "CreateStreamOnHGlobal�G���[")
    End If

    Dim Gdip As New CGdiplus
    Call Gdip.CreateFromStream(Stream)
    
    Dim uPicInfo As PICTDESC_BMP
    With uPicInfo
        .Size = Len(uPicInfo)
        .Type = PICTYPE_BITMAP
        .hBitmap = Gdip.ToHBITMAP
        .hPal = 0
    End With

    Dim gGuid As Guid
    Call CLSIDFromString(StrPtr(IID_IPictureDisp), gGuid)
    Dim IBitmap As IPicture
    Call OleCreatePictureIndirect(uPicInfo, gGuid, True, IBitmap)
    Set LoadImageFromResource = IBitmap
End Function

'*****************************************************************************
'[�T�v] ���{���̃R�[���o�b�N�֐������s����(Debug�p)
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Sub onAction2(Control As IRibbonControl)
    Select Case Control.ID
    Case "Bdmy1"
        Call GetRibbonUI.Invalidate
    Case "Bdmy2"
        ThisWorkbook.IsAddin = Not ThisWorkbook.IsAddin
    Case "Bdmy3"
        Call ThisWorkbook.Save
    End Select
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGdiplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

'Private Enum GpStatus
'   Ok = 0
'   GenericError = 1
'   InvalidParameter = 2
'   OutOfMemory = 3
'   ObjectBusy = 4
'   InsufficientBuffer = 5
'   NotImplemented = 6
'   Win32Error = 7
'   WrongState = 8
'   Aborted = 9
'   FileNotFound = 10
'   ValueOverflow = 11
'   AccessDenied = 12
'   UnknownImageFormat = 13
'   FontFamilyNotFound = 14
'   FontStyleNotFound = 15
'   NotTrueTypeFont = 16
'   UnsupportedGdiplusVersion = 17
'   GdiplusNotInitialized = 18
'   PropertyNotFound = 19
'   PropertyNotSupported = 20
'End Enum

Private Type Guid
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

Private Type BITMAPINFOHEADER
    Size          As Long
    Width         As Long
    Height        As Long
    Planes        As Integer
    BitCount      As Integer
    Compression   As Long
    SizeImage     As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed       As Long
    ClrImportant  As Long
End Type

Private Type BITMAPDATA
    Width       As Long
    Height      As Long
    stride      As Long
    PixelFormat As Long
    scan0       As LongPtr
    Reserved    As LongPtr
End Type

Private Const DIB_RGB_COLORS = 0&
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDc As LongPtr) As Long
Private Declare PtrSafe Function ReleaseDC Lib "User32" (ByVal hWnd As LongPtr, ByVal hDc As LongPtr) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmapptr As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As Any, ByVal wUsage As Long) As Long

Private Declare PtrSafe Function GdipSaveImageToStream Lib "GDIPlus" (ByVal Image As LongPtr, ByVal Stream As LongPtr, ByRef clsidEncoder As Guid, ByVal encoderParams As LongPtr) As Long
Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus.dll" (ByVal Image As LongPtr, ByVal Filename As LongPtr, ByRef clsidEncoder As Guid, ByVal encoderParams As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageHeight Lib "GDIPlus" (ByVal Image As LongPtr, Height As Long) As Long
Private Declare PtrSafe Function GdipGetImageWidth Lib "GDIPlus" (ByVal Image As LongPtr, Width As Long) As Long
Private Declare PtrSafe Function GdipBitmapGetPixel Lib "GDIPlus" (ByVal bitmap As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Color As Long) As Long
Private Declare PtrSafe Function GdipBitmapSetPixel Lib "GDIPlus" (ByVal bitmap As LongPtr, ByVal x As Long, ByVal y As Long, ByVal Color As Long) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus.dll" (ByVal nWidth As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByRef scan0 As Any, ByRef nBitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hbm As LongPtr, ByVal hPal As LongPtr, ByRef bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal Filename As LongPtr, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As LongPtr, hbmReturn As LongPtr, ByVal background As Long) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As LongPtr) As Long
Private Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" (ByRef token As LongPtr, ByRef inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "GDIPlus" (ByVal token As LongPtr) As Long
Private Declare PtrSafe Function GdipLoadImageFromFile Lib "GDIPlus" (ByVal Filename As LongPtr, ByRef Image As LongPtr) As Long
Private Declare PtrSafe Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal Stream As LongPtr, ByRef Image As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateSolidFill Lib "GDIPlus" (ByVal pColor As Long, ByVal brush As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "GDIPlus" (ByVal Image As LongPtr, graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipFillRectangle Lib "GDIPlus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
Private Declare PtrSafe Function GdipSetSmoothingMode Lib "GDIPlus" (ByVal mGraphics As LongPtr, ByVal mSmoothingMode As Long) As Long
Private Declare PtrSafe Function GdipDeleteBrush Lib "GDIPlus" (ByVal mBrush As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteGraphics Lib "GDIPlus" (ByVal graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipBitmapSetResolution Lib "GDIPlus" (ByVal pbitmap As LongPtr, ByVal xdpi As Single, ByVal ydpi As Single) As Long
Private Declare PtrSafe Function GdipCreatePen1 Lib "GDIPlus" (ByVal pColor As Long, ByVal Width As Long, ByVal unit As Long, ByRef hPen As LongPtr) As Long
Private Declare PtrSafe Function GdipDrawRectangle Lib "GDIPlus" (ByVal hGraphics As LongPtr, ByVal hPen As LongPtr, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
Private Declare PtrSafe Function GdipDeletePen Lib "GDIPlus" (ByVal hPen As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHICONFromBitmap Lib "GDIPlus" (ByVal bitmap As LongPtr, ByRef hIcon As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageThumbnail Lib "GDIPlus" (ByVal Image As LongPtr, ByVal thumbWidth As Long, ByVal thumbHeight As Long, ByRef thumbImage As LongPtr, ByVal callback As Long, ByRef callbackData As Any) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromGdiDib Lib "GDIPlus" (ByRef gdiBitmapInfo As Any, ByRef gdiBitmapData As Any, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal GpGraphics As LongPtr, ByVal GpImage As LongPtr, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipSetInterpolationMode Lib "GDIPlus" (ByVal graphics As LongPtr, ByVal nInterpolationMode As Long) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromGraphics Lib "GDIPlus" (ByVal Width As Long, ByVal Height As Long, ByVal Target As LongPtr, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipSetPixelOffsetMode Lib "GDIPlus" (ByVal graphics As LongPtr, ByVal PixOffsetMode As Long) As Long
Private Declare PtrSafe Function GdipBitmapLockBits Lib "GDIPlus" (ByVal bitmap As LongPtr, ByRef RECT As Any, ByVal flags As Long, ByVal PixelFormat As Long, ByRef lockedBitmapData As Any) As Long
Private Declare PtrSafe Function GdipBitmapUnlockBits Lib "GDIPlus" (ByVal bitmap As LongPtr, ByRef lockedBitmapData As Any) As Long
Private Declare PtrSafe Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hDc As LongPtr, ByRef graphics As LongPtr) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpszCLSID As LongPtr, ByRef pclsid As Guid) As Long

Private Enum ImageLockMode
    ImageLockModeRead = 1
    ImageLockModeWrite = 2
    ImageLockModeUserInputBuf = 4
End Enum

Private Enum InterpolationMode
    InterpolationModeInvalid = -1
    InterpolationModeDefault = 0&
    InterpolationModeLowQuality = 1&
    InterpolationModeHighQuality = 2&
    InterpolationModeBilinear = 3&
    InterpolationModeBicubic = 4&
    InterpolationModeNearestNeighbor = 5&
    InterpolationModeHighQualityBilinear = 6&
    InterpolationModeHighQualityBicubic = 7&
End Enum

Private Const CLSID_Suffix As String = "-1A04-11D3-9A73-0000F81EF32E}"
Private Const CLSID_BMP    As String = "{557CF400" & CLSID_Suffix
Private Const CLSID_JPG    As String = "{557CF401" & CLSID_Suffix
Private Const CLSID_GIF    As String = "{557CF402" & CLSID_Suffix
Private Const CLSID_EMF    As String = "{557CF403" & CLSID_Suffix
Private Const CLSID_WMF    As String = "{557CF404" & CLSID_Suffix
Private Const CLSID_TIF    As String = "{557CF405" & CLSID_Suffix
Private Const CLSID_PNG    As String = "{557CF406" & CLSID_Suffix
Private Const CLSID_ICO    As String = "{557CF407" & CLSID_Suffix
Private Const PixelOffsetModeHalf = 4
Private Const PixelFormat32bppARGB = &H26200A
Private Const S_OK = 0

Private FGdiPlus As LongPtr
Private FGdiImage As LongPtr

'*****************************************************************************
'[�T�v] �R���X�g���N�^
'*****************************************************************************
Private Sub Class_Initialize()
    Dim GdiInput As GdiplusStartupInput
    GdiInput.GdiplusVersion = 1&
    If GdiplusStartup(FGdiPlus, GdiInput) <> S_OK Then
        Call Err.Raise(513, "GdiplusStartup�G���[")
    End If
End Sub

'*****************************************************************************
'[�T�v] �f�X�g���N�^
'*****************************************************************************
Private Sub Class_Terminate()
    If FGdiImage <> 0 Then Call GdipDisposeImage(FGdiImage)
    If FGdiPlus <> 0 Then Call GdiplusShutdown(FGdiPlus)
End Sub

'*****************************************************************************
'[�T�v] �摜�̕�
'[����] �Ȃ�
'*****************************************************************************
Public Property Get Width() As Long
    Dim lngWidth As Long
    Call GdipGetImageWidth(FGdiImage, lngWidth)
    Width = lngWidth
End Property

'*****************************************************************************
'[�T�v] �摜�̍���
'[����] �Ȃ�
'*****************************************************************************
Public Property Get Height() As Long
    Dim lngHeight As Long
    Call GdipGetImageHeight(FGdiImage, lngHeight)
    Height = lngHeight
End Property

'*****************************************************************************
'[�T�v] �t�@�C������GdiImage�I�u�W�F�N�g���쐬
'[����] �t�@�C�����A���E����
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub CreateFromFile(ByVal strFile As String, Optional lngWidth As Long = 0, Optional lngHeight As Long = 0)
    If GdipLoadImageFromFile(ByVal StrPtr(strFile), FGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipLoadImageFromFile�G���[")
    End If
    Call GdipBitmapSetResolution(FGdiImage, 96#, 96#) '�W����DPI�ɐݒ�
    
    If lngWidth <> 0 And lngHeight <> 0 Then
        If lngWidth <> Me.Width And lngHeight <> Me.Height Then
            Call Me.Resize(lngWidth, lngHeight)
        End If
    End If
End Sub

'*****************************************************************************
'[�T�v] Stream�I�u�W�F�N�g����GdiImage�I�u�W�F�N�g���쐬
'[����] Stream�I�u�W�F�N�g�A���E����
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub CreateFromStream(ByVal Stream As IUnknown, Optional lngWidth As Long = 0, Optional lngHeight As Long = 0)
    If GdipLoadImageFromStream(ObjPtr(Stream), FGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipLoadImageFromStream�G���[")
    End If
    Call GdipBitmapSetResolution(FGdiImage, 96#, 96#) '�W����DPI�ɐݒ�
    
    If lngWidth <> 0 And lngHeight <> 0 Then
        If lngWidth <> Me.Width And lngHeight <> Me.Height Then
            Call Me.Resize(lngWidth, lngHeight)
        End If
    End If
End Sub

'*****************************************************************************
'[�T�v] �r�b�g�}�b�v�n���h������GdiImage�I�u�W�F�N�g���쐬
'[����] �r�b�g�}�b�v�n���h���A���E����
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub CreateFromHBITMAP(ByVal hBitmap As LongPtr, Optional lngWidth As Long = 0, Optional lngHeight As Long = 0)
    Dim Pixels() As Long
    Call GetPixelsFromHBITMAP(hBitmap, Pixels())
    Call CreateFromPixels(Pixels())
    
    If lngWidth <> 0 And lngHeight <> 0 Then
        If lngWidth <> Me.Width And lngHeight <> Me.Height Then
            Call Me.Resize(lngWidth, lngHeight)
        End If
    End If

'    ���A���t�@�`���l�������f����Ȃ����ߋp��
'    If GdipCreateBitmapFromHBITMAP(hBitmap, 0&, FGdiImage) <> S_OK Then
'        Call Err.Raise(513, "GdipCreateBitmapFromHBITMAP�G���[")
'    End If
End Sub

'*****************************************************************************
'[�T�v] hBitmap����Pixels�ɃA���t�@�`���l������̐F��ݒ肷��
'[����] Bitmap�n���h���A�ߒl:Pixels
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub GetPixelsFromHBITMAP(ByVal hBitmap As LongPtr, ByRef Pixels() As Long)
On Error GoTo Finalization
    Dim hDc   As LongPtr
    Dim BmpInfHeader   As BITMAPINFOHEADER
    hDc = CreateCompatibleDC(0)
    BmpInfHeader.Size = 40 '40�Œ�

    '���ƍ����̂ݎ擾
    Dim lngHeight  As Long
    Call GetDIBits(hDc, hBitmap, 0, 0, ByVal 0, BmpInfHeader, DIB_RGB_COLORS)
    lngHeight = Abs(BmpInfHeader.Height)
    ReDim Pixels(1 To BmpInfHeader.Width, 1 To lngHeight)
    
    With BmpInfHeader
        .Height = -lngHeight  '�{�g���A�b�v����g�b�v�_�E���ɕύX
        .Planes = 1 '���1
        .BitCount = 32 'Bit
        .Compression = 0 'BI_RGB(�����k)
        .SizeImage = 0 '�ȗ���
    End With

    'Pixels()�ɒl��ݒ�
    Call GetDIBits(hDc, hBitmap, 0, lngHeight, Pixels(1, 1), BmpInfHeader, DIB_RGB_COLORS)
Finalization:
    If hDc <> 0 Then Call DeleteDC(hDc)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'*****************************************************************************
'[�T�v] Pixel��񂩂�GdiImage�I�u�W�F�N�g���쐬
'[����] Pixel���:�Q�����z��
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub CreateFromPixels(ByRef Pixels() As Long)
    Dim lngWidth  As Long
    Dim lngHeight As Long
    lngWidth = UBound(Pixels, 1)
    lngHeight = UBound(Pixels, 2)
    
    'GdipDisposeImage(FGdiImage)����܂�Pixels()�̗̈�����L���邽�ߋp��
    '��CreateFromHBITMAP������s����鎞�́AGdipDisposeImage�̑O��Pixels()����������
'    If GdipCreateBitmapFromScan0(lngWidth, lngHeight, lngWidth * 4, PixelFormat32bppARGB, Pixels(1, 1), FGdiImage) <> S_OK Then
    
    'GdiImage�I�u�W�F�N�g���쐬
    If GdipCreateBitmapFromScan0(lngWidth, lngHeight, 0, PixelFormat32bppARGB, ByVal 0, FGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0�G���[")
    End If
    Call GdipBitmapSetResolution(FGdiImage, 96#, 96#) '�W����DPI�ɐݒ�
    
    Dim BmpData As BITMAPDATA
    With BmpData
        .Width = lngWidth
        .Height = lngHeight
        .PixelFormat = PixelFormat32bppARGB
        .scan0 = VarPtr(Pixels(1, 1))
        .stride = lngWidth * 4
    End With
    If GdipBitmapLockBits(FGdiImage, ByVal 0, ImageLockModeUserInputBuf Or ImageLockModeWrite, PixelFormat32bppARGB, BmpData) <> S_OK Then
        Call Err.Raise(513, "GdipBitmapLockBits�G���[")
    End If
    Call GdipBitmapUnlockBits(FGdiImage, BmpData)

'    Dim x As Long, y As Long
'    For x = 1 To lngWidth
'        For y = 1 To lngHeight
'            Call GdipBitmapSetPixel(FGdiImage, x - 1, y - 1, Pixels(x, y))
'        Next
'    Next
End Sub

'*****************************************************************************
'[�T�v] GdiImage�I�u�W�F�N�g����Pixel����ݒ�
'[����] �ߒl�FPixel���
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ToPixels(ByRef Pixels() As Long)
    Dim lngWidth  As Long
    Dim lngHeight As Long
    lngWidth = Me.Width
    lngHeight = Me.Height
    ReDim Pixels(1 To lngWidth, 1 To lngHeight)
    
    Dim BmpData As BITMAPDATA
    With BmpData
        .Width = lngWidth
        .Height = lngHeight
        .PixelFormat = PixelFormat32bppARGB
        .scan0 = VarPtr(Pixels(1, 1))
        .stride = lngWidth * 4
    End With
    
    If GdipBitmapLockBits(FGdiImage, ByVal 0, ImageLockModeUserInputBuf Or ImageLockModeRead, PixelFormat32bppARGB, BmpData) <> S_OK Then
        Call Err.Raise(513, "GdipBitmapLockBits�G���[")
    End If
    Call GdipBitmapUnlockBits(FGdiImage, BmpData)
End Sub

'*****************************************************************************
'[�T�v] GdiImage����r�b�g�}�b�v�n���h�����쐬
'[����] �Ȃ�
'[�ߒl] �r�b�g�}�b�v�n���h��
'*****************************************************************************
Public Function ToHBITMAP() As LongPtr
    Dim hBitmap As LongPtr
    If GdipCreateHBITMAPFromBitmap(FGdiImage, hBitmap, 0) <> S_OK Then
        Call Err.Raise(513, "GdipCreateHBITMAPFromBitmap�G���[")
    End If
    ToHBITMAP = hBitmap
End Function

'*****************************************************************************
'[�T�v] GdiImage����PNG�̃t�@�C��Stream���쐬
'[����] �ߒl�FPNG�̃t�@�C��Stream
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ToPNGStream(ByRef Stream As IUnknown)
    Dim gGuid As Guid
    Call CLSIDFromString(StrPtr(CLSID_PNG), gGuid)
    If GdipSaveImageToStream(FGdiImage, ByVal ObjPtr(Stream), gGuid, 0) <> S_OK Then
        Call Err.Raise(513, "GdipSaveImageToStream�G���[")
    End If
End Sub

'*****************************************************************************
'[�T�v] PNG�t�@�C���ŕۑ�����
'[����] �t�@�C����
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub SaveToPNGFile(ByVal strFile As String)
    Dim gGuid As Guid
    Call CLSIDFromString(StrPtr(CLSID_PNG), gGuid)
    If GdipSaveImageToFile(FGdiImage, ByVal StrPtr(strFile), gGuid, 0) <> S_OK Then
        Call Err.Raise(513, "GdipSaveImageToFile�G���[")
    End If
End Sub

'*****************************************************************************
'[�T�v] GdiImage�I�u�W�F�N�g�̉摜�̃T�C�Y��ύX����
'[����] �V�����T�C�Y
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub Resize(ByVal lngWidth As Long, ByVal lngHeight As Long)
On Error GoTo Finalization
    '�V�����T�C�Y��Bitmap�I�u�W�F�N�g���쐬
    Dim hNewBmp  As LongPtr
    If GdipCreateBitmapFromScan0(lngWidth, lngHeight, 0, PixelFormat32bppARGB, ByVal 0&, hNewBmp) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0�G���[")
    End If
 
    '�T�C�Y��ύX
    Dim hGraphics As LongPtr
    If GdipGetImageGraphicsContext(hNewBmp, hGraphics) = 0 Then
'       Call GdipSetInterpolationMode(hGraphics, InterpolationModeHighQuality)
        Call GdipSetInterpolationMode(hGraphics, InterpolationModeNearestNeighbor)
        Call GdipSetPixelOffsetMode(hGraphics, PixelOffsetModeHalf)
        '�V�����T�C�Y�ŕ`��
        Call GdipDrawImageRectI(hGraphics, FGdiImage, 0, 0, lngWidth, lngHeight)
    End If
    
    '�ύX�O��GdiImage�I�u�W�F�N�g�����
    Call GdipDisposeImage(FGdiImage)
    '�ύX���GdiImage�I�u�W�F�N�g��������
    FGdiImage = hNewBmp
Finalization:
    If hGraphics <> 0 Then Call GdipDeleteGraphics(hGraphics)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

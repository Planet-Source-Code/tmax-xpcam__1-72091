VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImgPng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)
Private Const InterpolationModeHighQualityBicubic = 7
Private Const UnitPixel = 2

Private m_token As Long      'token to GDI engine
Private m_Image As Long         'original file loaded
Private m_Graphic As Long       'adjusted image
Private m_FileName As String

Private Property Get Initialised() As Boolean
    If m_token <> 0 Then
        Initialised = True
    End If
End Property

Private Function DisposeImage()
    GdipDisposeImage m_Image
    m_Image = 0
End Function

Private Sub Class_Initialize()
    'start the GDI engine
    On Error GoTo Handler
    Dim GpInput As GdiplusStartupInput
    GpInput.GdiplusVersion = 1
    GdiplusStartup m_token, GpInput, ByVal 0&
    Exit Sub
Handler:
   ' ReportError Err.Number, "Initialising GDI+", Err.Description
    Resume Next
End Sub

Private Sub Class_Terminate()
    If Initialised Then
        DisposeImage
        Call GdiplusShutdown(m_token)
    End If
End Sub

Public Function Load(sFileName As String) As Boolean
    Dim retval As Long
    DisposeImage
    m_FileName = sFileName
    retval = GdipLoadImageFromFile(StrPtr(sFileName), m_Image)
    If retval = 0 Then
        Load = True
    End If
End Function

Public Function StretchDC(ByVal hDC As Long, X As Integer, Y As Integer, W As Long, H As Long)
    Dim ImageWidth As Long
    Dim ImageHeight As Long
    On Error Resume Next
    'get a link to the DC we are going to paint onto
    GdipCreateFromHDC hDC, m_Graphic
    GdipGetImageWidth m_Image, ImageWidth
    GdipGetImageHeight m_Image, ImageHeight
    If W = 0 Then W = H * ImageWidth / ImageHeight
    If H = 0 Then H = W * ImageHeight / ImageWidth
    GdipSetInterpolationMode m_Graphic, InterpolationModeHighQualityBicubic 'InterpolationModeHighQualityBilinear
    GdipDrawImageRectRectI m_Graphic, m_Image, X, Y, W, H, 0, 0, ImageWidth, ImageHeight, UnitPixel, 0, 0, 0
    GdipDeleteGraphics m_Graphic
End Function

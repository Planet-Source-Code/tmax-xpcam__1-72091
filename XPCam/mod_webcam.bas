Attribute VB_Name = "Cam"
Option Explicit
Public mCapHwnd As Long
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WM_CAP_START = &H400
Public Const WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10
Public Const WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11
Public Const WM_CAP_EDIT_COPY = WM_CAP_START + 30
Public Const WM_CAP_SEQUENCE = WM_CAP_START + 62
Public Const WM_CAP_GRAB_FRAME = WM_CAP_START + 60
Public Const WM_CAP_GRAB_FRAME_NOSTOP = WM_CAP_START + 61
Public Const WM_CAP_FILE_SAVEAS = WM_CAP_START + 23
Public Const WM_CAP_FILE_SET_CAPTURE_FILE = WM_CAP_START + 20
Public Const WM_CAP_FILE_ALLOCATE = WM_CAP_START + 22
Public Const WM_CAP_SET_SCALE = WM_CAP_START + 53
Public Const WM_CAP_SET_PREVIEWRATE = WM_CAP_START + 52
Public Const WM_CAP_SET_PREVIEW = WM_CAP_START + 50
Public Const WM_CAP_STOP As Long = WM_CAP_START + 68
Public Const WM_CAP_FILE_SAVEDIB = WM_CAP_START + 25
Public Const WM_CAP_SINGLE_FRAME As Long = WM_CAP_START + 72
Public Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Public Const WM_CAP_DLG_VIDEOSOURCE As Long = WM_CAP_START + 42
Public Const WM_CAP_DLG_VIDEODISPLAY As Long = WM_CAP_START + 43
Public Const WM_CAP_GET_VIDEOFORMAT As Long = WM_CAP_START + 44
Public Const WM_CAP_SET_VIDEOFORMAT As Long = WM_CAP_START + 45
Public Const WM_CAP_DLG_VIDEOCOMPRESSION As Long = WM_CAP_START + 46
'--The capGetDriverDescription function retrieves the version
' description of the capture driver--
Public Declare Function capGetDriverDescription Lib "avicap32.dll" Alias "capGetDriverDescriptionA" (ByVal wDriverIndex As Long, ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, ByVal cbVer As Long) As Long
'--The capCreateCaptureWindow function creates a capture window--
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
'--This function sends the specified message to a window or windows--
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAsAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public VideoFile$

'---list all the various video sources---
Public Sub ListVideoSources(LstVideoSource As ListBox)
    Dim DriverName As String
    Dim DriverVersion As String
    Dim i As Integer
    DriverName = Space(80)
    DriverVersion = Space(80)
    For i = 0 To 9
        If capGetDriverDescription(i, DriverName, 80, DriverVersion, 80) Then
            LstVideoSource.AddItem StripNull(DriverName)
        End If
    Next
End Sub

Private Function StripNull(ByVal sValue As String) As String
    Dim lPos As Long
    lPos = InStr(sValue, Chr$(0))
    If lPos > 0 Then
        StripNull = Left$(sValue, lPos - 1)
    Else
        StripNull = sValue
    End If
End Function


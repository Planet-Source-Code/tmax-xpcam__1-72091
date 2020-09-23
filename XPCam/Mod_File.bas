Attribute VB_Name = "ModFile"
Option Explicit

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const COLORONCOLOR = 3
' ----==== API Declarations ====----
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public fsys As New FileSystemObject

Public Function GetFname(FileName As String) As String
Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(FileName)
    GetFname = f.Name
    Set f = Nothing
    Set fs = Nothing
End Function

Public Function GetFPath(FileName As String) As String
Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(FileName)
    GetFPath = f.ParentFolder + IIf(Right(f.ParentFolder, 1) <> "\", "\", "")
    Set f = Nothing
    Set fs = Nothing
End Function

Public Function GetExt(FileName As String) As String
Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    GetExt = fs.GetExtensionName(FileName)
    Set fs = Nothing
End Function

Function DragForm(hwnd As Long)
  Dim ret As Long
  ret = ReleaseCapture()
  ret = SendMessage(hwnd, WM_NCLBUTTONDOWN, 2&, 0&)
End Function

Sub ShowTopMost(hwnd As Long, topmost As Boolean)
If topmost Then
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End Sub



VERSION 5.00
Begin VB.UserControl MediaCtrl 
   BackColor       =   &H0000DAFF&
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   ScaleHeight     =   3345
   ScaleWidth      =   4620
   ToolboxBitmap   =   "MediaCtrl.ctx":0000
   Begin VB.Frame FraVideo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "MediaCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MediaCtrl
'Playing media file by mciSendString API
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'Default Property Values:
Const m_def_FileName = ""
Const m_def_AliasName = ""
Const WS_CHILD = &H40000000
'Property Variables:
Dim m_FileName As String
Dim m_AliasName As String
Dim ParentHwnd As Long

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    m_FileName = New_FileName
    PropertyChanged "FileName"
End Property

Public Property Get AliasName() As String
    AliasName = m_AliasName
End Property

Public Property Let AliasName(ByVal New_AliasName As String)
    m_AliasName = New_AliasName
    PropertyChanged "AliasName"
End Property

Private Sub UserControl_InitProperties()
    m_FileName = m_def_FileName
    m_AliasName = m_def_AliasName
    ParentHwnd = FraVideo.hWnd
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
    m_AliasName = PropBag.ReadProperty("AliasName", m_def_AliasName)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
FraVideo.Width = UserControl.Width - FraVideo.Left * 2
FraVideo.Height = UserControl.Height - FraVideo.Top * 2
End Sub

Private Sub UserControl_Terminate()
    mmClose
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
    Call PropBag.WriteProperty("AliasName", m_AliasName, m_def_AliasName)
End Sub

Public Function IsPlaying() As Boolean
    Static s As String * 30
    mciSendString "status " & AliasName & " mode", s, Len(s), 0
    IsPlaying = (Mid$(s, 1, 7) = "playing")
End Function

Public Function mmPlay()
    Dim cmdToDo As String * 255
    Dim dwReturn As Long
    Dim ret As String * 128
    If Dir(FileName) = "" Then
        mmOpen = "Error with input file"
        Exit Function
    End If
    cmdToDo = "open " & FileName & " type MPEGVideo Alias " & AliasName & " Parent " & FraVideo.hWnd & " Style " & WS_CHILD
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    dwReturn = mciSendString("put " & AliasName & " window at 0 0 " & FraVideo.Width \ Screen.TwipsPerPixelX & " " & FraVideo.Height \ Screen.TwipsPerPixelY, 0&, 0&, 0&)
    If dwReturn <> 0 Then  'not success
        mciGetErrorString dwReturn, ret, 128
        mmOpen = ret
        MsgBox cmdToDo & vbCrLf & ret, vbCritical
        Exit Function
    End If
    mmPlay = "Success"
    mciSendString "play " & AliasName, 0, 0, 0
End Function

Public Function mmPause()
    mciSendString "pause " & AliasName, 0, 0, 0
End Function

Public Function mmStop() As String
    mciSendString "stop " & AliasName, 0, 0, 0
    mciSendString "close " & AliasName, 0, 0, 0
End Function

Public Function mmClose() As String
    mciSendString "Close All", 0&, 0&, 0&
End Function

Public Function PositionInSec()
    Static s As String * 30
    mciSendString "set " & AliasName & " time format milliseconds", 0, 0, 0
    mciSendString "status " & AliasName & " position", s, Len(s), 0
    PositionInSec = Round(Mid$(s, 1, Len(s)) / 1000)
End Function

Public Function Position()
    Static s As String * 30
    mciSendString "set " & AliasName & " time format milliseconds", 0, 0, 0
    mciSendString "status " & AliasName & " position", s, Len(s), 0
    sec = Round(Mid$(s, 1, Len(s)) / 1000)
    If sec < 60 Then Position = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        Position = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function

Public Function LengthInSec()
    Static s As String * 30
    mciSendString "set " & AliasName & " time format milliseconds", 0, 0, 0
    mciSendString "status " & AliasName & " length", s, Len(s), 0
    LengthInSec = Round(Val(Mid$(s, 1, Len(s))) / 1000)
End Function

Public Function Length()
    Static s As String * 30
    mciSendString "set " & AliasName & " time format milliseconds", 0, 0, 0
    mciSendString "status " & AliasName & " length", s, Len(s), 0
    sec = Round(Val(Mid$(s, 1, Len(s))) / 1000)
    If sec < 60 Then Length = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        Length = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function


Public Function SeekTo(Second)
    mciSendString "set " & AliasName & " time format milliseconds", 0, 0, 0
    If IsPlaying = True Then mciSendString "play " & AliasName & " from " & Second, 0, 0, 0
    If IsPlaying = False Then mciSendString "seek " & AliasName & " to " & Second, 0, 0, 0
End Function


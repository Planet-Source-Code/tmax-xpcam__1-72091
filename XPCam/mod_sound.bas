Attribute VB_Name = "Mod_Sound"
Option Explicit
Const SND_ASYNC& = &H1
Const SND_NODEFAULT& = &H2
Const SND_NOWAIT& = &H2000
Const SND_SYNC& = &H0
Const SND_NOSTOP = &H10
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub PlaySnd(fname As String)
Dim ret As Long
    ret = sndPlaySound(fname, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP Or SND_NOWAIT)
End Sub


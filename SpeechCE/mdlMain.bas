Attribute VB_Name = "mdlMain"
Public Seconds As Integer
Public Minutes As Integer
Public Limit1 As Integer
Public Limit2 As Integer
Public Limit3 As Integer

Public Temp As String
Public Declare Function sndPlaySound Lib "coredll.dll" Alias "sndPlaySoundW" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long



Public Sub Center(frm As Form)
Dim X As Integer
Dim Y As Integer

X = (Screen.Width - frm.Width) / 2
Y = (Screen.Height - frm.Height) / 2

frm.Left = X
frm.Top = 0 'Y
End Sub

Sub PlayWaveFile(WavFile)
    Dim ret
    Dim flags
    flags = SND_ASYNC Or SND_NODEFAULT
    ret = sndPlaySound(WavFile, flags)
End Sub

Attribute VB_Name = "Module1"
Option Explicit
'Public NumPlayers
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Public Highscore As Integer
Public ClicksCount
Public NewGame
Public Color
Public Game

Public tile(7)

Public Declare Function sndPlaySound Lib "coredll.dll" Alias "sndPlaySoundW" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Sub cblev1(ByRef Command1 As ButtonConstants, ByRef Command2 As ButtonConstants)

If Command2.Caption = "" Then
    Command2.Caption = Command1.Caption
    Command1.Caption = ""
    ClicksCount = ClicksCount + 1
    frmMain.Clicks.Caption = ClicksCount
    If frmMain.Label4.Caption = "1" Then Command2.BackColor = &HC0C0C0
    If frmMain.Label4.Caption = "2" Then Command2.BackColor = vbYellow
    If frmMain.Label4.Caption = "3" Then Command2.BackColor = vbBlue
    If frmMain.Label4.Caption = "4" Then Command2.BackColor = vbRed
    Command1.BackColor = &HFFFFFF
End If
 PlayWaveFile (App.Path & "\select.wav")
 
 
 
If frmMain.Label2.Caption = "1" Then
    If frmMain.Label3.Caption = "0" Then
        If frmMain.Command1.Caption = "1" And frmMain.Command2.Caption = "2" And frmMain.Command3.Caption = "3" And frmMain.Command4.Caption = "4" And frmMain.Command5.Caption = "5" And frmMain.Command6.Caption = "6" And frmMain.Command7.Caption = "7" And frmMain.Command8.Caption = "8" And frmMain.Command9.Caption = "" Then
           PlayWaveFile (App.Path & "\Cluck1.wav")
           MsgBox ("Well done! You solved the puzzle with " & ClicksCount & " moves!"), vbOKCancel, "CONGRATS"
         frmMain.Label2.Caption = "0"
         End If
        
    End If

    If frmMain.Label3.Caption = "1" Then
        If frmMain.Command1.Caption = "95LX" And frmMain.Command2.Caption = "100LX" And frmMain.Command3.Caption = "200LX" And frmMain.Command4.Caption = "300LX" And frmMain.Command5.Caption = "320LX" And frmMain.Command6.Caption = "J.680" And frmMain.Command7.Caption = "J.690" And frmMain.Command8.Caption = "J.720" And frmMain.Command9.Caption = "" Then
           PlayWaveFile (App.Path & "\Cluck1.wav")
           MsgBox ("Well done! You solved the puzzle with " & ClicksCount & " moves!"), vbOKCancel, "CONGRATS"
         frmMain.Label2.Caption = "0"
         End If
         
    End If
End If


     
   




End Sub
Sub cblev2(ByRef Command1 As ButtonConstants, ByRef Command2 As ButtonConstants)

If Command2.Caption = "" Then
    Command2.Caption = Command1.Caption
    Command1.Caption = ""
    ClicksCount = ClicksCount + 1
    frmfour.Clicks.Caption = ClicksCount
    'If frmMain.Label4.Caption = "1" Then Command2.BackColor = &HC0C0C0
    'If frmMain.Label4.Caption = "2" Then Command2.BackColor = vbYellow
    'If frmMain.Label4.Caption = "3" Then Command2.BackColor = vbBlue
    'If frmMain.Label4.Caption = "4" Then Command2.BackColor = vbRed
    Command1.BackColor = &HFFFFFF
    Command2.BackColor = &HC0C0C0
End If
 PlayWaveFile (App.Path & "\select.wav")
 
 
 

If frmfour.Command1.Caption = "1" And frmfour.Command2.Caption = "2" And frmfour.Command3.Caption = "3" And frmfour.Command4.Caption = "4" And frmfour.Command5.Caption = "5" And frmfour.Command6.Caption = "6" And frmfour.Command7.Caption = "7" And frmfour.Command8.Caption = "8" And frmfour.Command9.Caption = "9" And frmfour.Command10.Caption = "10" And frmfour.Command11.Caption = "11" And frmfour.Command12.Caption = "12" And frmfour.Command13.Caption = "13" And frmfour.Command14.Caption = "14" And frmfour.Command15.Caption = "15" And frmfour.Command16.Caption = "" Then
    PlayWaveFile (App.Path & "\Cluck1.wav")
    MsgBox ("Well done! You solved the puzzle with " & ClicksCount & " moves!"), vbOKCancel, "CONGRATS"
    'frmMain.Label2.Caption = "0"
End If




     
   




End Sub
Sub PlayWaveFile(WavFile)
    Dim ret
    Dim flags
    flags = SND_ASYNC Or SND_NODEFAULT
    ret = sndPlaySound(WavFile, flags)
End Sub

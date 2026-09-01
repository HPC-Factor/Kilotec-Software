Attribute VB_Name = "modDice"
'********************************************************************************
'*
'*  Dice - It emulates (to borrow a term) the rules of Yahtzee
'*
'*  Copyright (c) 1998 Microsoft Corporation

Dim GetValue

'Public NumPlayers
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Public Highscore As Integer

Const ID_NEW = 101
Const ID_ABOUT = 103
Const ID_EXIT = 105

Const ID_OPTIONS = 107
Const ID_HS = 109
Const ID_SOUND = 111

Public Player As String
Public Sound As Boolean
Public DiceRolls As Integer
Public Color As Integer



Public DBdaten(2, 5)
Public HS(5)
Public Aktuell(2)

Public IndexX As Integer
Public Zaehler As Integer

Public Temp As String
Public XX As Integer
Public YY As Integer
Public T As Integer
Public FF As Integer

Public TempDate As Date

Public SaveFlag As Integer

Public Declare Function sndPlaySound Lib "coredll.dll" Alias "sndPlaySoundW" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Function ChooseBitmap()
    Dim Die
    Randomize
    Die = Int((6) * Rnd + 1)
    GetValue = Die
    Select Case Die
        Case 1
            If Color = 0 Then ChooseBitmap = App.Path & "\1_Dice.bmp"
            If Color = 1 Then ChooseBitmap = App.Path & "\1_DiceB.bmp"
        Case 2
            If Color = 0 Then ChooseBitmap = App.Path & "\2_Dice.bmp"
            If Color = 1 Then ChooseBitmap = App.Path & "\2_DiceB.bmp"
        Case 3
            If Color = 0 Then ChooseBitmap = App.Path & "\3_Dice.bmp"
            If Color = 1 Then ChooseBitmap = App.Path & "\3_DiceB.bmp"
        Case 4
            If Color = 0 Then ChooseBitmap = App.Path & "\4_Dice.bmp"
            If Color = 1 Then ChooseBitmap = App.Path & "\4_DiceB.bmp"
        Case 5
            If Color = 0 Then ChooseBitmap = App.Path & "\5_Dice.bmp"
            If Color = 1 Then ChooseBitmap = App.Path & "\5_DiceB.bmp"
        Case 6
            If Color = 0 Then ChooseBitmap = App.Path & "\6_Dice.bmp"
            If Color = 1 Then ChooseBitmap = App.Path & "\6_DiceB.bmp"
    End Select
End Function
Sub PlayWaveFile(WavFile)
    Dim ret
    Dim flags
    flags = SND_ASYNC Or SND_NODEFAULT
    If Sound = True Then ret = sndPlaySound(WavFile, flags)
End Sub


Public Sub Center(frm As Form)
Dim X As Integer
Dim Y As Integer

X = (Screen.Width - frm.Width) / 2
Y = (Screen.Height - frm.Height) / 2

frm.Left = X
frm.Top = 0 'Y
End Sub
    
Public Sub DBladen()

T = 0
frmMain.File.Open App.Path & "\YHS.txt", fsModeAppend
frmMain.File.Close

frmMain.File.Open App.Path & "\YHS.txt", fsModeInput, fsAccessRead, fsLockRead

Do While Not frmMain.File.EOF
   DBdaten(0, T) = frmMain.File.LineInputString
   DBdaten(1, T) = frmMain.File.LineInputString
   DBdaten(2, T) = frmMain.File.LineInputString
      
    'Zaehler = Zaehler + 1
    T = T + 1
  Loop

Zaehler = T
frmMain.File.Close

Exit Sub
End Sub

Public Sub EinlesenEintraege()
' Lese in Listview ein;
' Möglich wäre auch eine Listbox
Dim j As Integer

HS(0) = CDbl(DBdaten(2, 0))
HS(1) = CDbl(DBdaten(2, 1))
HS(2) = CDbl(DBdaten(2, 2))
HS(3) = CDbl(DBdaten(2, 3))
HS(4) = CDbl(DBdaten(2, 4))
HS(5) = CDbl(DBdaten(2, 5))





For XX = 0 To 5
    For j = XX + 1 To 5
        If HS(XX) > HS(j) Then
            HS(XX) = HS(XX) + HS(j)
            HS(j) = HS(XX) - HS(j)
            HS(XX) = HS(XX) - HS(j)
        End If
    Next j



Next XX

frmHS.Score1.Caption = HS(5)
frmHS.Score2.Caption = HS(4)
frmHS.Score3.Caption = HS(3)
frmHS.Score4.Caption = HS(2)
frmHS.Score5.Caption = HS(1)



For XX = 0 To 5
    If HS(5) = CDbl(DBdaten(2, XX)) Then frmHS.Name1.Caption = DBdaten(0, XX)
    If HS(4) = CDbl(DBdaten(2, XX)) Then frmHS.Name2.Caption = DBdaten(0, XX)
    If HS(3) = CDbl(DBdaten(2, XX)) Then frmHS.Name3.Caption = DBdaten(0, XX)
    If HS(2) = CDbl(DBdaten(2, XX)) Then frmHS.Name4.Caption = DBdaten(0, XX)
    If HS(1) = CDbl(DBdaten(2, XX)) Then frmHS.Name5.Caption = DBdaten(0, XX)
    
    If HS(5) = CDbl(DBdaten(2, XX)) Then frmHS.Datum1.Caption = DBdaten(1, XX)
    If HS(4) = CDbl(DBdaten(2, XX)) Then frmHS.Datum2.Caption = DBdaten(1, XX)
    If HS(3) = CDbl(DBdaten(2, XX)) Then frmHS.Datum3.Caption = DBdaten(1, XX)
    If HS(2) = CDbl(DBdaten(2, XX)) Then frmHS.Datum4.Caption = DBdaten(1, XX)
    If HS(1) = CDbl(DBdaten(2, XX)) Then frmHS.Datum5.Caption = DBdaten(1, XX)
    
Next XX


End Sub

Public Sub AbInDieDatenbank()
DBdaten(0, 0) = "H/PC"
DBdaten(1, 0) = Date
DBdaten(2, 0) = 80
DBdaten(0, 1) = "H/PC"
DBdaten(1, 1) = Date
DBdaten(2, 1) = 90
DBdaten(0, 2) = "H/PC"
DBdaten(1, 2) = Date
DBdaten(2, 2) = 100
DBdaten(0, 3) = "H/PC"
DBdaten(1, 3) = Date
DBdaten(2, 3) = 150
DBdaten(0, 4) = "H/PC"
DBdaten(1, 4) = Date
DBdaten(2, 4) = 200
DBdaten(0, 5) = "H/PC"
DBdaten(1, 5) = Date
DBdaten(2, 5) = 0

EinlesenEintraege
DBspeichern

End Sub

Public Sub DBspeichern()

frmMain.File.Open App.Path & "\YHS.txt", fsModeOutput, fsAccessWrite, fsLockReadWrite

 For XX = 0 To 4
'If DBdaten(1, XX) <> "" Then
  frmMain.File.LinePrint DBdaten(0, XX)
  frmMain.File.LinePrint DBdaten(1, XX)
  frmMain.File.LinePrint DBdaten(2, XX)
 'End If
Next XX
  frmMain.File.Close
End Sub

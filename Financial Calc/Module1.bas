Attribute VB_Name = "Module1"
Option Explicit

Public SPIELER As String
Public TEAM As String

Public Mannschaft(17) As String
Public Punkte(17)
Public Tore(17)
Public Gegentore(17)
Public Form(17) As Long
Public Spiel(17) As Double
Public Spieltag As Integer
Public TabStop As String
Public xx As Long
Public Geld As Long
Public Ausgaben As Long
Public Eintritt As Integer
Public Zuschauer As Integer
Public Zuschauerschnitt As Integer
Public Stadion(17) As Integer
Public TeamIndex As Integer
Public s(10) As Double
Public ab(10) As Double

Public NurPaarungen As Boolean


Public Sub Center(frm As Form)
Dim x As Integer
Dim y As Integer

x = (Screen.Width - frm.Width) / 2
y = (Screen.Height - frm.Height) / 2

frm.Left = x
frm.Top = y
End Sub

Public Sub Paarungen(x As Boolean)
Dim Tore1 As Integer
Dim Tore2 As Integer

Dim gt1 As Integer
Dim gt2 As Integer

gt1 = 0
gt2 = 0


xx = Spieltag 'xx + 1
Dim y As Long
Dim spiele As Long
Dim spiele2 As Long

'For spiele = Spieltag To Spieltag
    y = xx + 1
    If y > 17 Then y = 1
    
        If Form(xx) > Form(y) Then gt1 = 2
        If Form(y) > Form(xx) Then gt2 = 2
        If Spiel(xx) > Spiel(y) Then gt1 = 3
        If Spiel(y) > Spiel(xx) Then gt2 = 3
    
    For spiele2 = 1 To 9
        
  
        If NurPaarungen = False Then
            Tore1 = Int(Rnd(1) * (gt1 + Spiel(xx))) '* (Form(xx) / 10))
            Tore2 = Int(Rnd(1) * (gt2 + Spiel(y))) '* (Form(y) / 10))
        End If
        
        frmErgebnis.List1.AddItem Mannschaft(xx) & Space(20 - Len(Mannschaft(xx))) & " - " & Space(4 - Len(" - ")) & Mannschaft(y) & Space(20 - Len(Mannschaft(y))) & Tore1 & ":" & Tore2
        frmSpieltag.List1.AddItem Mannschaft(xx) & Space(20 - Len(Mannschaft(xx))) & " - " & Space(4 - Len(" - ")) & Mannschaft(y)
        
        
      'frmErgebnis.Label1.Caption = frmErgebnis.Label1.Caption + ", " & y
       'frmErgebnis.Label2.Caption = frmErgebnis.Label2.Caption + ", " & xx
       
       If NurPaarungen = False Then
        Tore(xx) = Tore(xx) + Tore1
        Gegentore(xx) = Gegentore(xx) + Tore2
        Tore(y) = Tore(y) + Tore2
        Gegentore(y) = Gegentore(y) + Tore1
        
        If Tore1 > Tore2 Then
            Punkte(xx) = Punkte(xx) + 3
        End If
        If Tore2 > Tore1 Then
            Punkte(y) = Punkte(y) + 3
        End If
        If Tore1 = Tore2 Then
           Punkte(xx) = Punkte(xx) + 1
            Punkte(y) = Punkte(y) + 1
        End If
        End If
        
        xx = xx - 1
        y = y + 1

        If xx < 0 Then xx = 17
        If y > 17 Then y = 1
        
        
    
    Next spiele2
 
'Next spiele
If x = False Then xx = xx - 1
End Sub

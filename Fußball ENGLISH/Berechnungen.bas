Attribute VB_Name = "Berechnungen"
Option Explicit
 Public Function GetRandomNumber(min As Long, max As Long) As Long
    Randomize
    GetRandomNumber = Int(Rnd * (max - min + 1) + min)
End Function

Public Sub TabellenRechner()
  
    For i = 1 To 18
        Mannschaft2(i) = Mannschaft(i)
        Tore2(i) = Tore(i)
        Gegentore2(i) = Gegentore(i)
        Punkte2(i) = Punkte(i)
    Next i
    
    Dim j As Long
    
    For i = 1 To 18
        For j = i + 1 To 18
            If Punkte2(i) < Punkte2(j) Then
                Punkte2(i) = Punkte2(i) + Punkte2(j)
                Punkte2(j) = Punkte2(i) - Punkte2(j)
                Punkte2(i) = Punkte2(i) - Punkte2(j)
                
                
                temp = Mannschaft2(i)
                Mannschaft2(i) = Mannschaft2(j)
                Mannschaft2(j) = temp
                           
                temp2 = Tore2(i)
                Tore2(i) = Tore2(j)
                Tore2(j) = temp2
                
                temp3 = Gegentore2(i)
                Gegentore2(i) = Gegentore2(j)
                Gegentore2(j) = temp3
            End If
        Next j
Next i

End Sub
 
Public Sub Aktualisieren()
frmHaupt.lblSpieler.Caption = Spieler
frmHaupt.lblTeam.Caption = TEAM


Form(TeamIndex) = (Angriff(TeamIndex) + Mittelfeld(TeamIndex) + Abwehr(TeamIndex)) / 3

TabellenRechner
For i = 1 To 18
    If TEAM = Mannschaft2(i) Then Platz = i
    Form(i) = Int((Angriff(i) + Mittelfeld(i) + Abwehr(i)) / 3)
Next i



frmHaupt.lblGeld.Caption = "Kontostand: " & Int(Geld) & " €"
frmHaupt.lblPlatz.Caption = "Platz: " & Platz & " / " & "Stärke: " & Form(Int(TeamIndex))
frmHaupt.lblSpieltag.Caption = Spieltag & ". Spieltag"
frmHaupt.lblLiga.Caption = Liga & ". Bundesliga"
frmHaupt.lblSaison.Caption = "Saison: " & Saison


If Spieltag = 35 Then
    frmHaupt.Command1.Caption = "SAISONENDE"
    frmHaupt.lblSpieltag.Caption = "SAISONENDE"
    Exit Sub
End If

frmHaupt.lblGegner.Caption = ""
tempx = x

   frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & vbCrLf
    y = x + 1
    If y > (18 - 1) Then y = 1
    For spiele = 1 To 18 / 2 - 1

        If TEAM = Mannschaft(x) Then
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & Mannschaft(TeamIndex)
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & vbCrLf
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & Angriff(TeamIndex) & "-" & Mittelfeld(TeamIndex) & "-" & Abwehr(TeamIndex) & "-(" & Form(TeamIndex) & ")"
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & vbCrLf & vbCrLf
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & "gegen"
            Heimspiel = True
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & vbCrLf & vbCrLf
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & Mannschaft(y)
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & vbCrLf
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & Angriff(y) & "-" & Mittelfeld(y) & "-" & Abwehr(y) & "-(" & Form(y) & ")"
        End If
      
        
        If TEAM = Mannschaft(y) Then
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & Mannschaft(x)
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & vbCrLf
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & Angriff(x) & "-" & Mittelfeld(x) & "-" & Abwehr(x) & "-(" & Form(x) & ")"
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & vbCrLf & vbCrLf
            frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & "gegen"
            Heimspiel = False
        End If
        x = x - 1
        y = y + 1
        If x < 1 Then x = 18 - 1
        If y > (18 - 1) Then y = 1
    Next spiele
            
            If Heimspiel = False Then
                frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & vbCrLf & vbCrLf
                frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & Mannschaft(TeamIndex)
                frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & vbCrLf
                frmHaupt.lblGegner.Caption = frmHaupt.lblGegner.Caption & Angriff(TeamIndex) & "-" & Mittelfeld(TeamIndex) & "-" & Abwehr(TeamIndex) & "-(" & Form(TeamIndex) & ")"
                
            End If


x = tempx

End Sub

Public Sub Berechnen()
frmSpieltag.List1.Clear

tempx = x
   y = x + 1
    If y > (18 - 1) Then y = 1
    For spiele = 1 To 18 / 2 - 1
        frmSpieltag.List1.AddItem Mannschaft(x) & Space(20 - Len(Mannschaft(x))) & " - " & Space(4 - Len(" - ")) & Mannschaft(y)
 
        x = x - 1
        y = y + 1
        If x < 1 Then x = 18 - 1
        If y > (18 - 1) Then y = 1
    Next spiele
     
        frmSpieltag.List1.AddItem Mannschaft(18) & Space(20 - Len(Mannschaft(18))) & " - " & Space(4 - Len(" - ")) & Mannschaft(x)
 
x = tempx

End Sub

Public Sub Center(frm As Form)
Dim x As Integer
Dim y As Integer

x = (Screen.Width - frm.Width) / 2
y = (Screen.Height - frm.Height) / 2

frm.Left = x
frm.Top = y
End Sub

Public Sub DBspeichern()
Dim i As Integer

frmHaupt.File.Open App.Path & "\FM.txt", fsModeOutput, fsAccessWrite, fsLockReadWrite

    frmHaupt.File.LinePrint Platz
    frmHaupt.File.LinePrint Liga
    frmHaupt.File.LinePrint Spieltag
    frmHaupt.File.LinePrint TEAM
    frmHaupt.File.LinePrint TeamIndex
    frmHaupt.File.LinePrint x
    frmHaupt.File.LinePrint y
    frmHaupt.File.LinePrint Spieler
    frmHaupt.File.LinePrint x2
    frmHaupt.File.LinePrint Geld

For i = 1 To 18
    frmHaupt.File.LinePrint Mannschaft(i)
    frmHaupt.File.LinePrint Stadion(i)
    frmHaupt.File.LinePrint Angriff(i)
    frmHaupt.File.LinePrint Mittelfeld(i)
    frmHaupt.File.LinePrint Abwehr(i)
Next i

  
For i = 1 To 10
    frmHaupt.File.LinePrint ab(i)
Next i

For i = 1 To 18
    frmHaupt.File.LinePrint Tore(i)
    frmHaupt.File.LinePrint Gegentore(i)
    frmHaupt.File.LinePrint Punkte(i)
    frmHaupt.File.LinePrint Form(i)
Next i
    

  
  frmHaupt.File.Close
End Sub

Public Sub DBladen()

Dim i As Integer
frmHaupt.File.Open App.Path & "\FM.txt", fsModeInput
frmHaupt.File.Close

frmHaupt.File.Open App.Path & "\FM.txt", fsModeInput, fsAccessRead, fsLockRead

Do While Not frmHaupt.File.EOF
    Platz = frmHaupt.File.LineInputString
    Liga = frmHaupt.File.LineInputString
    Spieltag = frmHaupt.File.LineInputString
    TEAM = frmHaupt.File.LineInputString
    TeamIndex = frmHaupt.File.LineInputString
    x = frmHaupt.File.LineInputString
    y = frmHaupt.File.LineInputString
    Spieler = frmHaupt.File.LineInputString
    x2 = frmHaupt.File.LineInputString
    Geld = frmHaupt.File.LineInputString
    
For i = 1 To 18
    Mannschaft(i) = frmHaupt.File.LineInputString
    Stadion(i) = frmHaupt.File.LineInputString
    Angriff(i) = frmHaupt.File.LineInputString
    Mittelfeld(i) = frmHaupt.File.LineInputString
    Abwehr(i) = frmHaupt.File.LineInputString
Next i

For i = 1 To 10
   ab(i) = frmHaupt.File.LineInputString
Next i

For i = 1 To 18
    Tore(i) = frmHaupt.File.LineInputString
    Gegentore(i) = frmHaupt.File.LineInputString
    Punkte(i) = frmHaupt.File.LineInputString
    Form(i) = frmHaupt.File.LineInputString
Next i
  
  
  
  Loop
 
  

frmHaupt.File.Close

Exit Sub
End Sub


Public Sub Liga2Aktualisieren()
    TabellenRechner
    
    Mannschaft(1) = "Fortuna Düsseldorf"
    Mannschaft(2) = "Karlsruher SC"
    Mannschaft(3) = "Hamburger SV"
    Mannschaft(4) = "FC Schalke 04"
    Mannschaft(5) = "1. FC Magdeburg"
    Mannschaft(6) = "SC Paderborn"
    Mannschaft(7) = "Hannover 96"
    Mannschaft(8) = "Hertha BSC"
    Mannschaft(9) = "SV Elversberg"
    Mannschaft(10) = "1. FC Nürnberg"
    Mannschaft(11) = "1. FC Köln"
    Mannschaft(12) = "FC Kaiserslautern"
    Mannschaft(13) = "Greuther Fürth"
    Mannschaft(14) = "Darmstadt 98"
    Mannschaft(15) = "SSV Ulm 1846"
    Mannschaft(16) = "TSV Braunschweig"
    Mannschaft(17) = Mannschaft2(2)
    Mannschaft(18) = Mannschaft2(1)

    Stadion(1) = 54600
    Stadion(2) = 34302
    Stadion(3) = 57000
    Stadion(4) = 62271
    Stadion(5) = 30098
    Stadion(6) = 15000
    Stadion(7) = 49000
    Stadion(8) = 74667
    Stadion(9) = 10000
    Stadion(10) = 50000
    Stadion(11) = 50000
    Stadion(12) = 49350
    Stadion(13) = 16626
    Stadion(14) = 17810
    Stadion(15) = 17000
    Stadion(16) = 23325
    Stadion(17) = 12794
    Stadion(18) = 15210
    
    ab(1) = 152500
    ab(2) = 300000
    ab(3) = 442500
    ab(4) = 493000
    ab(5) = 570000
    ab(6) = 667500
    ab(7) = 795000
    ab(8) = 910000
    ab(9) = 1020000
    ab(10) = 1125000

    For i = 1 To 18
        Angriff(i) = GetRandomNumber(10, 80)
        Mittelfeld(i) = GetRandomNumber(10, 80)
        Abwehr(i) = GetRandomNumber(10, 80)
        Form(i) = Int((Angriff(i) + Mittelfeld(i) + Abwehr(i)) / 3)
    Next i

    Eintritt = 15
    Liga = 2
End Sub


Public Sub Liga1Aktualisieren()
    TabellenRechner
    
    Mannschaft(1) = "Bayern München"
    Mannschaft(2) = "Bayer Leverkusen"
    Mannschaft(3) = "Borussia Dortmund"
    Mannschaft(4) = "VFB Stuttgart"
    Mannschaft(5) = "FC Heidenheim"
    Mannschaft(6) = "FC Augsburg"
    Mannschaft(7) = "VFL Wolfsburg"
    Mannschaft(8) = "TSG Hoffenheim"
    Mannschaft(9) = "Werder Bremen"
    Mannschaft(10) = "RB Leipzig"
    Mannschaft(11) = "Mönchengladbach"
    Mannschaft(12) = "Eintracht Frankfurt"
    Mannschaft(13) = "Union Berlin"
    Mannschaft(14) = "SC Freiburg"
    Mannschaft(15) = "VFL Bochum"
    Mannschaft(16) = "FSV Mainz"
    Mannschaft(17) = Mannschaft2(2)
    Mannschaft(18) = Mannschaft2(1)

Stadion(1) = 75000
    Stadion(2) = 30210
    Stadion(3) = 81365
    Stadion(4) = 56589
    Stadion(5) = 44000
    Stadion(6) = 50000
    Stadion(7) = 28917
    Stadion(8) = 30150
    Stadion(9) = 42100
    Stadion(10) = 47069
    Stadion(11) = 54014
    Stadion(12) = 58000
    Stadion(13) = 22012
    Stadion(14) = 34700
    Stadion(15) = 26000
    Stadion(16) = 33305
    Stadion(17) = 45000
    Stadion(18) = 32350
    
    ab(1) = 252500
    ab(2) = 400000
    ab(3) = 542500
    ab(4) = 593000
    ab(5) = 670000
    ab(6) = 767500
    ab(7) = 895000
    ab(8) = 1010000
    ab(9) = 1120000
    ab(10) = 1225000

    For i = 1 To 18
        Angriff(i) = GetRandomNumber(20, 100)
        Mittelfeld(i) = GetRandomNumber(20, 100)
        Abwehr(i) = GetRandomNumber(20, 100)
        Form(i) = Int((Angriff(i) + Mittelfeld(i) + Abwehr(i)) / 3)
    Next i

    Eintritt = 25
    Liga = 1
End Sub


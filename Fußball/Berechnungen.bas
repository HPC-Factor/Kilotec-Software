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
Dim Heimspiel As Boolean

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
Next i
  
  
  
  Loop
 
  

frmHaupt.File.Close

Exit Sub
End Sub

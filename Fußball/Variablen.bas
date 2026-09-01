Attribute VB_Name = "Variablen"
Option Explicit

Public Geld As Long
Public Eintritt As Integer
Public TeamIndex As Integer
Public TEAM As String
Public Spieler As String
Public Spieltag As Integer
Public Spieltag2 As Integer
Public NurPaarungen As Boolean
Public Zuschauerschnitt As Long
Public Ausgaben As Long

Public Mannschaften As Integer
Public x, y As Integer
Public spiele As Integer

Public Liga As Integer

Public xx As Integer
Public yy As Integer


'MANNSCHAFT
Public Mannschaft(18)
Public Stadion(18)
Public Tore(19)
Public Gegentore(19)
Public Punkte(19)
Public Aufstellung(13)
Public Angriff(19)
Public Mittelfeld(19)
Public Abwehr(19)
Public Form(19)

Public ab(19)
Public s(19)
Public Position As String

'TABELLENRECHNER
Public VorNull As String
Public i As Long
Public temp As String
Public temp2 As Integer
Public temp3 As Integer
Public temp4 As Integer
Public Mannschaft2(19)
Public Tore2(19)
Public Gegentore2(19)
Public Punkte2(19)


Public Platz As Integer
Public tempx As Integer

'ERGEBNIS
Public Angriffsfaktor(19) As Double
Public Mittelfeldfaktor(19) As Double
Public Abwehrfaktor(19) As Double
Public x2 As Integer


Public Zuschauer As Long
Public Kosten As Long

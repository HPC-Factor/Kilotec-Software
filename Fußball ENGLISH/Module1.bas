Attribute VB_Name = "Module1"
Option Explicit


Public Geld As Long
Public Eintritt As Integer
Public TeamIndex As Integer
Public TEAM As String
Public Spieler As String
Public Spieltag As Integer
Public xx As Integer
Public NurPaarungen As Boolean
Public Zuschauerschnitt As Long
Public Ausgaben As Long

Public Mannschaften As Integer
Public x, y As Integer
Public spiele As Integer




Public Sub Center(frm As Form)
Dim x As Integer
Dim y As Integer

x = (Screen.Width - frm.Width) / 2
y = (Screen.Height - frm.Height) / 2

frm.Left = x
frm.Top = y
End Sub

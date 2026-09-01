Attribute VB_Name = "Module1"
Option Explicit

Const ID_ABOUT = 101
Const ID_EXIT = 103

Const ID_OPTIONS = 201

Public FCOLOR As Integer
Public BCOLOR As Integer
Public SIZE As Integer
Public FNAME As Integer



Public Sub Center(frm As Form)
Dim X As Integer
Dim Y As Integer

X = (Screen.Width - frm.Width) / 2
Y = (Screen.Height - frm.Height) / 2

frm.Left = X
frm.Top = Y
End Sub
Public Sub DBspeichern()

frmMain.File.Open "DeskClock.dat", fsModeOutput, fsAccessWrite, fsLockReadWrite
 
    frmMain.File.LinePrint FCOLOR
    frmMain.File.LinePrint BCOLOR
    frmMain.File.LinePrint SIZE
    frmMain.File.LinePrint FNAME

  frmMain.File.Close
End Sub
Public Sub DBladen()

frmMain.File.Open "DeskClock.dat", fsModeAppend
frmMain.File.Close

frmMain.File.Open "DeskClock.dat", fsModeInput, fsAccessRead, fsLockRead

Do While Not frmMain.File.EOF
    FCOLOR = frmMain.File.LineInputString
    BCOLOR = frmMain.File.LineInputString
    SIZE = frmMain.File.LineInputString
    FNAME = frmMain.File.LineInputString
   
  Loop
frmMain.File.Close

Exit Sub
End Sub

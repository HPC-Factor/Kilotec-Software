Attribute VB_Name = "mdlDatenbank"
Option Explicit
Public Filename As String
Const SIZE_SUPPORTSBUTTONS = 7500

Public Sub DBSpeichern()
Dim XX As Integer
frmMain.File.Open Filename, fsModeOutput, fsAccessWrite, fsLockReadWrite

For XX = 0 To frmMain.List1.ListCount - 1
    frmMain.File.LinePrint frmMain.List1.List(XX)
Next XX

frmMain.File.Close

End Sub
Public Sub DBladen()


On Error Resume Next
frmMain.File.Open Filename, fsModeInput, fsAccessRead, fsLockRead
frmMain.List1.Clear
Do While Not frmMain.File.EOF
    frmMain.List1.AddItem frmMain.File.LineInputString
Loop
 

frmMain.File.Close

Exit Sub
End Sub


Public Sub Center(frm As Form)
Dim X As Integer
Dim Y As Integer

X = (Screen.Width - frm.Width) / 2
Y = (Screen.Height - frm.Height) / 2

frm.Left = X
frm.Top = Y
End Sub

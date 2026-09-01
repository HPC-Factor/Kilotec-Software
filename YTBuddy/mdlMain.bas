Attribute VB_Name = "mdlMain"
Public Seconds As Integer
Public Minutes As Integer


Public Temp As String



Public Sub Center(frm As Form)
Dim X As Integer
Dim Y As Integer

X = (Screen.Width - frm.Width) / 2
Y = (Screen.Height - frm.Height) / 2

frm.Left = X
frm.Top = Y
End Sub


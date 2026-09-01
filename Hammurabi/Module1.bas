Attribute VB_Name = "Module1"
'Option Explicit

Public T
Public U
Public Z
Public P
Public S
Public H
Public E
Public Y
Public A
Public I
Public Q
Public B
Public Y2
Public C
Public QQ
Public D
Public Q1
Public D1
Public L
Public P1

Public temp1
Public temp2
Public temp3
Public temp4


Public Sub NewYear()
    Z = Z + 1
    P = P + I
    
    If Z = 11 Then
        frmGameOver.Show
        Exit Sub
    End If
    
    C = (Rnd(10) * 10) + 1
    Y = C + 17
    QQ = Int(10 * (2 * Rnd(1) - 0.3))
    'frmMain.Command2.Caption = QQ
    
    If QQ < 0 Then
        frmMain.PLAGE.Visible = True
        P = (P / 2)
    Else
        frmMain.PLAGE.Visible = False
    End If
    
    On Error Resume Next
    frmBuy.Text1.Text = ""
    frmSell.Text1.Text = ""
    frmFeed.Text1.Text = ""
    frmSeed.Text1.Text = ""
   
   Aktualisieren
    
End Sub

Public Sub Aktualisieren()
frmMain.Year.Caption = Z
frmMain.Starved.Caption = Int(D)
frmMain.Immigrants.Caption = Int(I)
frmMain.Acres.Caption = Int(A)
frmMain.Harvest.Caption = Int(Y2)
frmMain.Rats.Caption = Int(E)
frmMain.Land.Caption = Int(Y)
frmMain.Bushels.Caption = Int(S)
frmMain.Pop.Caption = Int(P)


End Sub

Public Sub Center(frm As Form)
Dim X As Integer
Dim Y As Integer

X = (Screen.Width - frm.Width) / 2
Y = (Screen.Height - frm.Height) / 2

frm.Left = X
frm.Top = 0 'Y
End Sub


Attribute VB_Name = "Module1"
Option Explicit
Const ID_NEW = 101
Const ID_EDIT = 102
Const ID_SAVE = 103
Const ID_CATEGORY = 104
Const ID_EXIT = 106

Const ID_LEARN = 201

Const ID_OPTIONS = 301

Const ID_HELP = 401
Const ID_INFO = 402

Const SIZE_SUPPORTSBUTTONS = 7500

Public strPath As String


Public n As Integer
Public EF As Integer
Public I As Integer
Public q As Integer

Public Interval As String
Public Autosave As String
Public RandomOrder As String
Public Spoiler As Integer

Public Sub AddToCombo(cbo As ComboBox)
 Dim strBuffer As String
 Dim Kat As String 'Einträge im Katalog (comboboxes)
Dim flKat As String
flKat = App.Path & "Category.txt"
  
 On Error Resume Next
  frmMain.File.Open flKat, fsModeAppend
  frmMain.File.Close
  'On Error Resume Next
    frmMain.File.Open flKat, fsModeInput, fsAccessRead, fsLockRead
  Do While Not frmMain.File.EOF
   Kat = frmMain.File.LineInputString
      'frmMain.Combo1.Text = frmMain.File.LineInputString
     If Kat <> Empty Then
            cbo.AddItem Kat
        End If
  Loop
 
  frmMain.File.Close
  
Exit Sub
End Sub

Public Sub EF_berechnen()
frmLearn.Command1.Enabled = False
frmLearn.Command2.Enabled = False
frmLearn.Command3.Enabled = False
frmLearn.Command4.Enabled = False
frmLearn.Command5.Enabled = False

If q >= 3 Then
    If n = 0 Then
        I = 1
    ElseIf n = 1 Then
        I = 6
    Else
        I = Round(I * EF)
    End If
    n = n + 1
Else
    n = 0
    I = 1
End If

EF = EF + (0.1 - (5 - q) * (0.08 + (5 - q) * 0.02))
If EF < 1.3 Then EF = 1.3

DBdaten(4, XX) = n
DBdaten(5, XX) = EF
DBdaten(6, XX) = I

If Index_Learn < frmMain.ListView1.ListItems.Count - 1 Then
    Index_Learn = Index_Learn + 1
Else
    Index_Learn = 0
End If


If Interval = "DAY" Then EinlesenEintraege_Day
If Interval = "HOUR" Then EinlesenEintraege_Hour
If Interval = "OFF" Then EinlesenEintraege_Off
End Sub

Public Sub Center(frm As Form)
Dim X As Integer
Dim Y As Integer

X = (Screen.Width - frm.Width) / 2
Y = (Screen.Height - frm.Height) / 2

frm.Left = X
frm.Top = Y
End Sub


Attribute VB_Name = "Module1"
Option Explicit
Const ID_ABOUT = 101
Const ID_SAVE = 103
Const ID_EXIT = 105

Const ID_ADD = 201
Const ID_EDIT = 202
Const ID_DELETE = 204


Const ID_KATEGORIE = 301
Const ID_STATISTIK = 302
Const ID_OPTIONS = 304

Const ID_FIND = 401


Public KategorieIndex As String

Public DBdaten(4, 600)

Public LANGUAGE As String
Public SYMBOL As String

'STATISTIK
Public GesamtBetrag As Long
Public Zeitraum As Integer
Public MaxBill As Integer
Public MinBill As Long

Public IndexX As Integer
Public Zaehler As Integer
Public XX As Integer
Public NeuerIndex As Long
Public ReplaceIndex As Long
Public T As Integer

Const SIZE_SUPPORTSBUTTONS = 7500



Public Sub Center(frm As Form)
Dim X As Integer
Dim Y As Integer

X = (Screen.Width - frm.Width) / 2
Y = (Screen.Height - frm.Height) / 2

frm.Left = X
frm.Top = Y
End Sub


Public Sub AddToCombo(cbo As ComboBox)
 Dim strBuffer As String
 Dim Kat As String 'Einträge im Katalog (comboboxes)
Dim flKat As String
flKat = App.Path & "KatBill.dat"
  
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
Public Sub AbInDieDatenbank()
    For XX = 0 To 500
    If DBdaten(0, XX) = "" And DBdaten(1, XX) = "" Then
    NeuerIndex = XX
    Exit For
    End If
    Next XX
    
        DBdaten(0, NeuerIndex) = frmAdd.Combo1.Text
        DBdaten(1, NeuerIndex) = frmAdd.txtBezeichnung.Text
        DBdaten(2, NeuerIndex) = frmAdd.txtDatum.Text
        DBdaten(3, NeuerIndex) = frmAdd.cboIntervall.List(frmAdd.cboIntervall.ListIndex)
        DBdaten(4, NeuerIndex) = frmAdd.txtBetrag.Text

Zaehler = Zaehler + 1

EinlesenEintraege

End Sub
Public Sub DBladen()

T = 0
frmMain.File.Open App.Path & "Bills.dat", fsModeAppend
frmMain.File.Close

frmMain.File.Open App.Path & "Bills.dat", fsModeInput, fsAccessRead, fsLockRead

Do While Not frmMain.File.EOF
    DBdaten(0, T) = frmMain.File.LineInputString
    DBdaten(1, T) = frmMain.File.LineInputString
    DBdaten(2, T) = frmMain.File.LineInputString
    DBdaten(3, T) = frmMain.File.LineInputString
    DBdaten(4, T) = frmMain.File.LineInputString
    T = T + 1
  Loop
 
  
Zaehler = T
frmMain.File.Close

Exit Sub
End Sub
Public Sub DBspeichern()

frmMain.File.Open App.Path & "Bills.dat", fsModeOutput, fsAccessWrite, fsLockReadWrite

 For XX = 0 To 500
If DBdaten(1, XX) <> "" Then
    frmMain.File.LinePrint DBdaten(0, XX)
    frmMain.File.LinePrint DBdaten(1, XX)
    frmMain.File.LinePrint DBdaten(2, XX)
    frmMain.File.LinePrint DBdaten(3, XX)
    frmMain.File.LinePrint DBdaten(4, XX)
 End If
 Next XX
  frmMain.File.Close
End Sub

Public Sub EinlesenEintraege()
Dim itmX As ListItem
Dim col As ColumnHeader
frmMain.ListView1.ColumnHeaders.Clear

If LANGUAGE = 1 Then
    frmMain.ListView1.ColumnHeaders.Add , "Col2", "Bezeichnung"
    frmMain.ListView1.ColumnHeaders.Add , "Col3", "Datum"
    frmMain.ListView1.ColumnHeaders.Add , "Col4", "Intervall"
    frmMain.ListView1.ColumnHeaders.Add , "Col5", "Betrag " & SYMBOL
    frmMain.ListView1.ColumnHeaders.Add , "Col6", "Nr."
Else
    frmMain.ListView1.ColumnHeaders.Add , "Col2", "Name"
    frmMain.ListView1.ColumnHeaders.Add , "Col3", "Date"
    frmMain.ListView1.ColumnHeaders.Add , "Col4", "Interval"
    frmMain.ListView1.ColumnHeaders.Add , "Col5", "Amount " & SYMBOL
    frmMain.ListView1.ColumnHeaders.Add , "Col6", "#"

End If
frmMain.ListView1.ListItems.Clear

frmMain.ListView1.ColumnHeaders(1).Width = 3000
frmMain.ListView1.ColumnHeaders(2).Width = 1000
frmMain.ListView1.ColumnHeaders(3).Width = 2000
frmMain.ListView1.ColumnHeaders(4).Width = 1000

frmMain.ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
frmMain.ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
frmMain.ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter



Dim Eintrage As Integer
Dim I As Integer
Eintrage = 0


   For XX = 0 To Zaehler
        If DBdaten(0, XX) = KategorieIndex And DBdaten(1, XX) <> Empty Then
            
            Set itmX = frmMain.ListView1.ListItems.Add(, , DBdaten(1, XX))
            itmX.SubItems(1) = DBdaten(2, XX)
            itmX.SubItems(2) = DBdaten(3, XX)
            itmX.SubItems(3) = DBdaten(4, XX)
            itmX.SubItems(4) = XX   'index
        End If
    Next XX
    

If frmMain.ListView1.ListItems.Count > 0 Then
    For XX = 0 To Zaehler
        'If DBdaten(2, XX) > Date Then frmMain.ListView1.ListItems(XX).SubItems(2).ForeColor = &HC0&          'ROT
        'If DBdaten(2, XX) < Date Then frmMain.ListView1.ListItems(XX).SubItems(2).ForeColor = &HC000&       'GRÜN
    Next XX
End If
    
'frmStatistik.Label1.Caption = frmMain.ListView1.ListItems.Count & " Eintrag in dieser Kategorie"
'frmStatistik.Label2.Caption = Zaehler & " Einträge insgesamt"
If LANGUAGE = 1 Then
    If Zaehler > 1 Then frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Einträge in " & frmMain.Combo1.Text
    If Zaehler = 1 Then frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Eintrag in " & frmMain.Combo1.Text
    If Zaehler = 0 Then frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Eintrag in " & frmMain.Combo1.Text
    If Zaehler > 1 Then frmMain.Label2.Caption = Zaehler & " Einträge insgesamt"
    If Zaehler = 1 Then frmMain.Label2.Caption = Zaehler & " Eintrag insgesamt"
    If Zaehler = 0 Then frmMain.Label2.Caption = Zaehler & " Einträge insgesamt"
Else
    If Zaehler > 1 Then frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " entries in " & frmMain.Combo1.Text
    If Zaehler = 1 Then frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " entry in " & frmMain.Combo1.Text
    If Zaehler = 0 Then frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " entries in " & frmMain.Combo1.Text
    
    If Zaehler > 1 Then frmMain.Label2.Caption = Zaehler & " total entries"
    If Zaehler = 1 Then frmMain.Label2.Caption = Zaehler & " total entry"
    If Zaehler = 0 Then frmMain.Label2.Caption = Zaehler & " total entries"
End If
End Sub

Public Sub LoescheEintrag()
' Einträge löschen
' Nur aus dem Array !
' Erst mit "Speichern" wird die DB aktualisiert
'Dim ind As Integer
ReplaceIndex = frmMain.ListView1.SelectedItem.SubItems(4)
   On Error Resume Next
    
    For XX = 0 To 4
       DBdaten(XX, ReplaceIndex) = ""
    Next XX

    Zaehler = Zaehler - 1
    EinlesenEintraege

End Sub
Public Sub DBbearbeiten()
' Geänderte Einträge übernehmen
       
       
        DBdaten(0, ReplaceIndex) = frmEdit.Combo1.Text
        DBdaten(1, ReplaceIndex) = frmEdit.txtBezeichnung.Text
        DBdaten(2, ReplaceIndex) = frmEdit.txtDatum.Text 'Name
        DBdaten(3, ReplaceIndex) = frmEdit.cboIntervall.Text
        DBdaten(4, ReplaceIndex) = frmEdit.txtBetrag.Text 'Name

    EinlesenEintraege

End Sub





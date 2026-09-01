Attribute VB_Name = "Module1"
Option Explicit
Public Pfad As String

Public KategorieIndex As String

Public DBdaten(5, 600)

Public IndexX As Integer
Public Zaehler As Integer
Public XX As Integer
Public NeuerIndex As Long
Public ReplaceIndex As Long
Public T As Integer

Const SIZE_SUPPORTSBUTTONS = 7500

Public Kontostand As Long
Public Einzahlungen As Long
Public Auszahlungen As Long
Public Saldo As Long



Public Sub Center(frm As Form)
Dim X As Integer
Dim Y As Integer

X = (Screen.Width - frm.Width) / 2
Y = (Screen.Height - frm.Height) / 2

frm.Left = X
frm.Top = 0 'Y
End Sub
Public Sub AddToCombo(cbo As ComboBox)
 Dim strBuffer As String
 Dim Kat As String
Dim flKat As String
flKat = App.Path & "Konto.dat"
  
 On Error Resume Next
  frmMain.File.Open flKat, fsModeAppend
  frmMain.File.Close
    frmMain.File.Open flKat, fsModeInput, fsAccessRead, fsLockRead
  Do While Not frmMain.File.EOF
   Kat = frmMain.File.LineInputString
     If Kat <> Empty Then
            cbo.AddItem Kat
        End If
  Loop
  frmMain.File.Close
Exit Sub
End Sub
Public Sub AddToCombo2(cbo As ComboBox)
 Dim strBuffer As String
 Dim Kat As String
Dim flKat As String
flKat = App.Path & "Zweck.dat"
  
 On Error Resume Next
  frmMain.File.Open flKat, fsModeAppend
  frmMain.File.Close
    frmMain.File.Open flKat, fsModeInput, fsAccessRead, fsLockRead
  Do While Not frmMain.File.EOF
   Kat = frmMain.File.LineInputString
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
    
        DBdaten(0, NeuerIndex) = frmAdd.Combo1.Text         'Konto
        DBdaten(1, NeuerIndex) = frmAdd.txtDatum.Text       'Datum
        If frmAdd.txtSaldo.Text <> "" Then
            DBdaten(2, NeuerIndex) = frmAdd.txtSaldo.Text       'Saldo
        Else
            DBdaten(2, NeuerIndex) = 0
        End If
        DBdaten(3, NeuerIndex) = frmAdd.cboZweck.Text       'Verwendungszweck
        DBdaten(4, NeuerIndex) = frmAdd.txtEinzahlung.Text  'Einzahlung
        DBdaten(5, NeuerIndex) = frmAdd.txtAuszahlung.Text  'Auszahlung
        
'Zähler
Zaehler = Zaehler + 1
        
' Aktualisieren
EinlesenEintraege

End Sub
Public Sub DBladen()

T = 0
frmMain.File.Open App.Path & "Kassenbuch.dat", fsModeAppend
frmMain.File.Close

frmMain.File.Open App.Path & "Kassenbuch.dat", fsModeInput, fsAccessRead, fsLockRead

Do While Not frmMain.File.EOF
    DBdaten(0, T) = frmMain.File.LineInputString
    DBdaten(1, T) = frmMain.File.LineInputString
    DBdaten(2, T) = frmMain.File.LineInputString
    DBdaten(3, T) = frmMain.File.LineInputString
    DBdaten(4, T) = frmMain.File.LineInputString
    DBdaten(5, T) = frmMain.File.LineInputString
    T = T + 1
  Loop
 
  
Zaehler = T
frmMain.File.Close

Exit Sub
End Sub
Public Sub DBspeichern()

frmMain.File.Open App.Path & "Kassenbuch.dat", fsModeOutput, fsAccessWrite, fsLockReadWrite

 For XX = 0 To 500
If DBdaten(1, XX) <> "" Then
    frmMain.File.LinePrint DBdaten(0, XX)
    frmMain.File.LinePrint DBdaten(1, XX)
    frmMain.File.LinePrint DBdaten(2, XX)
    frmMain.File.LinePrint DBdaten(3, XX)
    frmMain.File.LinePrint DBdaten(4, XX)
    frmMain.File.LinePrint DBdaten(5, XX)
 End If
 Next XX
  frmMain.File.Close
End Sub

Public Sub EinlesenEintraege()
Dim itmX As ListItem
Dim col As ColumnHeader
Dim XX2 As Integer
Dim Waehrung As String
Dim Waehrung2 As String

frmMain.ListView1.ColumnHeaders.Clear

frmMain.ListView1.ColumnHeaders.Add , "Col1", "v"
frmMain.ListView1.ColumnHeaders.Add , "Col2", "Nr."
frmMain.ListView1.ColumnHeaders.Add , "Col3", "Datum"
frmMain.ListView1.ColumnHeaders.Add , "Col4", "Buchung"
frmMain.ListView1.ColumnHeaders.Add , "Col5", "Einzahlung"
frmMain.ListView1.ColumnHeaders.Add , "Col6", "Auszahlung"

frmMain.ListView1.ColumnHeaders(1).Width = 300
frmMain.ListView1.ColumnHeaders(2).Width = 500
frmMain.ListView1.ColumnHeaders(3).Width = 1000
frmMain.ListView1.ColumnHeaders(4).Width = 3000
frmMain.ListView1.ColumnHeaders(5).Width = 1100
frmMain.ListView1.ColumnHeaders(6).Width = 1100

frmMain.ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
frmMain.ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
frmMain.ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
frmMain.ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter

frmMain.ListView1.ListItems.Clear
Kontostand = 0
Saldo = 0
Einzahlungen = 0
Auszahlungen = 0

   For XX = 0 To Zaehler
        Waehrung = ""
        Waehrung2 = ""
        
        If DBdaten(0, XX) = KategorieIndex And DBdaten(1, XX) <> Empty Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            
            Set itmX = frmMain.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
                 

            Saldo = Int(DBdaten(2, XX))
            If DBdaten(4, XX) <> "" Then
                Einzahlungen = Einzahlungen + DBdaten(4, XX)
            Else
                Einzahlungen = Einzahlungen + 0
            End If
            If DBdaten(5, XX) <> "" Then
                Auszahlungen = Auszahlungen + DBdaten(5, XX)
            Else
                Auszahlungen = Auszahlungen + 0
            End If
            
            Kontostand = Saldo + Einzahlungen - Auszahlungen

    
           End If
    Next XX
    
            frmMain.lblKontostand.Caption = "Kontostand: " & Kontostand & " €"
  
If frmMain.ListView1.ListItems.Count < 2 Then
    'frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Eintrag in dieser Kategorie"
Else
    'frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Einträge in dieser Kategorie"

End If
'frmMain.Label2.Caption = Zaehler & " Einträge insgesamt"
'frmMain.Label3.Caption = "Kategorie:"

End Sub

Public Sub LoescheEintrag()
' Einträge löschen
' Nur aus dem Array !
' Erst mit "Speichern" wird die DB aktualisiert
'Dim ind As Integer
ReplaceIndex = frmMain.ListView1.SelectedItem.SubItems(1)
   On Error Resume Next
    
    For XX = 0 To 5
       DBdaten(XX, ReplaceIndex) = ""
    Next XX

    Zaehler = Zaehler - 1
    EinlesenEintraege

End Sub
Public Sub DBbearbeiten()
' Geänderte Einträge übernehmen
       
        DBdaten(0, ReplaceIndex) = frmEdit.Combo1.Text         'Konto
        DBdaten(1, ReplaceIndex) = frmEdit.txtDatum.Text       'Datum
        If frmAdd.txtSaldo.Text <> "" Then
            DBdaten(2, ReplaceIndex) = frmEdit.txtSaldo.Text       'Saldo
        Else
            DBdaten(2, ReplaceIndex) = 0
        End If
        DBdaten(3, ReplaceIndex) = frmEdit.cboZweck.Text       'Verwendungszweck
        DBdaten(4, ReplaceIndex) = frmEdit.txtEinzahlung.Text  'Einzahlung
        DBdaten(5, ReplaceIndex) = frmEdit.txtAuszahlung.Text  'Auszahlung

    EinlesenEintraege

End Sub




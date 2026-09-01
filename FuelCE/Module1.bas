Attribute VB_Name = "Module1"
Option Explicit
Const ID_ABOUT = 101
Const ID_EXIT = 103

Const ID_ADD = 201
Const ID_EDIT = 202
Const ID_SEARCH = 203
Const ID_DELETE = 205
Const ID_SAVE = 207
Const ID_OVERVIEW = 209

Const ID_STATISTIK = 301
Const ID_EINZELHEITEN = 302
Const ID_BEWERTUNG = 304

Const ID_KATEGORIE = 401
Const ID_OPTIONEN = 403
Const ID_EXPORT = 405
Const ID_IMPORT = 406

Public KategorieIndex As String

Public DBdaten(5, 600)

Public IndexX As Integer
Public Zaehler As Integer
Public XX As Integer
Public NeuerIndex As Long
Public ReplaceIndex As Long
Public T As Integer

Public SPRACHE As Integer
Public WSYMBOL As String
Public MAXIM As Boolean
Public LITER As Boolean
Public KM As Boolean

Public KILOMETER As Integer
Public GESAMTKOSTEN As Double

Const SIZE_SUPPORTSBUTTONS = 7500



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
 Dim Kat As String 'Einträge im Katalog (comboboxes)
Dim flKat As String
flKat = App.Path & "\Car.dat"
  
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
    
        DBdaten(0, NeuerIndex) = Left(frmAdd.Combo1.Text, 2) 'KatalogNummer
        DBdaten(1, NeuerIndex) = Mid(frmAdd.Combo1.Text, 4) 'KatalogName
        DBdaten(2, NeuerIndex) = frmAdd.Kilometerstand.Text 'Name
        DBdaten(3, NeuerIndex) = frmAdd.Datum.Text 'MEMO
        DBdaten(4, NeuerIndex) = frmAdd.Getankt.Text 'MEMO
        DBdaten(5, NeuerIndex) = frmAdd.Preis.Text 'MEMO
          
           'index wird automatisch vergeben
        
'Zähler
Zaehler = Zaehler + 1
        
' Aktualisieren
EinlesenEintraege

End Sub


Public Sub DBladen()

T = 0
frmMain.File.Open App.Path & "\Fuel.dat", fsModeAppend
frmMain.File.Close

frmMain.File.Open App.Path & "\Fuel.dat", fsModeInput, fsAccessRead, fsLockRead

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

frmMain.File.Open App.Path & "\Fuel.dat", fsModeOutput, fsAccessWrite, fsLockReadWrite

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
Dim Ergebnis As Integer
Dim Length As Integer
Dim NewString As String

frmMain.ListView1.ColumnHeaders.Clear

If SPRACHE = 0 Then
    'frmMain.ListView1.ColumnHeaders.Add , "Col1", "Car"
    frmMain.ListView1.ColumnHeaders.Add , "Col1", "Mileage"
    frmMain.ListView1.ColumnHeaders.Add , "Col2", "Date"
    frmMain.ListView1.ColumnHeaders.Add , "Col3", "Refuel"
    frmMain.ListView1.ColumnHeaders.Add , "Col4", "Costs"
    frmMain.ListView1.ColumnHeaders.Add , "Col5", "Route"
    frmMain.ListView1.ColumnHeaders.Add , "Col6", "Total costs"
    frmMain.ListView1.ColumnHeaders.Add , "Col7", "Index"
End If

If SPRACHE = 1 Then
    'frmMain.ListView1.ColumnHeaders.Add , "Col1", "Auto"
    frmMain.ListView1.ColumnHeaders.Add , "Col1", "Kilometerstand"
    frmMain.ListView1.ColumnHeaders.Add , "Col2", "Datum"
    frmMain.ListView1.ColumnHeaders.Add , "Col3", "Getankt"
    frmMain.ListView1.ColumnHeaders.Add , "Col4", "Preis"
    frmMain.ListView1.ColumnHeaders.Add , "Col5", "Strecke"
    frmMain.ListView1.ColumnHeaders.Add , "Col6", "Gesamtkosten"
    frmMain.ListView1.ColumnHeaders.Add , "Col7", "Index"
End If

'frmMain.ListView1.ColumnHeaders(1).Width = 2200
frmMain.ListView1.ColumnHeaders(1).Width = 1500
frmMain.ListView1.ColumnHeaders(2).Width = 1000
frmMain.ListView1.ColumnHeaders(3).Width = 800
frmMain.ListView1.ColumnHeaders(4).Width = 800
frmMain.ListView1.ColumnHeaders(5).Width = 1200
frmMain.ListView1.ColumnHeaders(6).Width = 1500
frmMain.ListView1.ColumnHeaders(7).Width = 600

frmMain.ListView1.ListItems.Clear


   For XX = 0 To Zaehler
        If DBdaten(0, XX) = KategorieIndex And DBdaten(1, XX) <> Empty Then
            
           
            DBdaten(5, XX) = Replace(DBdaten(5, XX), ",", ".")
            GESAMTKOSTEN = CDbl(DBdaten(4, XX)) * CDbl(DBdaten(5, XX))
            Ergebnis = InStr(1, DBdaten(1, XX), "-")
            'Length =

            
            If XX = 0 Then KILOMETER = (CDbl(DBdaten(2, XX)) - Mid(DBdaten(1, XX), Ergebnis + 1, 7))
            If XX > 0 Then KILOMETER = CDbl(DBdaten(2, XX)) - CDbl(DBdaten(2, XX - 1))
            
            Set itmX = frmMain.ListView1.ListItems.Add(, , DBdaten(2, XX))
            itmX.SubItems(1) = DBdaten(3, XX)   'Datum
            itmX.SubItems(2) = DBdaten(4, XX) & " " & LITER  'Getankt
            itmX.SubItems(3) = DBdaten(5, XX) & " " & WSYMBOL & "/" & LITER  'Preis
            itmX.SubItems(4) = KILOMETER & " " & KM
            itmX.SubItems(5) = GESAMTKOSTEN & " " & WSYMBOL
            itmX.SubItems(6) = XX   'index
          End If
    Next XX

If SPRACHE = 0 Then
        frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " refueling for this vehicle"
        frmMain.Label2.Caption = Zaehler & " refuels total"
End If

If SPRACHE = 1 Then
    If frmMain.ListView1.ListItems.Count < 2 Then
        frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Betankung für dieses Fahrzeug"
    Else
        frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Betankungen für dieses Fahrzeug"
    End If
    frmMain.Label2.Caption = Zaehler & " Betankungen insgesamt"
End If
End Sub

Public Sub LoescheEintrag()
' Einträge löschen
' Nur aus dem Array !
' Erst mit "Speichern" wird die DB aktualisiert
'Dim ind As Integer
ReplaceIndex = frmMain.ListView1.SelectedItem.SubItems(3)
   On Error Resume Next
    
    For XX = 0 To 5
       DBdaten(XX, ReplaceIndex) = ""
    Next XX

    Zaehler = Zaehler - 1
    EinlesenEintraege

End Sub
Public Sub DBbearbeiten()
' Geänderte Einträge übernehmen
       
        DBdaten(0, ReplaceIndex) = Left(frmEdit.Combo1.Text, 2) 'KatalogNummer
        DBdaten(1, ReplaceIndex) = Mid(frmEdit.Combo1.Text, 4) 'KatalogName
        DBdaten(2, ReplaceIndex) = frmEdit.Bezeichnung.Text 'Name
        DBdaten(3, ReplaceIndex) = frmEdit.Anreise.Text 'MEMO
        DBdaten(4, ReplaceIndex) = frmEdit.Abreise.Text 'MEMO
        DBdaten(5, ReplaceIndex) = frmEdit.Bemerkungen.Text 'MEMO
        DBdaten(6, ReplaceIndex) = frmEdit.Combo2.Text 'MEMO
        DBdaten(7, ReplaceIndex) = frmEdit.Combo3.Text 'MEMO
        DBdaten(8, ReplaceIndex) = frmEdit.Ort.Text 'MEMO
        DBdaten(9, ReplaceIndex) = frmEdit.Land.Text 'MEMO

    EinlesenEintraege

End Sub


Public Sub ExportFile(ByVal FilePath As String)

   frmMain.File.Open FilePath, fsModeOutput, fsAccessWrite, fsLockReadWrite
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


Public Sub ImportFile(ByVal FilePath As String)

T = 0
frmMain.File.Open FilePath, fsModeOutput, fsAccessWrite, fsLockReadWrite

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
EinlesenEintraege
Exit Sub
End Sub


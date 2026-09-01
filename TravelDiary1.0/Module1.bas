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

Const ID_KOSTEN = 301
Const ID_EINZELHEITEN = 302
Const ID_BEWERTUNG = 304

Const ID_KATEGORIE = 401
Const ID_OPTIONEN = 403
Const ID_EXPORT = 405
Const ID_IMPORT = 406

Public KategorieIndex As String

Public DBdaten(23, 600)

Public IndexX As Integer
Public Zaehler As Integer
Public XX As Integer
Public NeuerIndex As Long
Public ReplaceIndex As Long
Public T As Integer

Public GK As Integer
Public GN As Double

Public AUTOSAVE As Boolean
Public TimerFlag As Integer
Public SAVE_MINUTEN As Integer
Public HINTERGRUNDFARBE As Integer
Public SPRACHE As Integer
Public WSYMBOL As String
Public MAXIM As Boolean

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
flKat = App.Path & "\Reiseart.dat"
  
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
Public Sub AddToCombo_Fortbewegungsmittel(cbo As ComboBox)
If SPRACHE = 0 Then
    cbo.AddItem "Car"
    cbo.AddItem "Plane"
    cbo.AddItem "Train"
    cbo.AddItem "Boat"
    cbo.AddItem "Walk"
    cbo.AddItem "Bike"
    cbo.AddItem "Subway"
    cbo.AddItem "Taxi"
    cbo.AddItem "Hitchhike"
    cbo.AddItem "Other"
End If
If SPRACHE = 1 Then
    cbo.AddItem "Auto"
    cbo.AddItem "Flugzeug"
    cbo.AddItem "Zug"
    cbo.AddItem "Schiff"
    cbo.AddItem "zu Fuß"
    cbo.AddItem "Fahrrad"
    cbo.AddItem "UBahn"
    cbo.AddItem "Taxi"
    cbo.AddItem "per Anhalter"
    cbo.AddItem "Sonstiges"
End If
End Sub
Public Sub AddToCombo_Uebernachtung(cbo As ComboBox)
If SPRACHE = 0 Then
    cbo.AddItem "Hotel"
    cbo.AddItem "Holiday home"
    cbo.AddItem "Domicile"
    cbo.AddItem "Camping"
    cbo.AddItem "Tent"
    cbo.AddItem "Caravan"
    cbo.AddItem "Mobile home"
    cbo.AddItem "Friends"
    cbo.AddItem "None"
    cbo.AddItem "Other"
End If
If SPRACHE = 1 Then
    cbo.AddItem "Hotel"
    cbo.AddItem "Ferienhaus"
    cbo.AddItem "Domizil"
    cbo.AddItem "Camping"
    cbo.AddItem "Zelt"
    cbo.AddItem "Wohnwagen"
    cbo.AddItem "Wohnmobil"
    cbo.AddItem "Freunde"
    cbo.AddItem "Keine"
    cbo.AddItem "Sonstiges"
End If
End Sub
Public Sub AddToCombo_Kosten(cbo As ComboBox)
    
    For XX = 0 To Zaehler - 1
    cbo.AddItem XX & " " & DBdaten(2, XX) & " " & DBdaten(8, XX)
    Next XX
    
End Sub
Public Sub AddToCombo_Bewertung(cbo As ComboBox)
If SPRACHE = 0 Then
    cbo.AddItem "0 - not applicable"
    cbo.AddItem "1 - very good"
    cbo.AddItem "2 - good"
    cbo.AddItem "3 - satisfactory"
    cbo.AddItem "4 - sufficient"
    cbo.AddItem "5 - deficient"
    cbo.AddItem "6 - insufficient"
End If
If SPRACHE = 1 Then
    cbo.AddItem "0 - nicht zutreffend"
    cbo.AddItem "1 - sehr gut"
    cbo.AddItem "2 - gut"
    cbo.AddItem "3 - befrieidigend"
    cbo.AddItem "4 - ausreichend"
    cbo.AddItem "5 - mangelhaft"
    cbo.AddItem "6 - ungenügend"
End If
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
        DBdaten(2, NeuerIndex) = frmAdd.Bezeichnung.Text 'Name
        DBdaten(3, NeuerIndex) = frmAdd.Anreise.Text 'MEMO
        DBdaten(4, NeuerIndex) = frmAdd.Abreise.Text 'MEMO
        DBdaten(5, NeuerIndex) = frmAdd.Bemerkungen.Text 'MEMO
        DBdaten(6, NeuerIndex) = frmAdd.Combo2.Text 'MEMO
        DBdaten(7, NeuerIndex) = frmAdd.Combo3.Text 'MEMO
        DBdaten(8, NeuerIndex) = frmAdd.Ort.Text 'MEMO
        DBdaten(9, NeuerIndex) = frmAdd.Land.Text 'MEMO
        
        DBdaten(10, NeuerIndex) = ""
        DBdaten(11, NeuerIndex) = ""
        DBdaten(12, NeuerIndex) = ""
        DBdaten(13, NeuerIndex) = ""
        DBdaten(14, NeuerIndex) = ""
        
        DBdaten(15, NeuerIndex) = ""
        DBdaten(16, NeuerIndex) = ""
        DBdaten(17, NeuerIndex) = ""
        DBdaten(18, NeuerIndex) = ""
        
        DBdaten(19, NeuerIndex) = ""
        DBdaten(20, NeuerIndex) = ""
        DBdaten(21, NeuerIndex) = ""
        DBdaten(22, NeuerIndex) = ""
        DBdaten(23, NeuerIndex) = ""
   
        
           'index wird automatisch vergeben
        
'Zähler
Zaehler = Zaehler + 1
        
' Aktualisieren
EinlesenEintraege

End Sub
Public Sub AbInDieDatenbank_Kosten()
    
        DBdaten(10, IndexX) = frmKosten.Reisekosten.Text
        DBdaten(11, IndexX) = frmKosten.Uebernachtung.Text
        DBdaten(12, IndexX) = frmKosten.Verpflegung.Text
        DBdaten(13, IndexX) = frmKosten.Freizeit.Text
        DBdaten(14, IndexX) = frmKosten.Sonstiges.Text

EinlesenEintraege

End Sub
Public Sub AbInDieDatenbank_Einzelheiten()
    
        DBdaten(15, IndexX) = frmEinzelheiten.Personen.Text
        DBdaten(16, IndexX) = frmEinzelheiten.Strecke.Text
        DBdaten(17, IndexX) = frmEinzelheiten.NameUnterkunft.Text
        DBdaten(18, IndexX) = frmEinzelheiten.Aktivitaeten.Text
        

EinlesenEintraege

End Sub
Public Sub AbInDieDatenbank_Bewertung()
    
        DBdaten(19, IndexX) = frmBewertung.Combo1.Text
        DBdaten(20, IndexX) = frmBewertung.Combo2.Text
        DBdaten(21, IndexX) = frmBewertung.Combo3.Text
        DBdaten(22, IndexX) = frmBewertung.Combo4.Text
        DBdaten(23, IndexX) = frmBewertung.Combo5.Text
        

EinlesenEintraege

End Sub
Public Sub DBladen()

T = 0
frmMain.File.Open App.Path & "\Travel.dat", fsModeAppend
frmMain.File.Close

frmMain.File.Open App.Path & "\Travel.dat", fsModeInput, fsAccessRead, fsLockRead

Do While Not frmMain.File.EOF
    DBdaten(0, T) = frmMain.File.LineInputString
    DBdaten(1, T) = frmMain.File.LineInputString
    DBdaten(2, T) = frmMain.File.LineInputString
    DBdaten(3, T) = frmMain.File.LineInputString
    DBdaten(4, T) = frmMain.File.LineInputString
    DBdaten(5, T) = frmMain.File.LineInputString
    DBdaten(6, T) = frmMain.File.LineInputString
    DBdaten(7, T) = frmMain.File.LineInputString
    DBdaten(8, T) = frmMain.File.LineInputString
    DBdaten(9, T) = frmMain.File.LineInputString
    
    DBdaten(10, T) = frmMain.File.LineInputString
    DBdaten(11, T) = frmMain.File.LineInputString
    DBdaten(12, T) = frmMain.File.LineInputString
    DBdaten(13, T) = frmMain.File.LineInputString
    DBdaten(14, T) = frmMain.File.LineInputString
    
    DBdaten(15, T) = frmMain.File.LineInputString
    DBdaten(16, T) = frmMain.File.LineInputString
    DBdaten(17, T) = frmMain.File.LineInputString
    DBdaten(18, T) = frmMain.File.LineInputString
    
    DBdaten(19, T) = frmMain.File.LineInputString
    DBdaten(20, T) = frmMain.File.LineInputString
    DBdaten(21, T) = frmMain.File.LineInputString
    DBdaten(22, T) = frmMain.File.LineInputString
    DBdaten(23, T) = frmMain.File.LineInputString

    T = T + 1
  Loop
 
  
Zaehler = T
frmMain.File.Close

Exit Sub
End Sub
Public Sub DBspeichern()

frmMain.File.Open App.Path & "\Travel.dat", fsModeOutput, fsAccessWrite, fsLockReadWrite

For XX = 0 To 500
If DBdaten(1, XX) <> "" Then
    frmMain.File.LinePrint DBdaten(0, XX)
    frmMain.File.LinePrint DBdaten(1, XX)
    frmMain.File.LinePrint DBdaten(2, XX)
    frmMain.File.LinePrint DBdaten(3, XX)
    frmMain.File.LinePrint DBdaten(4, XX)
    frmMain.File.LinePrint DBdaten(5, XX)
    frmMain.File.LinePrint DBdaten(6, XX)
    frmMain.File.LinePrint DBdaten(7, XX)
    frmMain.File.LinePrint DBdaten(8, XX)
    frmMain.File.LinePrint DBdaten(9, XX)
    
    frmMain.File.LinePrint DBdaten(10, XX)
    frmMain.File.LinePrint DBdaten(11, XX)
    frmMain.File.LinePrint DBdaten(12, XX)
    frmMain.File.LinePrint DBdaten(13, XX)
    frmMain.File.LinePrint DBdaten(14, XX)
    
    frmMain.File.LinePrint DBdaten(15, XX)
    frmMain.File.LinePrint DBdaten(16, XX)
    frmMain.File.LinePrint DBdaten(17, XX)
    frmMain.File.LinePrint DBdaten(18, XX)
    
    frmMain.File.LinePrint DBdaten(19, XX)
    frmMain.File.LinePrint DBdaten(20, XX)
    frmMain.File.LinePrint DBdaten(21, XX)
    frmMain.File.LinePrint DBdaten(22, XX)
    frmMain.File.LinePrint DBdaten(23, XX)
    
 End If
 Next XX
  frmMain.File.Close
End Sub

Public Sub EinlesenEintraege()
Dim itmX As ListItem
Dim col As ColumnHeader
Dim XX2 As Integer
frmMain.ListView1.ColumnHeaders.Clear

If SPRACHE = 0 Then
    frmMain.ListView1.ColumnHeaders.Add , "Col1", "Kind"
    frmMain.ListView1.ColumnHeaders.Add , "Col2", "Title"
    frmMain.ListView1.ColumnHeaders.Add , "Col3", "Destination"
    frmMain.ListView1.ColumnHeaders.Add , "Col4", "Index"
End If

If SPRACHE = 1 Then
    frmMain.ListView1.ColumnHeaders.Add , "Col1", "Art"
    frmMain.ListView1.ColumnHeaders.Add , "Col2", "Bezeichnung"
    frmMain.ListView1.ColumnHeaders.Add , "Col3", "Reiseziel"
    frmMain.ListView1.ColumnHeaders.Add , "Col4", "Index"
End If

frmMain.ListView1.ColumnHeaders(1).Width = 1550
frmMain.ListView1.ColumnHeaders(2).Width = 1550
frmMain.ListView1.ColumnHeaders(3).Width = 1550

frmMain.ListView1.ListItems.Clear

'frmMain.Question.Text = ""
'frmMain.Answer.Text = ""

   For XX = 0 To Zaehler
        If DBdaten(0, XX) = KategorieIndex And DBdaten(1, XX) <> Empty Then
            Set itmX = frmMain.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
            itmX.SubItems(1) = DBdaten(2, XX)   'Question
            itmX.SubItems(2) = DBdaten(8, XX)   'Answer
            itmX.SubItems(3) = XX   'index
          End If
    Next XX

If SPRACHE = 0 Then
    If frmMain.ListView1.ListItems.Count < 2 Then
        frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " item in this category"
    Else
        frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " items in this category"
    End If
    frmMain.Label2.Caption = Zaehler & " items total"
End If

If SPRACHE = 1 Then
    If frmMain.ListView1.ListItems.Count < 2 Then
        frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Eintrag in dieser Kategorie"
    Else
        frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Einträge in dieser Kategorie"
    End If
    frmMain.Label2.Caption = Zaehler & " Einträge insgesamt"
End If
End Sub

Public Sub LoescheEintrag()
' Einträge löschen
' Nur aus dem Array !
' Erst mit "Speichern" wird die DB aktualisiert
'Dim ind As Integer
ReplaceIndex = frmMain.ListView1.SelectedItem.SubItems(3)
   On Error Resume Next
    
    For XX = 0 To 23
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
    frmMain.File.LinePrint DBdaten(6, XX)
    frmMain.File.LinePrint DBdaten(7, XX)
    frmMain.File.LinePrint DBdaten(8, XX)
    frmMain.File.LinePrint DBdaten(9, XX)
    
    frmMain.File.LinePrint DBdaten(10, XX)
    frmMain.File.LinePrint DBdaten(11, XX)
    frmMain.File.LinePrint DBdaten(12, XX)
    frmMain.File.LinePrint DBdaten(13, XX)
    frmMain.File.LinePrint DBdaten(14, XX)
    
    frmMain.File.LinePrint DBdaten(15, XX)
    frmMain.File.LinePrint DBdaten(16, XX)
    frmMain.File.LinePrint DBdaten(17, XX)
    frmMain.File.LinePrint DBdaten(18, XX)
    
    frmMain.File.LinePrint DBdaten(19, XX)
    frmMain.File.LinePrint DBdaten(20, XX)
    frmMain.File.LinePrint DBdaten(21, XX)
    frmMain.File.LinePrint DBdaten(22, XX)
    frmMain.File.LinePrint DBdaten(23, XX)
    
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
    DBdaten(6, T) = frmMain.File.LineInputString
    DBdaten(7, T) = frmMain.File.LineInputString
    DBdaten(8, T) = frmMain.File.LineInputString
    DBdaten(9, T) = frmMain.File.LineInputString
    
    DBdaten(10, T) = frmMain.File.LineInputString
    DBdaten(11, T) = frmMain.File.LineInputString
    DBdaten(12, T) = frmMain.File.LineInputString
    DBdaten(13, T) = frmMain.File.LineInputString
    DBdaten(14, T) = frmMain.File.LineInputString
    
    DBdaten(15, T) = frmMain.File.LineInputString
    DBdaten(16, T) = frmMain.File.LineInputString
    DBdaten(17, T) = frmMain.File.LineInputString
    DBdaten(18, T) = frmMain.File.LineInputString
    
    DBdaten(19, T) = frmMain.File.LineInputString
    DBdaten(20, T) = frmMain.File.LineInputString
    DBdaten(21, T) = frmMain.File.LineInputString
    DBdaten(22, T) = frmMain.File.LineInputString
    DBdaten(23, T) = frmMain.File.LineInputString

    T = T + 1
  Loop
 
  
Zaehler = T
frmMain.File.Close
EinlesenEintraege
Exit Sub
End Sub
Public Sub EinlesenOverview()
Dim itmX As ListItem
Dim col As ColumnHeader
Dim XX2 As Integer
Dim GesamtkostenList As Double
Dim GesamtbewertungList As Double
Dim DauerList As Double
Dim Wieviel As Integer

frmOverview.ListView1.ColumnHeaders.Clear

If SPRACHE = 0 Then
    frmOverview.ListView1.ColumnHeaders.Add , "Col1", "Kind"
    frmOverview.ListView1.ColumnHeaders.Add , "Col2", "Title"
    frmOverview.ListView1.ColumnHeaders.Add , "Col3", "Destination"
    frmOverview.ListView1.ColumnHeaders.Add , "Col4", "Arrival"
    frmOverview.ListView1.ColumnHeaders.Add , "Col5", "Departure"
    frmOverview.ListView1.ColumnHeaders.Add , "Col6", "Duration"
    frmOverview.ListView1.ColumnHeaders.Add , "Col7", "Notes"
    frmOverview.ListView1.ColumnHeaders.Add , "Col8", "Traveling with"
    frmOverview.ListView1.ColumnHeaders.Add , "Col9", "Overnight stay"
    frmOverview.ListView1.ColumnHeaders.Add , "Col10", "Country"
    frmOverview.ListView1.ColumnHeaders.Add , "Col11", "Total costs"
    frmOverview.ListView1.ColumnHeaders.Add , "Col12", "Persons"
    frmOverview.ListView1.ColumnHeaders.Add , "Col13", "Route"
    frmOverview.ListView1.ColumnHeaders.Add , "Col14", "Lodging"
    frmOverview.ListView1.ColumnHeaders.Add , "Col15", "Activities"
    frmOverview.ListView1.ColumnHeaders.Add , "Col16", "Overall rating"
    frmOverview.ListView1.ColumnHeaders.Add , "Col17", "Index"
End If

If SPRACHE = 1 Then
    frmOverview.ListView1.ColumnHeaders.Add , "Col1", "Art"
    frmOverview.ListView1.ColumnHeaders.Add , "Col2", "Bezeichnung"
    frmOverview.ListView1.ColumnHeaders.Add , "Col3", "Reiseziel"
    frmOverview.ListView1.ColumnHeaders.Add , "Col4", "Anreise"
    frmOverview.ListView1.ColumnHeaders.Add , "Col5", "Abreise"
    frmOverview.ListView1.ColumnHeaders.Add , "Col6", "Dauer"
    frmOverview.ListView1.ColumnHeaders.Add , "Col7", "Bemerkungen"
    frmOverview.ListView1.ColumnHeaders.Add , "Col8", "Unterwegs mit"
    frmOverview.ListView1.ColumnHeaders.Add , "Col9", "Übernachtung"
    frmOverview.ListView1.ColumnHeaders.Add , "Col10", "Land"
    frmOverview.ListView1.ColumnHeaders.Add , "Col11", "Gesamtkosten"
    frmOverview.ListView1.ColumnHeaders.Add , "Col12", "Personen"
    frmOverview.ListView1.ColumnHeaders.Add , "Col13", "Strecke"
    frmOverview.ListView1.ColumnHeaders.Add , "Col14", "Unterkunft"
    frmOverview.ListView1.ColumnHeaders.Add , "Col15", "Aktivitäten"
    frmOverview.ListView1.ColumnHeaders.Add , "Col16", "Gesamtbewertung"
    frmOverview.ListView1.ColumnHeaders.Add , "Col17", "Index"
End If



frmOverview.ListView1.ColumnHeaders(1).Width = 1550
frmOverview.ListView1.ColumnHeaders(2).Width = 1550
frmOverview.ListView1.ColumnHeaders(3).Width = 1550

frmOverview.ListView1.ListItems.Clear



   For XX = 0 To Zaehler - 1
        GesamtkostenList = 0
        GesamtbewertungList = 0
        DauerList = 0
        Wieviel = 0
            
            
            If SPRACHE = 0 Then
                If DBdaten(19, XX) <> "0 - not applicable" Then Wieviel = Wieviel + 1
                If DBdaten(20, XX) <> "0 - not applicable" Then Wieviel = Wieviel + 1
                If DBdaten(21, XX) <> "0 - not applicable" Then Wieviel = Wieviel + 1
                If DBdaten(22, XX) <> "0 - not applicable" Then Wieviel = Wieviel + 1
                If DBdaten(23, XX) <> "0 - not applicable" Then Wieviel = Wieviel + 1
            End If
            If SPRACHE = 1 Then
                If DBdaten(19, XX) <> "0 - nicht zutreffend" Then Wieviel = Wieviel + 1
                If DBdaten(20, XX) <> "0 - nicht zutreffend" Then Wieviel = Wieviel + 1
                If DBdaten(21, XX) <> "0 - nicht zutreffend" Then Wieviel = Wieviel + 1
                If DBdaten(22, XX) <> "0 - nicht zutreffend" Then Wieviel = Wieviel + 1
                If DBdaten(23, XX) <> "0 - nicht zutreffend" Then Wieviel = Wieviel + 1
            End If
            
            
            GesamtkostenList = GesamtkostenList + CDbl(DBdaten(10, XX)) + CDbl(DBdaten(11, XX)) + CDbl(DBdaten(12, XX)) + CDbl(DBdaten(13, XX)) + CDbl(DBdaten(14, XX))
            
            GesamtbewertungList = GesamtbewertungList + (CDbl(Mid(DBdaten(19, XX), 1, 1)) + CDbl(Mid(DBdaten(20, XX), 1, 1)) + CDbl(Mid(DBdaten(21, XX), 1, 1)) + CDbl(Mid(DBdaten(22, XX), 1, 1)) + CDbl(Mid(DBdaten(23, XX), 1, 1))) / Wieviel
            
            If DBdaten(3, XX) <> Empty And DBdaten(4, XX) <> Empty Then
                DauerList = DateDiff("d", DBdaten(3, XX), DBdaten(4, XX))
            Else
                DauerList = ""
            End If
            
            Set itmX = frmOverview.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
            itmX.SubItems(1) = DBdaten(2, XX)   'Bezeichnung
            itmX.SubItems(2) = DBdaten(8, XX)   'Reiseziel
            itmX.SubItems(3) = DBdaten(3, XX)   'Anreise
            itmX.SubItems(4) = DBdaten(4, XX)   'Abreise
            itmX.SubItems(5) = DauerList       'Dauer
            itmX.SubItems(6) = DBdaten(5, XX)   'Bemerkungen
            itmX.SubItems(7) = DBdaten(6, XX)   'Unterwegs mit
            itmX.SubItems(8) = DBdaten(7, XX)   'AnsÜbernachtungwer
            itmX.SubItems(9) = DBdaten(9, XX)   'Land
            itmX.SubItems(10) = GesamtkostenList   'Gesamtkosten
            itmX.SubItems(11) = DBdaten(15, XX) 'Personen
            itmX.SubItems(12) = DBdaten(16, XX)   'Strecke
            itmX.SubItems(13) = DBdaten(17, XX)   'Unterkunft
            itmX.SubItems(14) = DBdaten(18, XX)   'Aktivitäten
            itmX.SubItems(15) = GesamtbewertungList   'Gesamtbewertung
            itmX.SubItems(16) = XX   'index
          'End If
    Next XX
  

End Sub

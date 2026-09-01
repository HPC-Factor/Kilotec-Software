Attribute VB_Name = "Module1"
Option Explicit
Const ID_ABOUT = 101
Const ID_EXIT = 103

Const ID_ADD = 201
Const ID_EDIT = 202
Const ID_DELETE = 204
Const ID_SAVE = 206

Const ID_KATEGORIE = 301
Const ID_COLUMN = 302
Const ID_OPTIONS = 304

Const ID_FIND = 401
'Const ID_INFO = 501

Public KategorieIndex As String

Public DBdaten(3, 600)
Public DBColumns(3, 600)

Public LANGUAGE As String
Public COLOR As Integer



Public Zaehler_Columns As Integer

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
frm.Top = 0 'Y
End Sub
Public Sub AddToCombo(cbo As ComboBox)
 Dim strBuffer As String
 Dim Kat As String 'Einträge im Katalog (comboboxes)
Dim flKat As String
flKat = "Kat.dat"
  
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
        DBdaten(2, NeuerIndex) = frmAdd.Entry1.Text 'Name
        DBdaten(3, NeuerIndex) = frmAdd.Entry2.Text 'MEMO
       
        'index wird automatisch vergeben
        
'Zähler
Zaehler = Zaehler + 1
        
' Aktualisieren
EinlesenEintraege

End Sub
Public Sub DBladen()

T = 0
frmMain.File.Open "SimplyDB.dat", fsModeAppend
frmMain.File.Close

frmMain.File.Open "SimplyDB.dat", fsModeInput, fsAccessRead, fsLockRead

Do While Not frmMain.File.EOF
    DBdaten(0, T) = frmMain.File.LineInputString
    DBdaten(1, T) = frmMain.File.LineInputString
    DBdaten(2, T) = frmMain.File.LineInputString
    DBdaten(3, T) = frmMain.File.LineInputString
    T = T + 1
  Loop
 
  
Zaehler = T
frmMain.File.Close

Exit Sub
End Sub
Public Sub DBspeichern()

frmMain.File.Open "SimplyDB.dat", fsModeOutput, fsAccessWrite, fsLockReadWrite

 For XX = 0 To 500
If DBdaten(1, XX) <> "" Then
    frmMain.File.LinePrint DBdaten(0, XX)
    frmMain.File.LinePrint DBdaten(1, XX)
    frmMain.File.LinePrint DBdaten(2, XX)
    frmMain.File.LinePrint DBdaten(3, XX)
 End If
 Next XX
  frmMain.File.Close
End Sub

Public Sub EinlesenEintraege()
Dim itmX As ListItem
Dim col As ColumnHeader
Dim XX2 As Integer
frmMain.ListView1.ColumnHeaders.Clear

If LANGUAGE = "English" Then
    frmMain.ListView1.ColumnHeaders.Add , "Col1", "Category"
End If
If LANGUAGE = "Deutsch" Then
    frmMain.ListView1.ColumnHeaders.Add , "Col1", "Kategorie"
End If

frmMain.ListView1.ColumnHeaders.Add , "Col2", "1"
frmMain.ListView1.ColumnHeaders.Add , "Col3", "2"
frmMain.ListView1.ColumnHeaders.Add , "Col4", "Index"

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
            itmX.SubItems(2) = DBdaten(3, XX)   'Answer
            itmX.SubItems(3) = XX   'index
            
            For XX2 = 0 To Zaehler_Columns
                If DBdaten(0, XX) = DBColumns(0, XX2) Then
                    frmMain.ListView1.ColumnHeaders(2).Text = DBColumns(2, XX2)
                    frmMain.ListView1.ColumnHeaders(3).Text = DBColumns(3, XX2)
                End If
            Next XX2
        End If
    Next XX

If frmMain.ListView1.ListItems.Count < 2 Then
    If LANGUAGE = "English" Then frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " entry of category"
    If LANGUAGE = "Deutsch" Then frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Eintrag in dieser Kategorie"

Else
    If LANGUAGE = "English" Then frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " entries of category"
    If LANGUAGE = "Deutsch" Then frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " Einträge in dieser Kategorie"

End If
If LANGUAGE = "English" Then frmMain.Label2.Caption = Zaehler & " entries total"
If LANGUAGE = "Deutsch" Then frmMain.Label2.Caption = Zaehler & " Einträge insgesamt"
If LANGUAGE = "English" Then frmMain.Label3.Caption = "Category:"
If LANGUAGE = "Deutsch" Then frmMain.Label3.Caption = "Kategorie:"

End Sub

Public Sub LoescheEintrag()
' Einträge löschen
' Nur aus dem Array !
' Erst mit "Speichern" wird die DB aktualisiert
'Dim ind As Integer
ReplaceIndex = frmMain.ListView1.SelectedItem.SubItems(3)
   On Error Resume Next
    
    For XX = 0 To 3
       DBdaten(XX, ReplaceIndex) = ""
    Next XX

    Zaehler = Zaehler - 1
    EinlesenEintraege

End Sub
Public Sub DBbearbeiten()
' Geänderte Einträge übernehmen
       
       
        DBdaten(0, ReplaceIndex) = Left(frmEdit.Combo1.Text, 2) 'KatalogNumm
        DBdaten(1, ReplaceIndex) = Mid(frmEdit.Combo1.Text, 4) 'KatalogName
        DBdaten(2, ReplaceIndex) = frmEdit.Entry1.Text 'Name
        DBdaten(3, ReplaceIndex) = frmEdit.Entry2.Text 'MEMO




    EinlesenEintraege

End Sub

Public Sub NeueColumns()
    For XX = 0 To 500
    If DBColumns(0, XX) = "" And DBColumns(1, XX) = "" Then
    NeuerIndex = XX
    Exit For
    End If
    Next XX
    
        DBColumns(0, NeuerIndex) = Left(frmColumn.Kategorie.Caption, 2) 'KatalogNummer
        DBColumns(1, NeuerIndex) = Mid(frmColumn.Kategorie.Caption, 4) 'KatalogName
        DBColumns(2, NeuerIndex) = frmColumn.Column1.Text 'Name
        DBColumns(3, NeuerIndex) = frmColumn.Column2.Text 'MEMO
       
        'index wird automatisch vergeben
        
'Zähler
Zaehler_Columns = Zaehler_Columns + 1
        
' Aktualisieren
EinlesenEintraege

End Sub

Public Sub Columnsladen()

T = 0
frmMain.File.Open "SimplyDB_C.dat", fsModeAppend
frmMain.File.Close

frmMain.File.Open "SimplyDB_C.dat", fsModeInput, fsAccessRead, fsLockRead

Do While Not frmMain.File.EOF
    DBColumns(0, T) = frmMain.File.LineInputString
    DBColumns(1, T) = frmMain.File.LineInputString
    DBColumns(2, T) = frmMain.File.LineInputString
    DBColumns(3, T) = frmMain.File.LineInputString
    T = T + 1
  Loop
 
  
Zaehler_Columns = T
frmMain.File.Close

Exit Sub
End Sub
Public Sub Columnsspeichern()

frmMain.File.Open "SimplyDB_C.dat", fsModeOutput, fsAccessWrite, fsLockReadWrite

 For XX = 0 To 500
If DBColumns(1, XX) <> "" Then
    frmMain.File.LinePrint DBColumns(0, XX)
    frmMain.File.LinePrint DBColumns(1, XX)
    frmMain.File.LinePrint DBColumns(2, XX)
    frmMain.File.LinePrint DBColumns(3, XX)
 End If
 Next XX
  frmMain.File.Close
End Sub


Public Sub Optionenladen()

frmMain.File.Open "Options.dat", fsModeAppend
frmMain.File.Close

frmMain.File.Open "Options.dat", fsModeInput, fsAccessRead, fsLockRead
'On Error Resume Next
Do While Not frmMain.File.EOF
    LANGUAGE = frmMain.File.LineInputString
    COLOR = frmMain.File.LineInputString
   

  Loop

frmMain.File.Close

Exit Sub
End Sub
Public Sub Optionenspeichern()

frmMain.File.Open "Options.dat", fsModeOutput, fsAccessWrite, fsLockReadWrite

    frmMain.File.LinePrint LANGUAGE
    frmMain.File.LinePrint COLOR


  frmMain.File.Close
End Sub

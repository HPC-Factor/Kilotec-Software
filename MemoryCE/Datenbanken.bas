Attribute VB_Name = "Datenbanken"
Option Explicit

Public KategorieIndex As String
Public Kategorie As String

Public DBdaten(8, 600)

Public IndexX As Integer
Public Zaehler As Integer
Public Index_Learn As Integer
Public XX As Integer
Public NeuerIndex As Long
Public ReplaceIndex As Long
Public T As Integer
Public Count As Integer


Public Sub AbInDieDatenbank()
    For XX = 0 To 500
    If DBdaten(0, XX) = "" And DBdaten(1, XX) = "" Then
    NeuerIndex = XX
    Exit For
    End If
    Next XX
    
        DBdaten(0, NeuerIndex) = Left(frmAdd.Combo1.Text, 2) 'KatalogNummer
        DBdaten(1, NeuerIndex) = Mid(frmAdd.Combo1.Text, 4) 'KatalogName
        DBdaten(2, NeuerIndex) = frmAdd.Question.Text 'Name
        DBdaten(3, NeuerIndex) = frmAdd.Answer.Text 'MEMO
        DBdaten(4, NeuerIndex) = 0
        DBdaten(5, NeuerIndex) = 2.5
        DBdaten(6, NeuerIndex) = 0
        DBdaten(7, NeuerIndex) = Date
        DBdaten(8, NeuerIndex) = Hour(Now)
        
        'index wird automatisch vergeben
        
'Zähler
Zaehler = Zaehler + 1
        
' Aktualisieren
EinlesenEintraege
If Autosave = "ON" Then DBspeichern
End Sub

Public Sub EinlesenEintraege()
Dim itmX As ListItem
Dim col As ColumnHeader
frmMain.ListView1.ColumnHeaders.Clear

frmMain.ListView1.ColumnHeaders.Add , "Col1", "Category"
frmMain.ListView1.ColumnHeaders.Add , "Col2", "Question"
frmMain.ListView1.ColumnHeaders.Add , "Col3", "Answer"
frmMain.ListView1.ColumnHeaders.Add , "Col4", "Rep."
frmMain.ListView1.ColumnHeaders.Add , "Col5", "EF"
frmMain.ListView1.ColumnHeaders.Add , "Col6", "Int."
frmMain.ListView1.ColumnHeaders.Add , "Col8", "Date"
frmMain.ListView1.ColumnHeaders.Add , "Col9", "Hour"
frmMain.ListView1.ColumnHeaders.Add , "Col10", "Index"

frmMain.ListView1.ColumnHeaders(1).Width = 1550
frmMain.ListView1.ColumnHeaders(2).Width = 1550
frmMain.ListView1.ColumnHeaders(3).Width = 1550
frmMain.ListView1.ColumnHeaders(4).Width = 550
frmMain.ListView1.ColumnHeaders(5).Width = 550
frmMain.ListView1.ColumnHeaders(6).Width = 550
frmMain.ListView1.ColumnHeaders(8).Width = 550
frmMain.ListView1.ColumnHeaders(9).Width = 600

frmMain.ListView1.ListItems.Clear

'frmMain.Question.Text = ""
'frmMain.Answer.Text = ""

If Spoiler = 1 Then
    For XX = 0 To Zaehler
        If DBdaten(0, XX) = KategorieIndex And DBdaten(1, XX) <> Empty Then
            Kategorie = DBdaten(1, XX)
            Set itmX = frmMain.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
            itmX.SubItems(1) = DBdaten(2, XX)   'Question
            itmX.SubItems(2) = DBdaten(3, XX)   'Answer
            itmX.SubItems(3) = DBdaten(4, XX)   'Rep
            itmX.SubItems(4) = DBdaten(5, XX)   'EF
            itmX.SubItems(5) = DBdaten(6, XX)   'I
            itmX.SubItems(6) = DBdaten(7, XX)   'Date
            itmX.SubItems(7) = DBdaten(8, XX)   'Date
            itmX.SubItems(8) = XX   'index
        End If
    Next XX
End If

If frmMain.ListView1.ListItems.Count < 2 Then
    frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " entry of category"
Else
    frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " entries of category"
End If
frmMain.Label2.Caption = Zaehler & " entries total"
'frmMain.Category.Caption = "Category: " & Kategorie
'frmMain.Item.Caption = "Item: " & IndexX + 1 & " of " & frmMain.ListView1.ListItems.Count
'frmMain.Label2.Caption = Zaehler & " total transactions"

End Sub
Public Sub EinlesenEintraege_Hour()

frmLearn.Question.Text = ""
frmLearn.Answer.Text = ""
On Error Resume Next

If RandomOrder = "ON" Then Index_Learn = Int(Rnd(1) * Zaehler)

For XX = Index_Learn To Zaehler
    If DBdaten(0, XX) = KategorieIndex And DBdaten(1, XX) <> Empty Then
          
       If Hour(Now) >= CDbl(DBdaten(8, XX)) + CDbl(DBdaten(4, XX)) Then
            frmLearn.Command6.Enabled = True
            frmLearn.Command7.Enabled = False
            
            frmLearn.Question.Text = DBdaten(2, XX)
           frmLearn.Label1.Caption = "Repetition: " & DBdaten(4, XX)
            frmLearn.Label2.Caption = "Easiness Factor: " & DBdaten(5, XX)
                        
            n = DBdaten(4, XX)
            EF = DBdaten(5, XX)
            I = DBdaten(6, XX)
            DBdaten(8, XX) = Hour(Now)
            DBdaten(7, XX) = Date
       
       Else
        If Index_Learn < frmMain.ListView1.ListItems.Count Then
        Index_Learn = Index_Learn + 1
       Else
        Index_Learn = 0
       End If
        EinlesenEintraege_Hour
        End If
       Exit For
     End If
Next XX

Ending_Session


End Sub
Public Sub EinlesenEintraege_Day()

frmLearn.Question.Text = ""
frmLearn.Answer.Text = ""
On Error Resume Next

If RandomOrder = "ON" Then Index_Learn = Int(Rnd(1) * Zaehler)

For XX = Index_Learn To Zaehler
    If DBdaten(0, XX) = KategorieIndex And DBdaten(1, XX) <> Empty Then
    
       
        If DateAdd("d", CDbl(DBdaten(4, XX)), DBdaten(7, XX)) = Date Or DateAdd("d", CDbl(DBdaten(4, XX)), DBdaten(7, XX)) < Date Then
            
            DBdaten(7, XX) = Date
            
            frmLearn.Command6.Enabled = True
            frmLearn.Command7.Enabled = False
            
            frmLearn.Question.Text = DBdaten(2, XX)
           frmLearn.Label1.Caption = "Repetition: " & DBdaten(4, XX)
            frmLearn.Label2.Caption = "Easiness Factor: " & DBdaten(5, XX)
                        
            n = DBdaten(4, XX)
            EF = DBdaten(5, XX)
            I = DBdaten(6, XX)
            DBdaten(8, XX) = Hour(Now)
       
       Else
        If Index_Learn < frmMain.ListView1.ListItems.Count Then
            Index_Learn = Index_Learn + 1
        Else
            Index_Learn = 0
            'Exit For
       End If
        EinlesenEintraege_Day
        End If
       Exit For
     End If
Next XX

Ending_Session

End Sub
Public Sub EinlesenEintraege_Off()

frmLearn.Question.Text = ""
frmLearn.Answer.Text = ""
On Error Resume Next

If RandomOrder = "ON" Then Index_Learn = Int(Rnd(1) * Zaehler)

For XX = Index_Learn To Zaehler
    If DBdaten(0, XX) = KategorieIndex And DBdaten(1, XX) <> Empty Then
        
       'If CDbl(DBdaten(8, XX)) + CDbl(DBdaten(4, XX)) = Hour(Now) Then
            frmLearn.Command6.Enabled = True
            frmLearn.Command7.Enabled = False
            
            frmLearn.Question.Text = DBdaten(2, XX)
            frmLearn.Label1.Caption = "Repetition: " & DBdaten(4, XX)
            frmLearn.Label2.Caption = "Easiness Factor: " & DBdaten(5, XX)
                        
            n = DBdaten(4, XX)
            EF = DBdaten(5, XX)
            I = DBdaten(6, XX)
            DBdaten(8, XX) = Hour(Now)
            DBdaten(7, XX) = Date
       
       'Else
       If Index_Learn < frmMain.ListView1.ListItems.Count Then
        Index_Learn = Index_Learn + 1
       Else
        Index_Learn = 0
       End If
        EinlesenEintraege_Off
        'End If
       Exit For
     End If
Next XX

Ending_Session


End Sub
Public Sub Ending_Session()
If frmLearn.Command1.Enabled = False And frmLearn.Command6.Enabled = False Then
    frmLearn.Command7.Enabled = True
    If Autosave = "ON" Then DBspeichern
End If
frmLearn.Category.Caption = "Category: " & Kategorie
frmLearn.Item.Caption = "Item: " & Index_Learn + 1 & " of " & frmMain.ListView1.ListItems.Count
End Sub

Public Sub DBladen()

T = 0
frmMain.File.Open App.Path & "MemoryCE.txt", fsModeAppend
frmMain.File.Close

frmMain.File.Open App.Path & "MemoryCE.txt", fsModeInput, fsAccessRead, fsLockRead

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
    
    'Zaehler = Zaehler + 1
    T = T + 1
  Loop
 
  
Zaehler = T
frmMain.File.Close

Exit Sub
End Sub
Public Sub DBspeichern()

frmMain.File.Open App.Path & "MemoryCE.txt", fsModeOutput, fsAccessWrite, fsLockReadWrite

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
 End If
 Next XX
  frmMain.File.Close
End Sub
Public Sub LoescheEintrag()

    For XX = 0 To 8
       DBdaten(XX, ReplaceIndex) = ""
    Next XX

   EinlesenEintraege
   Zaehler = Zaehler - 1

End Sub
Public Sub DBbearbeiten()
' Geänderte Einträge übernehmen

        DBdaten(0, ReplaceIndex) = Left(frmWork.Combo1.Text, 2) 'KatalogNumm
        DBdaten(1, ReplaceIndex) = Mid(frmWork.Combo1.Text, 4) 'KatalogName
        DBdaten(2, ReplaceIndex) = frmWork.Question.Text 'Name
        DBdaten(3, ReplaceIndex) = frmWork.Answer.Text 'MEMO
        If frmWork.Check1.Value = True Then
            DBdaten(4, ReplaceIndex) = 0
            DBdaten(5, ReplaceIndex) = 2.5
            DBdaten(6, ReplaceIndex) = 0
            DBdaten(7, ReplaceIndex) = Date
            DBdaten(8, ReplaceIndex) = Hour(Now)
       End If
    EinlesenEintraege

End Sub

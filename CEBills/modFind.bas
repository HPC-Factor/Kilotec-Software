Attribute VB_Name = "modFind"
Option Explicit

Public Sub Suchen()
Dim itmX As ListItem
Dim SuchText As String
SuchText = frmFind.Text1.Text
Dim Ergebnis As Integer
Dim XX2 As Integer

'On Error Resume Next





If LANGUAGE = 1 Then
    frmFind.ListView1.ColumnHeaders.Add , "Col2", "Bezeichnung"
    frmFind.ListView1.ColumnHeaders.Add , "Col3", "Datum"
    frmFind.ListView1.ColumnHeaders.Add , "Col4", "Intervall"
    frmFind.ListView1.ColumnHeaders.Add , "Col5", "Betrag " & SYMBOL
    frmFind.ListView1.ColumnHeaders.Add , "Col6", "Nr."
Else
    frmFind.ListView1.ColumnHeaders.Add , "Col2", "Name"
    frmFind.ListView1.ColumnHeaders.Add , "Col3", "Date"
    frmFind.ListView1.ColumnHeaders.Add , "Col4", "Interval"
    frmFind.ListView1.ColumnHeaders.Add , "Col5", "Amount " & SYMBOL
    frmFind.ListView1.ColumnHeaders.Add , "Col6", "#"

End If
frmFind.ListView1.ListItems.Clear
frmFind.ListView1.ColumnHeaders(1).Width = 2200
frmFind.ListView1.ColumnHeaders(2).Width = 1000
frmFind.ListView1.ColumnHeaders(3).Width = 2000
frmFind.ListView1.ColumnHeaders(4).Width = 1000

frmFind.ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
frmFind.ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
frmFind.ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter

'Me.MousePointer = vbHourglass

For XX = 0 To Zaehler
    If frmFind.Check1.Value = 1 Then
        If DBdaten(1, XX) = SuchText Or DBdaten(2, XX) = SuchText Or DBdaten(3, XX) = SuchText Or DBdaten(4, XX) = SuchText And DBdaten(0, XX) = frmFind.Combo1.Text Then
        
            Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(1, XX))
            itmX.SubItems(1) = DBdaten(2, XX)
            itmX.SubItems(2) = DBdaten(3, XX)
            itmX.SubItems(3) = DBdaten(4, XX)
            itmX.SubItems(4) = XX

        Ergebnis = 1
                
        End If
    ElseIf frmFind.Check1.Value = 0 Then
    
        If DBdaten(1, XX) = SuchText Or DBdaten(2, XX) = SuchText Or DBdaten(3, XX) = SuchText Or DBdaten(4, XX) = SuchText Then
        
            Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(1, XX))
            itmX.SubItems(1) = DBdaten(2, XX)
            itmX.SubItems(2) = DBdaten(3, XX)
            itmX.SubItems(3) = DBdaten(4, XX)
            itmX.SubItems(4) = XX
        
        
        'Suchergebnis
        Ergebnis = 1
        
        End If
    End If
    
Next XX
'Me.MousePointer = vbNormal

If Ergebnis <> 1 Then
If LANGUAGE = 0 Then MsgBox "Nothing found!", , "Info"
If LANGUAGE = 1 Then MsgBox "Nichts gefunden!", , "Info"
ElseIf Ergebnis = 1 Then
frmFind.Height = 2940
frmFind.Width = 7770
Center frmFind
End If

Exit Sub

'fehler:

'Me.MousePointer = vbNormal



End Sub



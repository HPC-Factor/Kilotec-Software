Attribute VB_Name = "modFind"
Option Explicit

Public Sub Suchen()
Dim itmX As ListItem
Dim SuchText As String
SuchText = frmFind.Text1.Text
Dim Ergebnis As Integer

'On Error Resume Next


frmFind.ListView1.ColumnHeaders.Clear
frmFind.ListView1.ListItems.Clear

If SPRACHE = 0 Then
    frmFind.ListView1.ColumnHeaders.Add , "Col1", "Kind"
    frmFind.ListView1.ColumnHeaders.Add , "Col2", "Title"
    frmFind.ListView1.ColumnHeaders.Add , "Col3", "Destination"
    frmFind.ListView1.ColumnHeaders.Add , "Col4", "Index"
End If

If SPRACHE = 1 Then
    frmFind.ListView1.ColumnHeaders.Add , "Col1", "Art"
    frmFind.ListView1.ColumnHeaders.Add , "Col2", "Bezeichnung"
    frmFind.ListView1.ColumnHeaders.Add , "Col3", "Reiseziel"
    frmFind.ListView1.ColumnHeaders.Add , "Col4", "Index"
End If


'Me.MousePointer = vbHourglass

For XX = 0 To Zaehler
    If frmFind.Check1.Value = 1 Then
        If DBdaten(2, XX) = SuchText And DBdaten(0, XX) = Mid(frmFind.Combo1.List(frmFind.Combo1.ListIndex), 1, 2) Then
        
            Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
            itmX.SubItems(1) = DBdaten(2, XX)
            itmX.SubItems(2) = DBdaten(8, XX)
            itmX.SubItems(3) = XX
           
        Ergebnis = 1
                
        End If
    ElseIf frmFind.Check1.Value = 0 Then
    
        If SPRACHE = 0 Then
            If DBdaten(2, XX) = SuchText And frmFind.Combo2.Text = "Title" Then
                Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
                itmX.SubItems(1) = DBdaten(2, XX)
                itmX.SubItems(2) = DBdaten(8, XX)
                itmX.SubItems(3) = XX
                Ergebnis = 1
             End If
            If DBdaten(8, XX) = SuchText And frmFind.Combo2.Text = "Destination" Then
                Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
                itmX.SubItems(1) = DBdaten(2, XX)
                itmX.SubItems(2) = DBdaten(8, XX)
                itmX.SubItems(3) = XX
                Ergebnis = 1
             End If
            If DBdaten(9, XX) = SuchText And frmFind.Combo2.Text = "Country" Then
                Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
                itmX.SubItems(1) = DBdaten(2, XX)
                itmX.SubItems(2) = DBdaten(8, XX)
                itmX.SubItems(3) = XX
                Ergebnis = 1
             End If
             If DBdaten(17, XX) = SuchText And frmFind.Combo2.Text = "Name of lodging" Then
                Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
                itmX.SubItems(1) = DBdaten(2, XX)
                itmX.SubItems(2) = DBdaten(8, XX)
                itmX.SubItems(3) = XX
                Ergebnis = 1
             End If
        End If
        If SPRACHE = 1 Then
            If DBdaten(2, XX) = SuchText And frmFind.Combo2.Text = "Bezeichnung" Then
                Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
                itmX.SubItems(1) = DBdaten(2, XX)
                itmX.SubItems(2) = DBdaten(8, XX)
                itmX.SubItems(3) = XX
                Ergebnis = 1
             End If
            If DBdaten(8, XX) = SuchText And frmFind.Combo2.Text = "Reiseziel" Then
                Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
                itmX.SubItems(1) = DBdaten(2, XX)
                itmX.SubItems(2) = DBdaten(8, XX)
                itmX.SubItems(3) = XX
                Ergebnis = 1
             End If
            If DBdaten(9, XX) = SuchText And frmFind.Combo2.Text = "Land" Then
                Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
                itmX.SubItems(1) = DBdaten(2, XX)
                itmX.SubItems(2) = DBdaten(8, XX)
                itmX.SubItems(3) = XX
                Ergebnis = 1
             End If
             If DBdaten(17, XX) = SuchText And frmFind.Combo2.Text = "Name Unterkunft" Then
                Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
                itmX.SubItems(1) = DBdaten(2, XX)
                itmX.SubItems(2) = DBdaten(8, XX)
                itmX.SubItems(3) = XX
                Ergebnis = 1
             End If
        End If
    
    
    
    End If
    
Next XX
'Me.MousePointer = vbNormal

If Ergebnis <> 1 Then
 If SPRACHE = 1 Then MsgBox "Nichts gefunden!", , "Info"
 If SPRACHE = 0 Then MsgBox "Nothing found!", , "Info"
ElseIf Ergebnis = 1 Then
frmFind.Height = 2940
frmFind.Width = 7770
End If

Exit Sub

'fehler:

'Me.MousePointer = vbNormal



End Sub

Public Sub Suchergebnis()

    Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
   
    frmFind.itmX.SubItems(1) = DBdaten(2, XX)
    frmFind.itmX.SubItems(2) = DBdaten(3, XX)
    frmFind.itmX.SubItems(3) = XX
End Sub

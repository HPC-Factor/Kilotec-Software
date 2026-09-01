Attribute VB_Name = "modFind"
Option Explicit

Public Sub Suchen()
Dim itmX As ListItem
Dim SuchText As String
SuchText = frmSuchen.Text1.Text
Dim Ergebnis As Integer
Dim Waehrung As String
Dim Waehrung2 As String

'On Error Resume Next

        frmSuchen.ListView1.ColumnHeaders.Clear
        frmSuchen.ListView1.ListItems.Clear
        
        frmSuchen.ListView1.ColumnHeaders.Add , "Col1", "v"
        frmSuchen.ListView1.ColumnHeaders.Add , "Col2", "Nr."
        frmSuchen.ListView1.ColumnHeaders.Add , "Col3", "Datum"
        frmSuchen.ListView1.ColumnHeaders.Add , "Col4", "Buchung"
        frmSuchen.ListView1.ColumnHeaders.Add , "Col5", "Einzahlung"
        frmSuchen.ListView1.ColumnHeaders.Add , "Col6", "Auszahlung"
        
        frmSuchen.ListView1.ColumnHeaders(1).Width = 300
        frmSuchen.ListView1.ColumnHeaders(2).Width = 500
        frmSuchen.ListView1.ColumnHeaders(3).Width = 1000
        frmSuchen.ListView1.ColumnHeaders(4).Width = 3000
        frmSuchen.ListView1.ColumnHeaders(5).Width = 1100
        frmSuchen.ListView1.ColumnHeaders(6).Width = 1100
        
        frmSuchen.ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
        frmSuchen.ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
        frmSuchen.ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
        frmSuchen.ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
        



For XX = 0 To Zaehler '- 1
    Waehrung = ""
    Waehrung2 = ""
    
    If frmSuchen.Check1.Value = 1 Then
        

    
        If DBdaten(1, XX) = SuchText And frmSuchen.Combo2.List(frmSuchen.Combo2.ListIndex) = "Datum" And DBdaten(0, XX) = frmSuchen.Combo1.List(frmSuchen.Combo1.ListIndex) Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            Set itmX = frmSuchen.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
            Ergebnis = 1
        End If
        If DBdaten(3, XX) = SuchText And frmSuchen.Combo2.List(frmSuchen.Combo2.ListIndex) = "Verwendungszweck" And DBdaten(0, XX) = frmSuchen.Combo1.List(frmSuchen.Combo1.ListIndex) Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            Set itmX = frmSuchen.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
            Ergebnis = 1
        End If
        If DBdaten(4, XX) = SuchText And frmSuchen.Combo2.List(frmSuchen.Combo2.ListIndex) = "Einzahlungen" And DBdaten(0, XX) = frmSuchen.Combo1.List(frmSuchen.Combo1.ListIndex) Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            Set itmX = frmSuchen.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
            Ergebnis = 1
        End If
        If DBdaten(5, XX) = SuchText And frmSuchen.Combo2.List(frmSuchen.Combo2.ListIndex) = "Auszahlungen" And DBdaten(0, XX) = frmSuchen.Combo1.List(frmSuchen.Combo1.ListIndex) Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            Set itmX = frmSuchen.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
            Ergebnis = 1
        End If
        If XX = SuchText And frmSuchen.Combo2.List(frmSuchen.Combo2.ListIndex) = "Buchungsnummer" And DBdaten(0, XX) = frmSuchen.Combo1.List(frmSuchen.Combo1.ListIndex) Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            Set itmX = frmSuchen.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
            Ergebnis = 1
        End If
    
    
    ElseIf frmSuchen.Check1.Value = 0 Then
        
        
        If DBdaten(1, XX) = SuchText And frmSuchen.Combo2.Text = "Datum" Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            Set itmX = frmSuchen.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
            Ergebnis = 1
        End If
        If DBdaten(3, XX) = SuchText And frmSuchen.Combo2.Text = "Verwendungszweck" Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            Set itmX = frmSuchen.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
            Ergebnis = 1
        End If
        If DBdaten(4, XX) = SuchText And frmSuchen.Combo2.Text = "Einzahlungen" Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            Set itmX = frmSuchen.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
            Ergebnis = 1
        End If
        If DBdaten(5, XX) = SuchText And frmSuchen.Combo2.Text = "Auszahlungen" Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            Set itmX = frmSuchen.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
            Ergebnis = 1
        End If
        If XX = SuchText And frmSuchen.Combo2.List(frmSuchen.Combo2.ListIndex) = "Buchungsnummer" Then
            If DBdaten(4, XX) <> "" Then Waehrung = " €"
            If DBdaten(5, XX) <> "" Then Waehrung2 = " €"
            Set itmX = frmSuchen.ListView1.ListItems.Add(, , "")
            itmX.SubItems(1) = XX
            itmX.SubItems(2) = DBdaten(1, XX)
            itmX.SubItems(3) = DBdaten(3, XX)
            itmX.SubItems(4) = DBdaten(4, XX) & Waehrung
            itmX.SubItems(5) = DBdaten(5, XX) & Waehrung2
            Ergebnis = 1
        End If

    End If
    
Next XX
'Me.MousePointer = vbNormal

If Ergebnis <> 1 Then
MsgBox "Nichts gefunden!", , "Info"
ElseIf Ergebnis = 1 Then
frmSuchen.Height = 3660

End If

Exit Sub



End Sub

Public Sub Suchergebnis()

    Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
   
    frmFind.itmX.SubItems(1) = DBdaten(2, XX)
    frmFind.itmX.SubItems(2) = DBdaten(3, XX)
    frmFind.itmX.SubItems(3) = XX
End Sub

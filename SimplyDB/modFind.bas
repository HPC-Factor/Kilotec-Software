Attribute VB_Name = "modFind"
Option Explicit

Public Sub Suchen()
Dim itmX As ListItem
Dim SuchText As String
SuchText = frmFind.Text1.Text
Dim Ergebnis As Integer
Dim XX2 As Integer

'On Error Resume Next


frmFind.ListView1.ColumnHeaders.Clear
frmFind.ListView1.ListItems.Clear

If LANGUAGE = "English" Then
    frmFind.ListView1.ColumnHeaders.Add , "Col1", "Category"
End If
If LANGUAGE = "Deutsch" Then
    frmFind.ListView1.ColumnHeaders.Add , "Col1", "Kategorie"
End If
frmFind.ListView1.ColumnHeaders.Add , "Col2", "1"
frmFind.ListView1.ColumnHeaders.Add , "Col3", "2"
frmFind.ListView1.ColumnHeaders.Add , "Col4", "Index"

'Me.MousePointer = vbHourglass

For XX = 0 To Zaehler
    If frmFind.Check1.Value = 1 Then
        If DBdaten(2, XX) = SuchText And DBdaten(0, XX) = Mid(frmFind.Combo1.List(frmFind.Combo1.ListIndex), 1, 2) Then
        
            Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
            itmX.SubItems(1) = DBdaten(2, XX)
            itmX.SubItems(2) = DBdaten(3, XX)
            itmX.SubItems(3) = XX
            
            For XX2 = 0 To Zaehler_Columns
                If DBdaten(0, XX) = DBColumns(0, XX2) Then
                    frmFind.ListView1.ColumnHeaders(2).Text = DBColumns(2, XX2)
                    frmFind.ListView1.ColumnHeaders(3).Text = DBColumns(3, XX2)
                End If
            Next XX2
        'Suchergebnis
        Ergebnis = 1
                
        End If
    ElseIf frmFind.Check1.Value = 0 Then
    
        If DBdaten(2, XX) = SuchText Then
        
        
            Set itmX = frmFind.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))
            itmX.SubItems(1) = DBdaten(2, XX)
            itmX.SubItems(2) = DBdaten(3, XX)
            itmX.SubItems(3) = XX
        
        For XX2 = 0 To Zaehler_Columns
                If DBdaten(0, XX) = DBColumns(0, XX2) Then
                    frmFind.ListView1.ColumnHeaders(2).Text = DBColumns(2, XX2)
                    frmFind.ListView1.ColumnHeaders(3).Text = DBColumns(3, XX2)
                End If
            Next XX2
        
        'Suchergebnis
        Ergebnis = 1
        
        End If
    End If
    
Next XX
'Me.MousePointer = vbNormal

If Ergebnis <> 1 Then
If LANGUAGE = "English" Then MsgBox "Nothing found!", , "Info"
If LANGUAGE = "Deutsch" Then MsgBox "Nichts gefunden!", , "Info"
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

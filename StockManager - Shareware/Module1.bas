Attribute VB_Name = "Module1"
'Option Explicit

Public DBdaten(5, 600)

Public IndexX As Integer
Public Zaehler As Integer

Public Temp As String
Public XX As Integer
Public YY As Integer
Public T As Integer
Public FF As Integer

Public BuySell As String



Public NeuerIndex As Long
Public ReplaceIndex As Long
Public KategorieIndex As String


Public Sub AddToCombo(cbo As ComboBox)
 Dim strBuffer As String
 Dim Kat As String 'Einträge im Katalog (comboboxes)
Dim flKat As String
flKat = "Stocks.dat"


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
' Speichere einen neuen Datensatz
' Allerdings voerst nur im Array !
' Erst mit "Speichern" wird in die DB übernommen.



    For XX = 0 To 500
    If DBdaten(0, XX) = "" And DBdaten(1, XX) = "" Then
    NeuerIndex = XX
    Exit For
    End If
    Next XX
    
        If frmAdd.Option1.Value = True Then BuySell = "Buy"
        If frmAdd.Option2.Value = True Then BuySell = "Sell"
        
        DBdaten(0, NeuerIndex) = Left(frmAdd.Combo1.Text, 2) 'KatalogNummer
        DBdaten(1, NeuerIndex) = Mid(frmAdd.Combo1.Text, 4) 'KatalogName
        DBdaten(2, NeuerIndex) = frmAdd.Datum.Text 'Name
        DBdaten(3, NeuerIndex) = frmAdd.Price.Text 'MEMO
        DBdaten(4, NeuerIndex) = frmAdd.Shares.Text 'MEMO
        DBdaten(5, NeuerIndex) = BuySell
        
        
        'index wird automatisch vergeben
        
'Zähler
Zaehler = Zaehler + 1
        
' Aktualisieren
EinlesenEintraege

End Sub

Public Sub DBladen()

T = 0
frmMain.File.Open "Invests.dat", fsModeAppend
frmMain.File.Close

frmMain.File.Open "Invests.dat", fsModeInput, fsAccessRead, fsLockRead

Do While Not frmMain.File.EOF
   DBdaten(0, T) = frmMain.File.LineInputString
   DBdaten(1, T) = frmMain.File.LineInputString
   DBdaten(2, T) = frmMain.File.LineInputString
   DBdaten(3, T) = frmMain.File.LineInputString
   DBdaten(4, T) = frmMain.File.LineInputString
   DBdaten(5, T) = frmMain.File.LineInputString
    
    'Zaehler = Zaehler + 1
    T = T + 1
  Loop
 
  


Zaehler = T
frmMain.File.Close

Exit Sub
End Sub


Public Sub EinlesenEintraege()
' Lese in Listview ein;
' Möglich wäre auch eine Listbox

Dim itmX As ListItem
Dim col As ColumnHeader
frmMain.ListView1.ColumnHeaders.Clear

frmMain.ListView1.ColumnHeaders.Add , "Col1", "Stock"
frmMain.ListView1.ColumnHeaders.Add , "Col2", "Date"
frmMain.ListView1.ColumnHeaders.Add , "Col3", "Price"
frmMain.ListView1.ColumnHeaders.Add , "Col4", "Shares"
frmMain.ListView1.ColumnHeaders.Add , "Col5", "Buy/Sell"
frmMain.ListView1.ColumnHeaders.Add , "Col6", "Index"

frmMain.ListView1.ListItems.Clear


For XX = 0 To Zaehler

If DBdaten(0, XX) = KategorieIndex And DBdaten(1, XX) <> Empty Then

Set itmX = frmMain.ListView1.ListItems.Add(, , DBdaten(0, XX) & " " & DBdaten(1, XX))

itmX.SubItems(1) = DBdaten(2, XX) 'Date
itmX.SubItems(2) = DBdaten(3, XX)   'Price
itmX.SubItems(3) = DBdaten(4, XX)   'Shares
itmX.SubItems(4) = DBdaten(5, XX)   'BuySell
itmX.SubItems(5) = XX   'index
End If


Next XX

frmMain.Label1.Caption = frmMain.ListView1.ListItems.Count & " transaction of this stock"
frmMain.Label2.Caption = Zaehler & " total transactions"

End Sub
Public Sub DBspeichern()

frmMain.File.Open "Invests.dat", fsModeOutput, fsAccessWrite, fsLockReadWrite

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
Public Sub LoescheEintrag()
' Einträge löschen
' Nur aus dem Array !
' Erst mit "Speichern" wird die DB aktualisiert
'Dim ind As Integer
'ind = frmMain.ListView1.SelectedItem.SubItems(4)
    For XX = 0 To 5
       DBdaten(XX, ReplaceIndex) = ""
    Next XX

   EinlesenEintraege


    Zaehler = Zaehler - 1

End Sub
Public Sub DBbearbeiten()
' Geänderte Einträge übernehmen

        If frmWork.Option1.Value = True Then BuySell = "Buy"
        If frmWork.Option2.Value = True Then BuySell = "Sell"
        
        
        DBdaten(0, ReplaceIndex) = Left(frmWork.Combo1.Text, 2) 'KatalogNumm
        DBdaten(1, ReplaceIndex) = Mid(frmWork.Combo1.Text, 4) 'KatalogName
        DBdaten(2, ReplaceIndex) = frmWork.Datum.Text 'Name
        DBdaten(3, ReplaceIndex) = frmWork.Price.Text 'MEMO
        DBdaten(4, ReplaceIndex) = frmWork.Shares.Text 'MEMO
        DBdaten(5, ReplaceIndex) = BuySell



    EinlesenEintraege

End Sub
Public Sub EinlesenStatistik()
Dim Total As Integer
Dim TotalBuyAmount As Integer
Dim TotalSellAmount As Integer
Dim TotalBoughtShares As Integer
Dim TotalSoldShares As Integer
Dim TotalHoldShares As Integer

For XX = 0 To Zaehler
    If DBdaten(0, XX) = KategorieIndex And DBdaten(1, XX) <> Empty Then
        Total = Total + 1
        If DBdaten(5, XX) = "Buy" Then TotalBuyAmount = TotalBuyAmount + (DBdaten(4, XX) * DBdaten(3, XX))
        If DBdaten(5, XX) = "Sell" Then TotalSellAmount = TotalSellAmount + (DBdaten(4, XX) * DBdaten(3, XX))
        
        If DBdaten(5, XX) = "Buy" Then TotalBoughtShares = TotalBoughtShares + CDbl(DBdaten(4, XX))
        If DBdaten(5, XX) = "Sell" Then TotalSoldShares = TotalSoldShares + CDbl(DBdaten(4, XX))
        
        TotalHoldShares = TotalBoughtShares - TotalSoldShares
        
    
    
    
    End If
Next XX


frmStat.TotalTrans.Caption = Total
frmStat.TotalBuy.Caption = TotalBuyAmount
frmStat.TotalSell.Caption = TotalSellAmount
frmStat.Profit.Caption = TotalSellAmount - TotalBuyAmount

frmStat.BShares.Caption = TotalBoughtShares
frmStat.SShares.Caption = TotalSoldShares
frmStat.SHold.Caption = TotalHoldShares
End Sub

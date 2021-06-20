Attribute VB_Name = "Module1"

Public Sub VBAofWallStreet()

Dim ticker As String
Dim i As Long
Dim lastrow As Long
Dim tickervolume As LongLong
Dim tickername As String
Dim tickertablerow As Integer
Dim firstopen As Double
Dim lastclose As Double


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
'Assign Table Headers

lastrow = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
tickertablerow = Cells(Rows.Count, 9).End(xlUp).Row
firstopen = Cells(2, 3).Value

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    'If the Ticker has changed from the previous
    
    tickername = Cells(i, 1).Value
    'Set Ticker Name
    
    tickervolume = tickervolume + Cells(i, 7).Value
    'Add to Stock Volume
    
    lastclose = Cells(i, 6).Value
    'Store last close price
    
    Range("I" & tickertablerow + 1).Value = tickername
    'Enter Ticker into Tavle
    
    Range("L" & tickertablerow + 1).Value = tickervolume
    'Enter Stock volume into table
    
    Range("J" & tickertablerow + 1).Value = (lastclose - firstopen)
    'Enter Yearly Change
    
    Range("K" & tickertablerow + 1).Value = (lastclose - firstopen) / firstopen
    'Enter percent change
    
    Cells(tickertablerow + 1, 11).NumberFormat = "0.00%"
    'Cells(tickertablerow + 1, 11).Style = "Percent"
    'Format Column for Percentage
    
    tickertablerow = Cells(Rows.Count, 9).End(xlUp).Row
    'Update last row of table
    
    tickervolume = 0
    'Reset Stock volume
    
    firstopen = Cells(i + 1, 3).Value
    'Store first open price
     
    Else
    'If the Ticker is the same as the previous
    
    tickervolume = tickervolume + Cells(i, 7).Value
    'Update stock volume
    
    End If
    
Next i
    

'Ticker Symbol - Value of Ticker Cell
'Yearly Change = Last entry - First Entry (by date)
'Percent Change = (Close - Open)/100
'Total Stock Volume = Sum of shared Ticker
'Cells(i + 1, 1).Value <> Cells(i, 1).Value
End Sub
Public Sub VBAofWallStreet()
Dim ticker As String
Dim i As Integer
Dim lastrow As Integer
Dim tickervolume As Integer


lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    'If the Ticker has changed from the previous
    
    Else
    'If the Ticker is the same as the previous
    
    tickervolume = tickervolume + Cells(i, 7)
    
    

Next i
'Ticker Symbol - Value of Ticker Cell
'Yearly Change = Last entry - First Entry (by date)
'Percent Change = (Close - Open)/100
'Total Stock Volume = Sum of shared Ticker
'Cells(i + 1, 1).Value <> Cells(i, 1).Value
End Sub

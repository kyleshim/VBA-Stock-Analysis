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
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvolume As LongLong
Dim greatestincreaseticker As String
Dim greatestdecreaseticker As String
Dim greatestvolumeticker As String
'Define variable types

sheet_count = Sheets.Count
'Determine # of sheets

For ws = 1 To sheet_count
'Cycle throughsheets

    Sheets(ws).Activate
    'Ensure current worksheet will be modified

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    'Assign Table Headers

    tickertablerow = 0
    lastrow = 0
    greatestincrease = 0
    greatestdecrease = 0
    greatestvolume = 0
    firstopen = Cells(2, 3).Value
    'Save First Open Stock value

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    tickertablerow = Cells(Rows.Count, 9).End(xlUp).Row
    'Set Row Counts

    For i = 2 To lastrow
    'Loop through all rows
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'If the Ticker has changed from the previous
        
            tickername = Cells(i, 1).Value
            'Set Ticker Name
        
            tickervolume = tickervolume + Cells(i, 7).Value
            'Add to Stock Volume
        
            lastclose = Cells(i, 6).Value
            'Store last close price
        
            Range("I" & tickertablerow + 1).Value = tickername
            'Enter Ticker into Table
        
            Range("J" & tickertablerow + 1).Value = (lastclose - firstopen)
            'Enter Yearly Change
        
            If Cells(tickertablerow + 1, 10).Value > 0 Then
            
                Cells(tickertablerow + 1, 10).Interior.ColorIndex = 4
            
            ElseIf Cells(tickertablerow + 1, 10).Value < 0 Then
                Cells(tickertablerow + 1, 10).Interior.ColorIndex = 3
            
            End If
            'Format cell as green or red
        
            If firstopen = 0 Then
            'Prevent dividing by 0
            
                Range("K" & tickertablerow + 1).Value = 0
            
            Else
            
                Range("K" & tickertablerow + 1).Value = (lastclose - firstopen) / firstopen
                'Enter percent change
            
                If Cells(tickertablerow + 1, 11).Value > greatestincrease Then
                
                    greatestincrease = Cells(tickertablerow + 1, 11).Value
                    greatestincreaseticker = Cells(tickertablerow + 1, 9).Value
                
                ElseIf Cells(tickertablerow + 1, 11).Value < greatestdecrease Then
                
                    greatestdecrease = Cells(tickertablerow + 1, 11).Value
                    greatestdecreaseticker = Cells(tickertablerow + 1, 9).Value
                
                End If
                'Set greatest increase % and greatest decrease %
            
            End If
            
         Range("L" & tickertablerow + 1).Value = tickervolume
         'Enter Stock volume into table
         
         If Cells(tickertablerow + 1, 12).Value > greatestvolume Then
         
            greatestvolume = Cells(tickertablerow + 1, 12).Value
            greatestvolumeticker = Cells(tickertablerow + 1, 9).Value
         
         End If
         'Set greatest stock volume
        
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
            
        Range("P2").Value = greatestincreaseticker
        Range("Q2").Value = greatestincrease
        Range("P3").Value = greatestdecreaseticker
        Range("Q3").Value = greatestdecrease
        Range("P4").Value = greatestvolumeticker
        Range("Q4").Value = greatestvolume
        Range("Q2:Q3").NumberFormat = "0.00%"
        'Fill out advanced table

            
        Range("I1:L" & tickertablerow).Columns.AutoFit
        Range("O1:Q4").Columns.AutoFit
        'Adjust Column Width
        
Next ws

End Sub


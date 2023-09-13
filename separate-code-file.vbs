Sub Scriptv1()
Dim ws As Worksheet

For Each ws In Worksheets
'Define the variables
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim JCount As Double
    
    'Format the colums in the return infos to be currency or percentage
    ws.Range("J:J").NumberFormat = "$#,##0.00"
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("P2", "P3").NumberFormat = "0.00%"
    
    'Format columns in Yearly Change to be green if positive, red if negative
    ws.Range("J:J").FormatConditions.Add xlCellValue, xlLess, "0"
    ws.Range("J:J").FormatConditions(1).Interior.ColorIndex = 3
    ws.Range("J:J").FormatConditions.Add xlCellValue, xlGreater, "0"
    ws.Range("J:J").FormatConditions(2).Interior.ColorIndex = 4
    
    'Couldn't figure out how to not make J1 green, so instead of fighting I'll just change it back
    ws.Range("J1").FormatConditions.Delete
    
    'Set initial values
    TotalVolume = 0
    ReturnInfoRow = 2
    
    'Set headers
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    'Set the total number of rows
    Dim rowcount As Double
    rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Set the new ticker symbol in ticker symbol column, take in first volume, and store opening value
    For x = 2 To rowcount
        If ws.Cells(x, 1).Value <> ws.Cells(x - 1, 1).Value Then
            YearOpen = ws.Cells(x, 3).Value
            ws.Cells(ReturnInfoRow, 9).Value = ws.Cells(x, 1).Value
            TotalVolume = TotalVolume + ws.Cells(x, 7).Value
    
    'Check for ticker symbol change and add to total volume
        ElseIf ws.Cells(x, 1).Value = ws.Cells(x + 1, 1).Value Then
            TotalVolume = TotalVolume + ws.Cells(x, 7).Value
     'Once ticker symbol changes, obtain the last volume number, add to total, obtain final close value, calculate yearly and percent change, and print return values in columns
        ElseIf ws.Cells(x, 1).Value <> ws.Cells(x + 1, 1).Value Then
            TotalVolume = TotalVolume + ws.Cells(x, 7)
            ws.Cells(ReturnInfoRow, 12).Value = TotalVolume
            YearClose = ws.Cells(x, 6).Value
            YearChange = YearClose - YearOpen
            ws.Cells(ReturnInfoRow, 10).Value = YearChange
            PercentChange = YearChange / YearOpen
            ws.Cells(ReturnInfoRow, 11).Value = PercentChange
      'Reset TotalVolume and advance to next ReturnInfoRow for the new ticker. Shouldn't need to reset other values
            TotalVolume = 0
            YearOpen = 0
            PercentChange = 0
            ReturnInfoRow = ReturnInfoRow + 1
            End If
        Next x
        
        'Now that all of the values are in place, check for greatest %increase and place the value in the ticker
        'First, create a way to cycle through all of the rows for all 3 values, and create variables to store the highest value'
        Dim percentchecker As Double
        Dim greatincrease As Double
        Dim greatdecrease As Double
        Dim greatticker As String
        Dim LeastTicker As String
        Dim GreatVolume As Double
        Dim VolumeTicker As String
        
        percentchecker = ws.Cells(Rows.Count, "K").End(xlUp).Row
        volumechecker = ws.Cells(Rows.Count, "L").End(xlUp).Row
        
        greatincrease = 0
        greatdecrease = 0
        GreatVolume = 0
        
        'Check for both greatest increase and greatest decrease, do them both at the same time to save one cycle
        
        For y = 2 To percentchecker
        
            If ws.Cells(y, 11).Value > greatincrease Then
            'Store the value of the greatest increase
            greatincrease = ws.Cells(y, 11).Value
            'Store that corresponding ticker symbol
            greatticker = ws.Cells(y, 9).Value
            
            'Now do decrease
            ElseIf ws.Cells(y, 11).Value < greatdecrease Then
            'Store greatest decrease value
            greatdecrease = ws.Cells(y, 11).Value
            'Store corresponding ticker
            LeastTicker = ws.Cells(y, 9).Value
            
            End If
        Next y
        
    'Time to cycle through for the greatest volume
    For Z = 2 To volumechecker
        If ws.Cells(Z, 12).Value > GreatVolume Then
            'Store the value of the great volume
            GreatVolume = ws.Cells(Z, 12).Value
            'Store corresponding ticker
            VolumeTicker = ws.Cells(Z, 9).Value
        End If
        Next Z
        
    'Now print the values for all 3
    ws.Cells(2, 15).Value = greatticker
    ws.Cells(2, 16).Value = greatincrease
    ws.Cells(3, 15).Value = LeastTicker
    ws.Cells(3, 16).Value = greatdecrease
    ws.Cells(4, 15).Value = VolumeTicker
    ws.Cells(4, 16).Value = GreatVolume
    Next ws
            
End Sub
Attribute VB_Name = "Module11"
Sub calcStocks()
    
    Dim numSheets As Integer
    numSheets = ActiveWorkbook.Worksheets.Count
    
    Dim numRows As Long, row As Long
    numRows = ActiveSheet.UsedRange.Rows.Count

    Dim currentTicker As String
    'The Ticker on the line with row number = row
    Dim rowTicker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim runningTotalVolume As Double
    Dim summaryRow As Long
    Dim greatestPIncreaseTicker As String
    Dim greatestPIncreaseVal As Double
    Dim greatestPDecreaseTicker As String
    Dim greatestPDecreaseVal As Double
    Dim greatestTotalVolTicker As String
    Dim greatestTotalVol As Double
    
    'Iterate through each sheet
    For sheeti = 1 To numSheets
    Worksheets(sheeti).Select
    
        'Define Initial Values
        currentTicker = ""
        rowTicker = ""
        openingPrice = -1
        closingPrice = -1
        runningTotalVolume = 0
        summaryRow = 2
        greatestPIncreaseTicker = "none"
        greatestPIncreaseVal = 0
        greatestPDecreaseTicker = "none"
        greatestPDecreaseVal = 50000
        greatestTotalVolTicker = "none"
        greatestTotalVol = 0
    
        '---------------------------
        ' Write Summary Table Labels
        '---------------------------
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
    
        'Iterate through each row
        For row = 2 To numRows
        
            rowTicker = Cells(row, 1).Value
            
            'When this condition is true we are on the first line of a new stock
            If (currentTicker <> rowTicker) Then
                currentTicker = rowTicker
                openingPrice = Cells(row, 3).Value
                closingPrice = -1 'Set to nonsense value, so I always know
                runningTotalVolume = 0
            End If
            
            runningTotalVolume = runningTotalVolume + Int(Cells(row, 7).Value)
            
            'When this condition is true we are on the last line of the current stock
            If (Cells(row + 1, 1).Value <> rowTicker) Then
                    
                closingPrice = Cells(row, 6).Value
                    
                Dim percentChange As Double
                
                If (openingPrice <> 0) Then
                    percentChange = (closingPrice - openingPrice) / openingPrice
                Else
                    percentChange = 0
                End If
                    
                'Check if current percent change is the largest or smallest we've seen yet, and set the corresponding variables if so
                If percentChange > greatestPIncreaseVal Then
                    greatestPIncreaseVal = percentChange
                    greatestPIncreaseTicker = currentTicker
                End If
                
                If percentChange < greatestPDecreaseVal Then
                    greatestPDecreaseVal = percentChange
                    greatestPDecreaseTicker = currentTicker
                End If
                
                'Beucause this code only runs on the last line, the running total volume, and the current line's volume has already been added
                ' The runningTotalVolume = the total volume of the stock when this is run.
                If runningTotalVolume > greatestTotalVol Then
                    greatestTotalVol = runningTotalVolume
                    greatestTotalVolTicker = currentTicker
                End If
                
                '-------------------------------------------
                'Write out a row into the summary table
                '-------------------------------------------
                Cells(summaryRow, 9).Value = currentTicker
                Cells(summaryRow, 10).Value = closingPrice - openingPrice
                Cells(summaryRow, 11).Value = percentChange
                Cells(summaryRow, 12).Value = runningTotalVolume
                
                'Apply Color Formatting based on percent change
                If (percentChange) < 0 Then
                    Cells(summaryRow, 10).Interior.ColorIndex = 3
                    Cells(summaryRow, 11).Interior.ColorIndex = 3
                Else
                    Cells(summaryRow, 10).Interior.ColorIndex = 4
                    Cells(summaryRow, 11).Interior.ColorIndex = 4
                End If
                
                summaryRow = summaryRow + 1
            End If
            
        Next row
        
        'Write out summaries for some statistics across the entire sheet
        Cells(2, 14).Value = "Greatest Percent Increase"
        Cells(3, 14).Value = "Greatest Percent Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        Cells(2, 15).Value = greatestPIncreaseTicker
        Cells(2, 16).Value = greatestPIncreaseVal
        Cells(3, 15).Value = greatestPDecreaseTicker
        Cells(3, 16).Value = greatestPDecreaseVal
        Cells(4, 15).Value = greatestTotalVolTicker
        Cells(4, 16).Value = greatestTotalVol
        
        'Format cells to show correctly
        Range("K2:K" & numRows).NumberFormat = "0.00%"
        Cells(2, 16).NumberFormat = "0.00%"
        Cells(3, 16).NumberFormat = "0.00%"
        Cells().EntireColumn.AutoFit
        
    Next sheeti
End Sub


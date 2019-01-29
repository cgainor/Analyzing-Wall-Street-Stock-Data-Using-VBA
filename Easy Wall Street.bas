Attribute VB_Name = "Module1"
Sub EasyWallStreet()
' Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
' You will also need to display the ticker symbol to coincide with the total volume.

    ' Loop through all sheets
    For Each ws In Worksheets
    
        ' For each sheet, create results table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        
        ' Make variable tickerRow that will increment to fill in results table
        tickerRow = 1
        
        ' For each sheet, find what the last row number is and put it into lastRow variable
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' For each sheet, initialize totalStockVolume value to be 0
        totalStockVolume = 0
        
        ' For ticker column, want to list all unique ticker names
        ' Loops through all rows
        For i = 2 To lastrow
            
            ' Search for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Print the ticker name in a new row on the results table
                tickerRow = tickerRow + 1
                ws.Cells(tickerRow, 9).Value = ws.Cells(i, 1).Value
                
                ' Add the current stock volume to our stockVolumeTotal
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
                ' Print the totalStockVolume value in the Total Stock Volume column of the results table
                ws.Cells(tickerRow, 10).Value = totalStockVolume
                ' Reset the totalStockVolume to equal 0 for the next ticker name total
                totalStockVolume = 0
                
                Else
                ' Make variable to add total stock volume
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
            
End Sub

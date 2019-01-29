Attribute VB_Name = "Module3"
Sub HardWallStreet()
' Create a script that will loop through each year of stock data and grab the following:
        ' Ticker symbol
        ' Yearly change from what the stock opened the year at to what the closing price was
        ' The percent change from the opening at the beginning of the year, to the closing at the end
        ' The total amount of volume each stock had over the year
    ' Locate the stock with the "Greatest % Increase"
    ' Locate the stock with the "Greatest % Decrease"
    ' Locate the stock with the "Greatest Total Volume"

    ' Loop through all sheets
    For Each ws In Worksheets
    
        ' For each sheet, create results table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Make variable tickerRow that will increment to fill in results table
        tickerRow = 1
        
        ' For each sheet, find what the last row number is and put it into lastRow variable
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' For each sheet, initialize totalStockVolume value to be 0
        totalStockVolume = 0
        ' Initialize Open Price Value to be first open value
        Dim openPrice As Double
        openPrice = ws.Cells(2, 3).Value
        ' Initialize Close Price Value to be 0
        Dim closePrice As Double
        closePrice = 0
        ' Initialize Yearly Change Value to be 0
        Dim yearlyChange As Double
        yearlyChange = 0
        ' Initialize Percent Change Value to be 0
        Dim percentChange As Double
        percentChange = 0
        ' Initialize Greatest Percent Increase to be 0
        maxPerInc = 0
        ' Initialize Greatest Percent Decrease to be 0
        maxPerDec = 0
        ' Initialize Greatest Total Volume to be 0
        greatestTotalVolume = 0
        
        ' For ticker column, want to list all unique ticker names
        ' Loops through all rows
        For i = 2 To lastrow
            
            ' Search for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Print the ticker name in a new row on the results table
                tickerRow = tickerRow + 1
                ws.Cells(tickerRow, 9).Value = ws.Cells(i, 1).Value
                
                ' Add the current stock volume to our total
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value

                ' Print the totalStockVolume value in the Total Stock Volume column of the results table
                ws.Cells(tickerRow, 12).Value = totalStockVolume
                
                ' Set the close price
                closePrice = ws.Cells(i, 6).Value
                
                ' Print the yearlyChange value in the Yearly Change column of the results table
                yearlyChange = closePrice - openPrice
                ws.Cells(tickerRow, 10).Value = yearlyChange
                
                ' Set the fill color of yearlyChange to green if positive
                If yearlyChange >= 0 Then
                    ws.Cells(tickerRow, 10).Interior.ColorIndex = 4
                ' Set the fill color of yearlyChange to red if negative
                Else
                    ws.Cells(tickerRow, 10).Interior.ColorIndex = 3
                End If
                ' Print the percentChange value in the Percent Change column of the results table
                ' percentChange = (yearlyChange / openSum) * 100
                    ' Cover the case where open sum is 0
                    If openPrice = 0 Then
                        ws.Cells(tickerRow, 11).Value = Format(0, "percent")
                    Else
                        percentChange = (yearlyChange / openPrice)
                        ws.Cells(tickerRow, 11).Value = Format(percentChange, "percent")
                    End If
                
                ' Reset the totalStockVolume, yearlyChange, and percentChange to equal 0 for the next ticker name total
                totalStockVolume = 0
                yearlyChange = 0
                percentChange = 0
                
                ' Set new open price
                openPrice = ws.Cells(i + 1, 3).Value
                
                Else
                ' Make variable to add total stock volume
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
                
            End If
            
            ' Find Greatest Percent Increase amount
            If ws.Cells(i, 11).Value > maxPerInc Then
                maxPerInc = ws.Cells(i, 11).Value
            ' Initialize and find the stock name for maxPerInc
                Dim maxIncName As String
                maxIncName = ws.Cells(i, 9).Value
            End If
            
            ' Find Greatest Percent Decrease amount
            If ws.Cells(i, 11).Value < maxPerDec Then
                maxPerDec = ws.Cells(i, 11).Value
             ' Initialize and find the stock name for maxPerDec
                Dim maxDecName As String
                maxDecName = ws.Cells(i, 9).Value
             End If
             
            ' Find Greatest Total Volume amount
            If ws.Cells(i, 12).Value > greatestTotalVolume Then
                greatestTotalVolume = ws.Cells(i, 12).Value
             ' Initialize and find the stock name for greatestTotalVolume
                Dim greatestTVName As String
                greatestTVName = ws.Cells(i, 9).Value
             End If
            
        Next i
    
    ' Make second summary table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    ' Print the stock name/amount with the "Greatest % Increase" then reset value to 0 for next sheet
        ws.Range("Q2").Value = Format(maxPerInc, "Percent")
        ws.Range("P2").Value = maxIncName
        maxPerInc = 0
    ' Print the stock name/amount with the "Greatest % Decrease" then reset value to 0 for next sheet
        ws.Range("Q3").Value = Format(maxPerDec, "Percent")
        ws.Range("P3").Value = maxDecName
        maxPerDec = 0
    ' Print the stock name/amount with the "Greatest Total Volume" then reset value to 0 for next sheet
        ws.Range("Q4").Value = greatestTotalVolume
        ws.Range("P4").Value = greatestTVName
        greatestTotalVolume = 0
    
    'Automate the columns to fit the width of the values
        ws.Columns("A:Q").AutoFit
    Next ws
End Sub

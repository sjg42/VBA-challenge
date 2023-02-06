Attribute VB_Name = "Module1"
Sub VBA_Challenge():
'    Loop through all the stocks for one year and ouput the following information
'       The ticker symbol
'
'       Yearly change from the opening price at the beginning of a
'       given year to the closing price at the end of that year
'
'        The percentage change from the opening price at the beginning
'        of a given year to the closing price at the end of that year
'
'        The total stock volume of the stock.

'    Note: Must use Long data type instead of integer!!!


    For Each ws In Worksheets
    
    
        ws.Cells(1, 9) = "Ticker" 'Column I
        ws.Cells(1, 10) = "Yearly Change" 'Column J
        ws.Cells(1, 11) = "Percent Change" 'Column K
        ws.Cells(1, 12) = "Total Stock Volume" 'Column L
        
        'first, find the last row in the sheet
        Dim finalRow As Long
        finalRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'check on the ticker name
        Dim ticker As String
        
        'make variable for our total stock volume
        Dim totalSV As Double
        totalSV = 0
        
        'variable to hold the rows in the totals columns I-L
        Dim tickerRow As Long
        tickerRow = 2 'first row to populate in columns I-L
        
        'loop through the rows and check the changes in the credit cards
        Dim row As Long
        
        'track first row of new ticker
        Dim firstRow As Long
        firstRow = 2 'set to 2 which is the first row where a ticker shows up
        
        'track last row of new ticker
        Dim lastRow As Long
        lastRow = 0
        
        For row = 2 To finalRow
            'add to ticker total vol from column G (column 7)
            totalSV = totalSV + ws.Cells(row, 7).Value
            
            'check the changes in the ticker
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                'set the ticker name
                ticker = ws.Cells(row, 1).Value ' grabs the value from column A BEFORE the change
                
                'display the ticker name on the current row of the the ticker column
                ws.Cells(tickerRow, 9).Value = ticker
                
                'display the total stock volume on the current row of the
                'tickerRow in column L (column 12)
                ws.Cells(tickerRow, 12).Value = totalSV
                            
                'set lastRow as the current row
                lastRow = row
                
                'Set opening and closing price
                Dim openPrice As Double
                openPrice = ws.Cells(firstRow, 3).Value
                
                Dim closePrice As Double
                closePrice = ws.Cells(lastRow, 6)
                
                'Calculate Yearly Change
                Dim YearlyChange As Double
                YearlyChange = Round(openPrice - closePrice, 2)
                ws.Cells(tickerRow, 10).Value = YearlyChange
                
                'Color Cell According to Yearly Change (Zero gets no color change)
                'MsgBox ("Color Time! The yearly change is" + Str(YearlyChange))
                If YearlyChange < 0 Then
                    ws.Cells(tickerRow, 10).Interior.ColorIndex = 3
                ElseIf YearlyChange > 0 Then
                    ws.Cells(tickerRow, 10).Interior.ColorIndex = 4
                End If
                
                'Calculate Percent Change
                Dim percentChange As Double
                If openPrice = 0 Then
                    percentChange = closePrice
                Else
                    percentChange = (openPrice - closePrice) / openPrice
                End If
                
                'Put percentChange in percent format at desired location
                
                ws.Cells(tickerRow, 11).Value = FormatPercent(Str(percentChange))
                
                'set the next first row
                firstRow = row + 1
                
                'Move to next credit card row
                tickerRow = tickerRow + 1
                
                'reset credit card total for next credit card
                totalSV = 0
                
            End If
                
        Next row
        'Make set up for putting greatest percent changes and volume change
        ws.Cells(1, 16) = "Ticker" 'Column P
        ws.Cells(1, 17) = "Value" 'Column Q
        ws.Cells(2, 15) = "Greatest % Increase" 'Column O
        ws.Cells(3, 15) = "Greatest % Increase" 'Column O
        ws.Cells(4, 15) = "Greatest Total Volume" 'Column O
        
        
        
        'initialize variables for Greatest Percent Increase(GPI), Greatest Percent Decrease (GPD)
        'and Greatest Total Volume (GTV)
        'Must be double to hold entire decimal
        Dim GPI, GPD, GTV As Double
        
        'Use Max and Min workbook functions to find the GPI, GPD, and GTV
        GPI = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(2, 17).Value = FormatPercent(Str(GPI))
        
        GPD = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(3, 17).Value = FormatPercent(Str(GPD))
        
        GTV = WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(4, 17).Value = GTV
        
        
        'Use Match workbook funciton to find Tickers of the respective values
        Dim matchGPI, matchGPD, matchGTV As Long
        matchGPI = WorksheetFunction.Match(GPI, ws.Range("K:K"), 0)
        ws.Cells(2, 16).Value = ws.Cells(matchGPI, 9).Value
        
        matchGPD = WorksheetFunction.Match(GPD, ws.Range("K:K"), 0)
        ws.Cells(3, 16).Value = ws.Cells(matchGPD, 9).Value
        
        matchGTV = WorksheetFunction.Match(GTV, ws.Range("L:L"), 0)
        ws.Cells(4, 16).Value = ws.Cells(matchGTV, 9).Value
        
        'Autofit Columns
        ws.Range("A:Q").Columns.AutoFit
        
    Next ws
End Sub

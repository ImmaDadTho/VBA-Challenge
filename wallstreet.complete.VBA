Attribute VB_Name = "Module111"
Sub wallstreet():
    For Each ws In Worksheets
            Dim ticker As String, volume, greatest_volume, closeprice, greatest_increase, greatest_decrease, ticker_index As Double
            
            'variable that will declare what row the data will begin to display on
            summarytable = 2
            
            'variable to hold the openening year price for first ticker symbol
            openprice = ws.Cells(2, 3).Value
    
            'finding the last row in the sheet
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            'set the headers for all worksheets and format them to fit in the columns
            ws.Range("I1").Value = "Ticker"
                ws.Range("I1").Columns.AutoFit
            ws.Range("J1").Value = "Yearly Change"
                ws.Range("J1").Columns.AutoFit
            ws.Range("K1").Value = "Percent Change"
                ws.Range("K1").Columns.AutoFit
            ws.Range("L1").Value = "Total Stock Volume"
                ws.Range("L1").Columns.AutoFit
                 
        'loop from row 2 to last row in column A
        For Row = 2 To lastrow
             
            'see if the value changes, if it does go back up one and capture the symbol
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
                'variable to hold ticker symbols
                ticker = ws.Range("A" & Row).Value
                
                'add the symbols to the "ticker" column
                ws.Range("I" & summarytable).Value = ticker
                    
                'variable to capture the total volume
                volume = volume + ws.Range("G" & Row).Value
                
                'add the total volume per stock symbol to the column "total stock volume"
                ws.Range("L" & summarytable).Value = volume
                    
                'variable to capture the last prices of the year
                closeprice = ws.Range("F" & Row).Value
                        
                'calculate yearly change into column
                ws.Range("J" & summarytable).Value = closeprice - openprice
                
                'variable to hold results from yearlychangecolumn
                yearlychange = ws.Range("j" & summarytable).Value
                
                    'conditional formatting of "Yearlychange" column
                    If ws.Range("J" & summarytable).Value > 0 Then
                        ws.Range("J" & summarytable).Interior.ColorIndex = 4
                    ElseIf ws.Range("J" & summarytable).Value < 0 Then
                        ws.Range("J" & summarytable).Interior.ColorIndex = 3
                    Else: ws.Range("J" & summarytable).Interior.ColorIndex = 15
                    
                    End If
                
                'calculate the percent change
                ws.Range("k" & summarytable).Value = (yearlychange / openprice)
                
                'format results into percantage
                ws.Range("k" & summarytable).NumberFormat = "0.00%"
                
                'variable to capture yearly opening date price starting from row second ticker symbol
                openprice = ws.Range("C" & Row + 1).Value
                        
    
                summarytable = summarytable + 1
                
                'reset the totals to 0
                volume = 0
            
            Else
                 
                'continue adding the totals until the stock symbol change
                volume = volume + ws.Range("G" & Row).Value
            
            End If
        Next Row
            'functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
            Dim percentmax, percentmin, volumemax As LongLong, tickerindex As Integer
            
            'variables to hold  each value
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            
            'formatting each cell to fit values
            ws.Range("O:Q").Columns.AutoFit
            
            'finding the last row of the summarytable where we will look for the values in
            lastrow_bonus = Cells(Rows.Count, 9).End(xlUp).Row
            
            'locating and displaying the greatest percent increase and the ticker
            percentmax = WorksheetFunction.max(ws.Range("K2:K" & lastrow_bonus))
            tickerindex = WorksheetFunction.Match(percentmax, ws.Range("K2:K" & lastrow_bonus), 0)
            ws.Range("P2").Value = ws.Range("I" & tickerindex + 1).Value
            ws.Range("Q2").Value = percentmax
            ws.Range("Q2").NumberFormat = "0.00%"
            
            'locating and displaying the greatest percent decrease and the ticker
            percentmin = WorksheetFunction.Min(ws.Range("K2:K" & lastrow_bonus))
            tickerindex = WorksheetFunction.Match(percentmin, ws.Range("K2:K" & lastrow_bonus), 0)
            ws.Range("P3").Value = ws.Range("I" & tickerindex + 1).Value
            ws.Range("Q3").Value = percentmin
            ws.Range("Q3").NumberFormat = "0.00%"
            
            ''locating and displaying the total greatest volume
            volumemax = WorksheetFunction.max(ws.Range("L2:L" & lastrow_bonus))
            tickerindex = WorksheetFunction.Match(volumemax, ws.Range("L2:L" & lastrow_bonus), 0)
            ws.Range("P4").Value = ws.Range("I" & tickerindex + 1).Value
            ws.Range("Q4").Value = volumemax
           
    Next ws
End Sub


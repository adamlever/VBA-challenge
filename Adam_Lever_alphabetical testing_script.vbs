Attribute VB_Name = "Module1"
Sub alphabeticaltesting()

'Loop through all sheets
For Each ws In Worksheets

    'Set Variables
    Dim TickerSymbol As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim SummaryTableRow As Integer
    Dim lastRow As Double
    Dim Max As Double
    Dim Min As Double
    Dim GreatestVolume As Double
    
    'Create First Summary Table Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change ($)"
    ws.Cells(1, 11).Value = "Percent Change (%)"
    ws.Cells(1, 12).Value = "Total Stock Volume"
             
    'Set Initial Summary Table Row Value
    SummaryTableRow = 2
    
    'Set Initial TotalVolume
    TotalVolume = 0
            
    'Set Last Row with Data
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
    'Loop Through Ticker Column
    For i = 2 To lastRow
    
        'When the current row and the next row are not the same
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            
            'Set the TickerSymbol
            TickerSymbol = ws.Cells(i, 1).Value

            'Print Ticker Symbol in Summary Table
            ws.Cells(SummaryTableRow, 9).Value = TickerSymbol
            
            'Set the Closing Price
            ClosingPrice = ws.Cells(i, 6).Value
            
            'Calculate Yearly Change in Price For Each Ticker
            YearlyChange = ClosingPrice - OpeningPrice
            
            'Print Yearly Change in Summary Table
            ws.Cells(SummaryTableRow, 10).Value = YearlyChange
                        
                'Change Colour of Yearly Change value cells to green for positive change or red for negative change
                If ws.Cells(SummaryTableRow, 10).Value > 0 Then
                    ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                End If
                
           'Calculate Yearly Change of Price in % For Each Ticker
            If YearlyChange = 0 Then
                PercentChange = 0
            ElseIf OpeningPrice = 0 Then
                PercentChange = 0
            ElseIf ClosingPrice = 0 Then
                PercentChange = 0
            Else
                PercentChange = YearlyChange / OpeningPrice
            End If
            
            'Print Percent Change of Closing Pricse from Opening Prices in Summary Table
            ws.Cells(SummaryTableRow, 11).Value = PercentChange
            
            'Change cell format of Percent Change to Percentage
            ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
            
                'Change Colour of Percent Change value cells to green for positive change or red for negative change
                If ws.Cells(SummaryTableRow, 11).Value > 0 Then
                    ws.Cells(SummaryTableRow, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(SummaryTableRow, 11).Interior.ColorIndex = 3
                End If
            
            'Add volume of Tickers last row to TotalVolume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            'Print Total Volume in Summary Table
            ws.Cells(SummaryTableRow, 12).Value = TotalVolume
                
            ' Add one to the Summary Table Row
            SummaryTableRow = SummaryTableRow + 1
                
            'Reset TotalVolume
            TotalVolume = 0
                
        Else
            'When the current row and the previous row are not the same
            If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
            
            'Find Each Tickers Yearly Opening Price
            OpeningPrice = ws.Cells(i, 3).Value
            End If
                       
            'When the current row and the next row are the same
            If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
            
            'Add to the Total Volume Tally
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
                       
        End If

    'Next Row in Loop
    Next i
           
     
    'Create Second Summary Table Headers
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
    'Find Greatest Percentage Increase
    MaxPercent = WorksheetFunction.Max(ws.Columns(11))
    
    'Print Greatest Percentage Increase in Second Summary Table
    ws.Cells(2, 17).Value = MaxPercent
    
    'Change Cell Format of Greatest Percentage Increase to Percent
    ws.Cells(2, 17).NumberFormat = "0.00%"
 
 
    'Find Greatest Percentage Decrease
    MinPercent = WorksheetFunction.Min(ws.Columns(11))
    
    'Print Greatest Decrease Percentage in Second Summary Table
    ws.Cells(3, 17).Value = MinPercent
    
    'Change Cell Format of Greatest Percentage Decrease to Percent
    ws.Cells(3, 17).NumberFormat = "0.00%"
 
 
    'Find Greatest Total Volume
    GreatestVolume = WorksheetFunction.Max(ws.Columns(12))
    
    'Print Greatest Total Volume in Second Summary Table
    ws.Cells(4, 17).Value = GreatestVolume
    
    
    'Set Last Row of data in Column 11 of First Summary Table
    LastSummaryRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    
    'Loop Through First Summary Table to Retrieve Ticker Symbols for the Greatest Percentage Increase, Decrease and Total Volume
    For i = 2 To LastSummaryRow
        If ws.Cells(i, 11) = MaxPercent Then
            TickerMax = ws.Cells(i, 9)
        End If
    
        If ws.Cells(i, 11) = MinPercent Then
            TickerMin = ws.Cells(i, 9)
        End If
    
        If ws.Cells(i, 12) = GreatestVolume Then
            TickerGreatestVolume = ws.Cells(i, 9)
        End If
    
    'Next Row in Loop
    Next i
    
    
    'Print Ticker of Greatest Percentage Increase, Decrease and Total Volume in Second Summary Table
    ws.Cells(2, 16).Value = TickerMax
    ws.Cells(3, 16).Value = TickerMin
    ws.Cells(4, 16).Value = TickerGreatestVolume

    
    'Autofit columns to display data
    ws.Columns("A:Q").AutoFit


'Go to next Worksheet to Repeat Sub
Next ws


End Sub

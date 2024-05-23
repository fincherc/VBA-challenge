Attribute VB_Name = "Module1"

Sub alpha_beta()
    'Ensure this does not get ruined
    
    'Get the Worksheet Name
    Dim WorksheetName As String
    
    'Need the rows
    Dim StartRow As Long
    Dim CurrentRow As Long
    Dim EndRow As Long
    
    'Total Stock
    Dim TotalStockVolume As Variant
    
    'Create the ticker
    Dim LastTicker As String
    Dim CurrentTicker As String
    
    'Conditional Formatting
    Dim condition1 As FormatCondition
    Dim condition2 As FormatCondition
    
    '----------------
    'Second Part
    '----------------
    
    'Greatest Portion
    Dim greatest_ticker As String
    
    Dim ticker_percent As Double
    Dim greatest_percent As Double
    Dim ticker_greatest_total As Variant
    Dim ticker_total As Variant
    
    For Each ws In Worksheets

        ' Grabbed the WorksheetName and LastRow
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Setup the Columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Setup the for loop to check the range of items, get the entry that doesn't exist
        For CurrentRow = 2 To LastRow
        
            'Get the ticker
            CurrentTicker = ws.Cells(CurrentRow, 1).Value
            
            'We have the ticker and the range, get the last Ticker row
            TickerLastRow = ws.Range("I" & Rows.Count).End(xlUp).Row
            
            'Grab ticker entry in other column
            LastTicker = ws.Cells(TickerLastRow, 9).Value
            
            'If ticker and the last entry are not the same, add it to the entry
            If (CurrentTicker <> LastTicker) Then
                
                TotalStockVolume = 0
                
                ' --------------------------------------------
                ' Ticker
                ' --------------------------------------------
                
                'Loop through the range to check if the entry changes
                For StartRow = CurrentRow To LastRow
                    If (CurrentTicker = ws.Cells(StartRow, 1).Value) Then
                        TotalStockVolume = TotalStockVolume + ws.Cells(StartRow, 7).Value
                    End If
                    
                    If (CurrentTicker <> ws.Cells(StartRow, 1).Value Or StartRow = LastRow) Then
                        EndRow = StartRow
                        Exit For
                    End If
                Next StartRow
                
                'Enter the ticker into the cell
                ws.Cells(TickerLastRow + 1, 9).Value = CurrentTicker
                
                ' --------------------------------------------
                ' Quarterly Change
                ' from the opening price at the beginning of a
                ' given quarter to the closing price at the end of that quarter
                ' --------------------------------------------
                
                'Enter the Percentage Change into the cell
                'Close (end) minus Open (start)/Close
                'Format to Percentage
                ws.Cells(TickerLastRow + 1, 10).Value = (ws.Cells(EndRow - 1, 6) - ws.Cells(CurrentRow, 3))
                ws.Cells(TickerLastRow + 1, 10).NumberFormat = "0.00"
                
                ' --------------------------------------------
                ' Percentage Change
                ' from the opening price at the beginning of a
                ' given quarter to the closing price at the end of that quarter
                ' --------------------------------------------
                
                'Enter the Percentage Change into the cell
                'Close (end) minus Open (start)/Close
                'Format to Percentage
                ws.Cells(TickerLastRow + 1, 11).Value = (ws.Cells(EndRow - 1, 6) - ws.Cells(CurrentRow, 3)) / ws.Cells(CurrentRow, 3)
                ws.Cells(TickerLastRow + 1, 11).NumberFormat = "0.00%"
                
                ' --------------------------------------------
                ' Total Stock Volume
                ' The total stock volume of the stock
                ' --------------------------------------------
                ws.Cells(TickerLastRow + 1, 12).Value = TotalStockVolume
                
            End If
        Next CurrentRow

        '-------------------------
        ' Finished with the main code, perform conditional formatting and proceed with the Greatest sections
        '-------------------------
        
        ' Add conditional formatting to the Quarterly Change
        Set condition1 = ws.Range("J2:J" & LastRow).FormatConditions.Add(xlCellValue, xlGreater, "=0")
        Set condition2 = ws.Range("J2:J" & LastRow).FormatConditions.Add(xlCellValue, xlLess, "=0")
        
        condition1.Interior.ColorIndex = 4
        condition2.Interior.ColorIndex = 3
        
        ' Grab the LastRow and StartRow
        LastRow = ws.Range("K" & Rows.Count).End(xlUp).Row
        StartRow = 2
        
        '-------------------------
        ' Greatest % Increase
        '-------------------------
        
        ' Establish our "greatest" one before looping
        greatest_ticker = ws.Cells(StartRow, 9).Value
        greatest_percent = ws.Cells(StartRow, 11).Value
        
        'For Loop, start at 3 to check previous entry - greatest percent
        For CurrentRow = 3 To LastRow
        
            'Get the ticker
            CurrentTicker = ws.Cells(CurrentRow, 9).Value
            ticker_percent = ws.Cells(CurrentRow, 11).Value
            
            If (greatest_percent < ticker_percent) Then
                greatest_ticker = CurrentTicker
                greatest_percent = ticker_percent
                
            'Continue to the next iteration
            End If
        Next CurrentRow
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = greatest_ticker
        ws.Cells(2, 16).Value = greatest_percent
        ws.Cells(2, 16).NumberFormat = "0.00%"
        
        '------------------------
        ' Greatest % Decrease
        '------------------------
        
        ' Establish our "greatest" one before looping
        greatest_ticker = ws.Cells(StartRow, 9).Value
        greatest_percent = ws.Cells(StartRow, 11).Value
        
        'For Loop, start at 3 to check previous entry
        For CurrentRow = 3 To LastRow
        
            'Get the ticker
            CurrentTicker = ws.Cells(CurrentRow, 9).Value
            ticker_percent = ws.Cells(CurrentRow, 11).Value
            
            If (greatest_percent > ticker_percent) Then
                greatest_ticker = CurrentTicker
                greatest_percent = ticker_percent
                
            'Continue to the next iteration
            End If
        Next CurrentRow
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = greatest_ticker
        ws.Cells(3, 16).Value = greatest_percent
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
        '--------------------------
        ' Greatest Total Volume
        '--------------------------
        
        ' Establish our "greatest" one before looping
        greatest_ticker = ws.Cells(StartRow, 9).Value
        ticker_greatest_total = ws.Cells(StartRow, 12).Value
        
        'For Loop, start at 3 to check previous entry
        For CurrentRow = 3 To LastRow
        
            'Get the ticker
            CurrentTicker = ws.Cells(CurrentRow, 9).Value
            ticker_total = ws.Cells(CurrentRow, 12).Value
            
            If (ticker_greatest_total < ticker_total) Then
                greatest_ticker = CurrentTicker
                ticker_greatest_total = ticker_total
                
            'Continue to the next iteration
            End If
        Next CurrentRow
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = greatest_ticker
        ws.Cells(4, 16).Value = ticker_greatest_total
    Next ws
    
End Sub




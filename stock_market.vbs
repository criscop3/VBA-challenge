Sub stock_market()

    Dim i, j, last_record As Long
    Dim ticker_row As Integer
    Dim ticker_name As String
    Dim total_stock As Double
    Dim yearly_change, percent_change, stock_open, stock_close, count As Double
    Dim ws As Worksheet

'Loops through the worksheets
For Each ws In Worksheets

    ws.Activate
    last_record = Cells(Rows.count, 1).End(xlUp).Row
    last_summary_row = Cells(Rows.count, 9).End(xlUp).Row
    ticker_row = 2
    total_stock = 0
    count = 0
    
    'Loop through initial data set in active worksheet

    For i = 2 To last_record
        count = count + 1
        stock_open = Cells(i - (count - 1), 3).Value
        stock_close = Cells(i, 6).Value
        
        'Checks for a change in ticker and calculates for that ticker
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ticker_name = Cells(i, 1).Value
            yearly_change = stock_close - stock_open
            percent_change = FormatPercent((stock_close - stock_open) / stock_open, 2)
            total_stock = total_stock + Cells(i, 7)
            
            'Populates summary section
            
            Range("I" & ticker_row).Value = ticker_name
            Range("J" & ticker_row).Value = yearly_change
                If Range("J" & ticker_row).Value < 0 Then
                    Range("J" & ticker_row).Interior.ColorIndex = 3
                ElseIf Range("J" & ticker_row).Value > 0 Then
                    Range("J" & ticker_row).Interior.ColorIndex = 4
                Else
                    Range("J" & ticker_row).Interior.ColorIndex = 2
                End If
            Range("K" & ticker_row).Value = percent_change
                If Range("K" & ticker_row).Value < 0 Then
                    Range("K" & ticker_row).Interior.ColorIndex = 3
                ElseIf Range("K" & ticker_row).Value > 0 Then
                    Range("K" & ticker_row).Interior.ColorIndex = 4
                Else
                    Range("K" & ticker_row).Interior.ColorIndex = 2
                End If
                   
            Range("L" & ticker_row).Value = total_stock
            
            'Moves down a row in summary for next ticker calculation
            
            ticker_row = ticker_row + 1
            
            'Resets count and stock calculation
            
            count = 0
            total_stock = 0
        Else
            total_stock = total_stock + Cells(i, 7)
        End If
        
    Next i
    
    'Loops through summary section
    
    max_stock = WorksheetFunction.Max(Range("L2:L" & last_summary_row))
    max_percent = WorksheetFunction.Max(Range("K2:K" & last_summary_row))
    min_percent = WorksheetFunction.Min(Range("K2:K" & last_summary_row))
    
   For j = 2 To last_summary_row
        If Cells(j, 12).Value = max_stock Then
            Range("R4").Value = Cells(j, 12).Value
            Range("Q4").Value = Cells(j, 9).Value
        End If
          
        If Cells(j, 11).Value = min_percent Then
            Range("R3").Value = FormatPercent(Cells(j, 11).Value, 2)
            Range("Q3").Value = Cells(j, 9).Value
        End If
        
        If Cells(j, 11).Value = max_percent Then
            Range("R2").Value = FormatPercent(Cells(j, 11).Value, 2)
            Range("Q2").Value = Cells(j, 9).Value
        End If
         
   Next j

Next ws
    
End Sub

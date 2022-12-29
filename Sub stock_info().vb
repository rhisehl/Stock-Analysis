Sub stock_info()

Dim ws As Worksheet

'loop through worksheets'
For Each ws In ThisWorkbook.Worksheets
    
    'Dim summary_row, r As Integer
    r = 1
    summary_row = 1
    
   
    
    Dim start_stock, end_stock, yr_change, per_change, volume, max_inc, max_dec, max_vol As Double
    start_stock = 0#
    volume = 0#
    max_inc = 0#
    max_dec = 0#
    max_vol = 0#
    
    
    'loop through row A until empty'
    Do Until IsEmpty(ws.Cells(r, 1))
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            If ws.Cells(r, 1) = "<ticker>" Then
                start_stock = ws.Cells(r + 1, 3).Value
                summary_row = summary_row + 1
            Else
                ws.Cells(summary_row, 9).Value = ws.Cells(r, 1).Value
                ws.Cells(summary_row, 12).Value = Round(volume, 2)
                end_stock = ws.Cells(r, 6).Value
                yr_change = end_stock - start_stock
                per_change = yr_change / start_stock
                ws.Cells(summary_row, 10).Value = Round(yr_change, 2)
                If ws.Cells(summary_row, 10).Value > 0 Then
                    ws.Cells(summary_row, 10).Interior.Color = vbGreen
                ElseIf ws.Cells(summary_row, 10).Value < 0 Then
                    ws.Cells(summary_row, 10).Interior.Color = vbRed
                End If
                ws.Cells(summary_row, 11).Value = FormatPercent(per_change)
                start_stock = ws.Cells(r + 1, 3).Value
                volume = volume + ws.Cells(r, 7).Value
                If per_change > max_inc Then
                    max_inc = per_change
                    ws.Cells(2, 17).Value = FormatPercent(max_inc)
                    ws.Cells(2, 16).Value = ws.Cells(summary_row, 9).Value
                End If
                If per_change < max_dec Then
                    max_dec = per_change
                    ws.Cells(3, 17).Value = FormatPercent(max_dec)
                    ws.Cells(3, 16).Value = ws.Cells(summary_row, 9).Value
                End If
                If volume > max_vol Then
                    max_vol = volume
                    ws.Cells(4, 17).Value = max_vol
                    ws.Cells(4, 16).Value = ws.Cells(summary_row, 9).Value
                 End If
                
                summary_row = summary_row + 1
                volume = 0
            End If
        Else
            volume = volume + ws.Cells(r, 7).Value
        End If
        r = r + 1
    Loop
    
    'autofit the current sheet'
    ws.Range("a1:f1").ColumnWidth = 9
    ws.Range("g1:i1").ColumnWidth = 11.5
    ws.Range("j1:k1").ColumnWidth = 14.5
    ws.Range("L1").ColumnWidth = 18
    ws.Range("o1").ColumnWidth = 21
    ws.Range("P1:Q1").ColumnWidth = 8
    ws.Range("J1").Interior.ColorIndex = 0
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
Next ws
End Sub


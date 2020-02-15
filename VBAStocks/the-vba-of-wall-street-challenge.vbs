Sub summary()

For Each ws In Worksheets

    ' Insert table headings and row labels
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    ' Find the last row in the table summarizing the yearly values
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Set range from which to detrmine maximum and minimum values
    Dim rng_p As Range
        Set rng_p = ws.Range("K2:K" & lastrow)
    
    Dim rng_v As Range
        Set rng_v = ws.Range("L2:L" & lastrow)
    
    ' Create variables for maximum and minimum values
    Dim MaxPercent As Double
        MaxPercent = Application.WorksheetFunction.Max(rng_p)
    
    Dim MinPercent As Double
        MinPercent = Application.WorksheetFunction.Min(rng_p)
    
    Dim MaxVolume As LongLong
        MaxVolume = Application.WorksheetFunction.Max(rng_v)
    
    ' Find the maximum and minimum values and add to summary table
    For i = 2 To lastrow
    
        If (ws.Cells(i, 11).Value = MaxPercent) Then
            ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
        
        End If
        
        If (ws.Cells(i, 11).Value = MinPercent) Then
            ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
        
       End If
        
        If (ws.Cells(i, 12).Value = MaxVolume) Then
            ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
        
        End If
    
    Next i
    
    ' Format greatest % increase/decreaes as percentages in summary table
    ws.Range("P2:P3").Style = "Percent"
    
Next ws

End Sub

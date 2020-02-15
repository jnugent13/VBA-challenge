Sub testing()

For Each ws In Worksheets

    ' Insert headings for summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ' Create variable to track ticker name
    Dim ticker As String

    ' Create variable to track place in summary table
    Dim SummaryTableRow As Integer
      SummaryTableRow = 2
  
    ' Create variable for starting price
    Dim openPrice As Double
      openPrice = 0

    ' Create variable for ending price
    Dim closePrice As Double
      closePrice = 0

    ' Create variable to track stock volume
    Dim volume As LongLong
      volume = 0

    ' Find the last row on the sheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through each row
    For i = 2 To lastrow

    ' If ticker doesn't match one before, set the opening price
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
      openPrice = ws.Cells(i, 3).Value
    
    ' Search if we're still using the same ticker. If not...
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ' Set the closing price
        closePrice = ws.Cells(i, 6).Value
        
        ' Add ticker symbol to table
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & SummaryTableRow).Value = ticker
        
        ' Add price change to table
        ws.Range("J" & SummaryTableRow).Value = closePrice - openPrice
 
        ' Add percent change to table
        If openPrice > 0 Then
            ws.Range("K" & SummaryTableRow).Value = closePrice / openPrice - 1
        Else
            ws.Range("K" & SummaryTableRow).Value = 0
        End If
        
        ' Add stock volume to table
        volume = volume + ws.Cells(i, 7).Value
        ws.Range("L" & SummaryTableRow).Value = volume
        
        ' Add one to the summary table row
        SummaryTableRow = SummaryTableRow + 1
        
        ' Reset price change
        priceChange = 0
        
        ' Reset stock volume
        volume = 0
        
    ' If cell immediately following a row is the same ticker
    Else
        
        ' Add stock volume to previous stock volume
        volume = volume + ws.Cells(i, 7).Value
    
    End If
    
Next i
    
    For i = 2 To lastrow
    ' Format percentage change column as percent
    ws.Cells(i, 11).Style = "Percent"
    
    ' Format the cells to highlight positive changes in green and negative in red
    If (ws.Cells(i, 10).Value > 0) Then
      ws.Cells(i, 10).Interior.ColorIndex = 4
      
    ElseIf (ws.Cells(i, 10).Value < 0) Then
      ws.Cells(i, 10).Interior.ColorIndex = 3

    End If
    
Next i
Next ws

End Sub


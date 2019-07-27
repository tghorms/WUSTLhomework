Sub TickerTest()

'Perform action across all worksheets
For Each ws In Worksheets

    'Set Variables we want to extract
    Dim Ticker As String

    Dim Stock_Volume As Double

    Stock_Volume = 0
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Assigning Titles to Columns
    ws.Cells(1, 9) = "Ticker"
    
    ws.Cells(1, 10) = "Total Stock Volume"
    
    'Location for each diferrent Ticker
    Dim Summary_Table_Row As Integer
    
    Summary_Table_Row = 2
    
    'Collecting all different Tickers
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Set the Ticker
        Ticker = ws.Cells(i, 1).Value
        
        'Add to The Stock Volume Total
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    
        'Print the Ticker In the Summary column
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        'Print  the Stock Volume in Summary Column
        ws.Range("J" & Summary_Table_Row).Value = Stock_Volume
    
        Summary_Table_Row = Summary_Table_Row + 1
    
        'Reset Stock Volume
        Stock_Volume = 0
    
    Else
    
        'Add Stock Volume Total
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    
    End If

    Next i

Next ws
    
End Sub


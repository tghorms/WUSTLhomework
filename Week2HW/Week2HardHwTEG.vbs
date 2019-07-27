Sub TickerTest()

Dim ws As Worksheet

'Perform action across all worksheets
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
    'Set Variables we want to extract
        Dim Ticker As String
        Dim Stock_Volume As Double
        Dim Open_Price As Double
        Dim Closing_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        
        Stock_Volume = 0
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Assigning Titles to Columns
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Change"
    
    'Location for each diferrent Ticker
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        Dim Column As Double
        Column = 1
        
    'Set Opening Price
        Open_Price = ws.Cells(2, Column + 2).Value
        
    'Collecting all different Tickers
        Dim i As Long
    
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Set the Ticker
            Ticker = ws.Cells(i, 1).Value
            
           'Set Closing Price
            Closing_Price = ws.Cells(i, Column + 5).Value
            
            'Add Yearly Change
            Yearly_Change = Closing_Price - Open_Price
            ws.Cells(Summary_Table_Row, Column + 9).Value = Yearly_Change
            
            'Add Percent Change
                    If (Open_Price = 0 And Closing_Price = 0) Then
                        Percent_Change = 0
                    ElseIf (Open_Price = 0 And Closing_Price <> 0) Then
                        Pecernt_Change = 1
                    Else
                        Percent_Change = Yearly_Change / Open_Price
                        ws.Cells(Summary_Table_Row, Column + 10).Value = Percent_Change
                        ws.Cells(Summary_Table_Row, Column + 10).NumberFormat = "0.00%"
                    End If
            
                'Add to The Stock Volume Total
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                 'Print the Ticker In the Summary column
                    ws.Range("I" & Summary_Table_Row).Value = Ticker
                'Print  the Stock Volume in Summary Column
                    ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                    Summary_Table_Row = Summary_Table_Row + 1
                    Open_Price = ws.Cells(i + 1, Column + 2)
    
                'Reset Stock Volume
                    Stock_Volume = 0
    
            Else
             'Add Stock Volume Total
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    
            End If
    Next i
    
    'Finding the last Row of Yearly Change
    LastRowYC = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row
    
    'Set Colors
    For j = 2 To LastRowYC
        If (ws.Cells(j, Column + 9).Value > 0 Or ws.Cells(j, Column + 9).Value = 0) Then
            ws.Cells(j, Column + 9).Interior.ColorIndex = 4
        ElseIf ws.Cells(j, Column + 9).Value < 0 Then
            ws.Cells(j, Column + 9).Interior.ColorIndex = 3
        End If
    Next j
    
    'Hard Summary Table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Loop through Rows to find greatest value and ticker
    For K = 2 To LastRowYC
        If ws.Cells(K, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRowYC)) Then
            ws.Cells(2, Column + 15).Value = ws.Cells(K, Column + 8).Value
            ws.Cells(2, Column + 16).Value = ws.Cells(K, Column + 10).Value
            ws.Cells(2, Column + 16).NumberFormat = "0.00%"
        
        ElseIf ws.Cells(K, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRowYC)) Then
            ws.Cells(3, Column + 15).Value = ws.Cells(K, Column + 8).Value
            ws.Cells(3, Column + 16).Value = ws.Cells(K, Column + 10).Value
            ws.Cells(3, Column + 16).NumberFormat = "0.00%"
            
        ElseIf ws.Cells(K, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRowYC)) Then
            ws.Cells(4, Column + 15).Value = ws.Cells(K, Column + 8).Value
            ws.Cells(4, Column + 16).Value = ws.Cells(K, Column + 11).Value
        End If
        
    Next K
    
Next ws
    
End Sub





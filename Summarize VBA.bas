Sub Summarize()

    Call column_names
    Call Tickers
    
End Sub

Sub column_names()

    Dim ws As Worksheet
    
    'Rename the top row of cells
    For Each ws In Worksheets
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("J:J").ColumnWidth = 6.29
        ws.Range("K:K").ColumnWidth = 12.86
        ws.Range("L:L").ColumnWidth = 14.29
        ws.Range("M:M").ColumnWidth = 17.57
    Next ws
End Sub

Sub Tickers()

    'Create a variable to track the last row in each sheet
    Dim lastrow As Integer
    
    'Create a counter to keep track of the row in the summary chart
    Dim a As Integer
    
    'Create a variable to store opening value
    Dim opening As Double
    
    'Create a variable to store closing value
    Dim closing As Double
    
    'Create a variable to calculate the percent change
    Dim change As Double
    
    'Create a variable to add total volume
    Dim volume As Variant
    
    volume = 0
    
    'Loop through each worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
        'Define last row for each worksheet
        lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
        
        'Reset a in each new worksheet
        a = 2
            
        'Loop through rows on each worksheet
        For i = 2 To lastrow
            
            'Determine when ticker name changes
            If ws.Cells(i, 1).Value <> ws.Cells(a - 1, 10).Value Then
                
                'Write down the new ticker name in the summary chart
                ws.Cells(a, 10).Value = ws.Cells(i, 1).Value
                
                'Store the opening value
                opening = ws.Cells(i, 3).Value
                
                'Store the volume value
                volume = ws.Cells(i, 7).Value
                
                'Increase the counter
                a = a + 1
                
            Else
                'Store the closing value each time the ticker doesn't change
                closing = ws.Cells(i, 6).Value
                
                'Document the yearly change, overwrite this value until the ticker changes
                ws.Cells(a - 1, 11).Value = Val(closing) - Val(opening)
                
                'Calculate the percentage change
                change = Val(ws.Cells(a - 1, 11).Value) / Val(opening)
                
                'Format the percentage change
                ws.Cells(a - 1, 12).Value = Format(change, "Percent")
                
                'Format the color of the cell to change with value
                If ws.Cells(a - 1, 12).Value >= 0 Then
                    ws.Cells(a - 1, 12).Interior.ColorIndex = 10
                Else
                    ws.Cells(a - 1, 12).Interior.ColorIndex = 3
                End If
                
                'Add the volume each time the ticker hasn't changed
                volume = volume + ws.Cells(i, 7).Value
                
                'Store the volume each time the ticker hasn't changed
                ws.Cells(a - 1, 13).Value = volume
            End If
        Next i
    
    Next ws

End Sub

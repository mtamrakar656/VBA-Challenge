Attribute VB_Name = "Module1"
Sub wall_street():
    
For Each ws In Worksheets

    ws.Range("J1") = "Ticker"
    ws.Range("K1") = "Total Stock Volume"
    
    'Declare variables
    Dim ticker As String
    Dim total_volume As Double
    Dim summary_table_row As Integer
    
    'Initialize variable
    total_volume = 0
    summary_table_row = 2
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
        'Calculate the total volume for each ticker
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            total_volume = total_volume + ws.Cells(i, 7).Value
        
            'Store ticker and corresponding total value to cells
            ws.Range("J" & summary_table_row).Value = ticker
            ws.Range("K" & summary_table_row).Value = total_volume
        
            'Move to next row to store value
            summary_table_row = summary_table_row + 1
        
            Else
            total_volume = total_volume + Cells(i, 3).Value
        
            End If
    
        Next i
    
    Next ws
    
End Sub

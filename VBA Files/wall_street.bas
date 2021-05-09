Attribute VB_Name = "wall_street"
Sub wall_street():
    
For Each ws In Worksheets

    'Assign Headers
    ws.Range("J1") = "Ticker"
    ws.Range("K1") = "Yearly Change"
    ws.Range("L1") = "Percent Change"
    ws.Range("M1") = "Total Stock Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    
    'Declare variables
    Dim ticker As String
    Dim total_volume As Double
    Dim open_value As Double
    Dim close_value As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim summary_table_row As Integer
    Dim lastrow As Long
    Dim i As Long
    Dim y As Long
    Dim increase_row_num As Integer
    Dim decrease_row_num As Integer
    Dim volume_row_num As Integer
                
    'Initialize variables
    total_volume = 0
    summary_table_row = 2
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
        'Calculate the total volume for each ticker
        For i = 2 To lastrow
                 
            'If the ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            ticker = ws.Cells(i, 1).Value
            
            'Stores result in variable
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            'Store ticker and corresponding total values to cells
            ws.Range("J" & summary_table_row).Value = ticker
            ws.Range("M" & summary_table_row).Value = total_volume
            
            'Reset Total Volume to zero
            total_volume = 0
                      
            'Set close_value
            close_value = ws.Cells(i, 6)
            
                'Calculate Yearly Change and Percent Change
                If open_value = 0 Then
                    yearly_change = 0
                    percent_change = 0
                Else:
                yearly_change = close_value - open_value
                percent_change = (close_value - open_value) / open_value
                End If
            
            'Store Yearly change and Percent change values to cells
            ws.Range("K" & summary_table_row).Value = yearly_change
            ws.Range("L" & summary_table_row).Value = percent_change
            ws.Range("L" & summary_table_row).Style = "Percent"
            ws.Range("L" & summary_table_row).NumberFormat = "0.00%"
            
             'Move to next row to store value
            summary_table_row = summary_table_row + 1
            
            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
                open_value = ws.Cells(i, 3)
                
            Else:
            'Keep adding volume
            total_volume = total_volume + Cells(i, 7).Value
        
            End If
           
        Next i
        
        For y = 2 To lastrow
        
        If ws.Range("K" & y).Value > 0 Then
            ws.Range("K" & y).Interior.ColorIndex = 4
            
        ElseIf ws.Range("K" & y).Value < 0 Then
            ws.Range("K" & y).Interior.ColorIndex = 3
            
        End If
        
        Next y
        
'Bonus
    
    'Take the max and min % increase and decrease
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("L2:L" & lastrow)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("L2:L" & lastrow)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("M2:M" & lastrow))
    
    'Get the corresponding ticker row minus the header row
    increase_row_num = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
    decrease_row_num = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
    volume_row_num = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("M2:M" & lastrow)), ws.Range("M2:M" & lastrow), 0)
    
    'Get the ticker using corresponding ticker row
    ws.Range("P2") = ws.Cells(increase_row_num + 1, 10)
    ws.Range("P3") = ws.Cells(decrease_row_num + 1, 10)
    ws.Range("P4") = ws.Cells(volume_row_num + 1, 10)
    
    Next ws
    
End Sub


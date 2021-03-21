Attribute VB_Name = "Module1"
Sub stickers()

' Creating new variables
Dim ticker, greatest_increase_ticker, greatest_decrease_ticker, greatest_total_ticker As String
Dim initial_open, final_close, summary_table_row As Integer
Dim yearly_change, percent_change, greatest_percent_increase, greatest_percent_decrease As Double
Dim total_stock_volume, greatest_total_volume As LongLong

'Counting Number of Rows
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Setting initial values
initial_open = Cells(2, 3).Value
summary_table_row = 2

'Formatting
Range("I:L").ColumnWidth = 16
Range("O:O").ColumnWidth = 20
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Range("K:K").NumberFormat = "0.00%"
Range("Q2:Q3").NumberFormat = "0.00%"

'Loop Summary Table
For i = 2 To lastRow
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ticker = Cells(i, 1).Value
        final_close = Cells(i, 6).Value
        yearly_change = final_close - initial_open
        
        If initial_open = 0 Then
            
            percent_change = 0
            
        Else
            
            percent_change = yearly_change / initial_open
        
        End If
        
        total_stock_volume = total_stock_volume + Cells(i, 7)
        Cells(summary_table_row, 9).Value = ticker
        Cells(summary_table_row, 10).Value = yearly_change
        
        If Cells(summary_table_row, 10).Value > 0 Then
            
            Cells(summary_table_row, 10).Interior.ColorIndex = 4
                    
        ElseIf Cells(summary_table_row, 10).Value < 0 Then
            
            Cells(summary_table_row, 10).Interior.ColorIndex = 3
        
        End If
                        
        Cells(summary_table_row, 11).Value = percent_change
        Cells(summary_table_row, 12).Value = total_stock_volume
        summary_table_row = summary_table_row + 1
        initial_open = Cells(i + 1, 3).Value
        total_stock_volume = 0
        
    Else
        
        total_stock_volume = total_stock_volume + Cells(i, 7)
        
    End If
    
Next i
                                           
'Reseting number of rows
lastRow = Cells(Rows.Count, 11).End(xlUp).Row

'Setting initial values, had to be done after the first loop
greatest_increase_ticker = Cells(2, 9).Value
greatest_percent_increase = Cells(2, 11).Value
greatest_decrease_ticker = Cells(2, 9).Value
greatest_percent_decrease = Cells(2, 11).Value
greatest_total_ticker = Cells(2, 9).Value
greatest_total_volume = Cells(2, 12).Value

'Loop Bonus
For j = 2 To lastRow
    
    If Cells(j, 11).Value > greatest_percent_increase Then
    
        greatest_increase_ticker = Cells(j, 9).Value
        greatest_percent_increase = Cells(j, 11).Value
    
    End If
    
    If Cells(j, 11).Value < greatest_percent_decrease Then
    
        greatest_decrease_ticker = Cells(j, 9).Value
        greatest_percent_decrease = Cells(j, 11).Value

    End If
    
    If Cells(j, 12).Value > greatest_total_volume Then
    
        greatest_total_ticker = Cells(j, 9).Value
        greatest_total_volume = Cells(j, 12).Value
        
    End If
    
Next j

'Printing final values
Cells(2, 16).Value = greatest_increase_ticker
Cells(3, 16).Value = greatest_decrease_ticker
Cells(4, 16).Value = greatest_total_ticker
Cells(2, 17).Value = greatest_percent_increase
Cells(3, 17).Value = greatest_percent_decrease
Cells(4, 17).Value = greatest_total_volume

End Sub

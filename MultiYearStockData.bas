Attribute VB_Name = "Module1"
Sub MultiYearStockData()
Dim ticker_symbol As String
Dim total_volume As Double
total_volume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim ws As Worksheet
yearly_change = 0
open_price = Cells(2, 3).Value
Dim j As Long



For Each ws In Worksheets
j = 2
total_volume = 0
Summary_Table_Row = 2
yearly_change = 0
open_price = Cells(2, 3).Value

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Insert Titles For Summary Tables
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    
    'Insert Titles for Greatest change table
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Set ticker symbol
            ticker_symbol = ws.Cells(i, 1).Value
            
            'Add to volume total
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            'Define yearly change and percent change variables
            close_price = ws.Cells(i, 6).Value
        
            
            'Yearly change and Percent change formulas
            yearly_change = close_price - open_price
            percent_change = yearly_change / open_price
            '
            'Format numbers
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00"
            ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
            
            'Print Ticker, Volume, Yearly change and Percent change
            'ws.Range("J" & Summary_Table_Row).Value = ticker_symbol
            'ws.Range("M" & Summary_Table_Row).Value = total_volume
            'ws.Range("K" & Summary_Table_Row).Value = yearly_change
            'ws.Range("L" & Summary_Table_Row).Value = percent_change
            
             ws.Cells(j, 10).Value = ticker_symbol
             ws.Cells(j, 11).Value = yearly_change
             ws.Cells(j, 12).Value = percent_change
             ws.Cells(j, 13).Value = total_volume
             j = j + 1
            
            
            
            'Conditional formating
            If ws.Range("K" & Summary_Table_Row).Value > 0 Then
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf ws.Range("K" & Summary_Table_Row).Value < 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            If ws.Range("L" & Summary_Table_Row).Value > 0 Then
                ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf ws.Range("L" & Summary_Table_Row).Value < 0 Then
                    ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            'Find max and min with ws.Function
            greatest_percent_increase = Application.WorksheetFunction.Max(ws.Range("L:L").Value)
            greatest_percent_decrease = Application.WorksheetFunction.Min(ws.Range("L:L").Value)
            max_total_volume = Application.WorksheetFunction.Max(ws.Range("M:M").Value)
                    
            'Format cells for percentages
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(3, 17).NumberFormat = "0.00%"
                    
            
            'Print Values
            If ws.Range("L" & Summary_Table_Row).Value = greatest_percent_increase Then
                ws.Cells(2, 16).Value = ws.Range("J" & Summary_Table_Row).Value
                ws.Cells(2, 17).Value = greatest_percent_increase
            End If
            If ws.Range("L" & Summary_Table_Row).Value = greatest_percent_decrease Then
                ws.Cells(3, 16).Value = ws.Range("J" & Summary_Table_Row).Value
                ws.Cells(3, 17).Value = greatest_percent_decrease
            End If
            If ws.Range("M" & Summary_Table_Row).Value = max_total_volume Then
                ws.Cells(4, 16).Value = ws.Range("J" & Summary_Table_Row).Value
                ws.Cells(4, 17).Value = max_total_volume
            End If
           
            
            'Add one to summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            'Reset opening price
            If ws.Cells(i + 1).Value <> 0 Then
            open_price = ws.Cells(i + 1, 3).Value
            End If
            
            'Reset volume total
            total_volume = 0
        'If its the same as the previous value
        Else
            'Add to volume total
            total_volume = total_volume + ws.Cells(i, 7).Value
        End If
    Next i
    
     
    
    
    'Print Values
    'ws.Cells(2, 17).Value = greatest_percent_increase
    'ws.Cells(3, 17).Value = greatest_percent_decrease
    'ws.Cells(4, 17).Value = max_total_volume
    
    'Add corresponding ticker symbol
    'WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    'ws.Cells(2, 16).Value = Application.WorksheetFunction.Match(ws.Cells(2, 17).Value, ws.Range("J:J" & Summary_Table_Row).Value, 0)
    'ws.Cells(3, 16).Value = ws.Cells(greatest_percent_decrease, 10)
    'ws.Cells(4, 16).Value = ws.Cells(max_total_volume, 10)
    
    
Next ws

End Sub


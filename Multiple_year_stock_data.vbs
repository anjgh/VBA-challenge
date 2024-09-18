Sub multiple_year_stock_data()

    'declaring variables
    Dim i As Long
    Dim lastRow As Long
    Dim opening_number As Double
    Dim closing_number As Double
    Dim output_counter As Integer
    Dim opening_number_flag As Boolean
    Dim percent_change As Double
    Dim total_volume As LongLong
    Dim greatest_percent_increase As Double
    Dim greatest_percent_decrease As Double
    Dim greatest_percent_volume As LongLong
    Dim greatest_increase_ticker_symbol As String
    Dim greatest_decrease_ticker_symbol As String
    Dim greatest_volume_ticker_symbol As String
    Dim ws As Worksheet
    
    'lopp through all sheets
    For Each ws In Worksheets
    
        'titles for the new columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
            
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'initilizing the counter variables
        output_counter = 1
        opening_number_flag = False
        
        For i = 2 To lastRow
            
            'starts when opening_number_flag is 0, reads the opening number and makes opening_number_flag a 1
            'this if statement reads the opening_number and assigns it to the variable named opening_number, at the start of every unique ticker
            If opening_number_flag = False Then
                opening_number = ws.Cells(i, 3).Value
                opening_number_flag = True
                
                'set total volume to be 0 so that in the next ticker's total volume can be calculated
                total_volume = 0
                
            'starts when the ticker value in row i and row i+1 are not the same
            'reads closing_number, adds one to output_counter
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                'read closing number at the end of each ticker
                closing_number = ws.Cells(i, 6).Value
                
                'position information in the correct cells
                output_counter = output_counter + 1
                
                'write the quarterly change on column 10, by subtracting closing_number to opening_number and assigning it to cell location
                ws.Cells(output_counter, 10).Value = closing_number - opening_number
                ws.Cells(output_counter, 10).NumberFormat = "0.00"
                
                'color quarterly change
                If ws.Cells(output_counter, 10).Value > 0 Then
                    ws.Cells(output_counter, 10).Interior.ColorIndex = 4
                ElseIf Cells(output_counter, 10).Value < 0 Then
                    ws.Cells(output_counter, 10).Interior.ColorIndex = 3
                End If
                
                'write the percent change on column 11, by dividing quarterly change by opening_number
                ws.Cells(output_counter, 11).Value = ws.Cells(output_counter, 10).Value / opening_number
                ws.Cells(output_counter, 11).NumberFormat = "0.00%"
                
                'write the ticker symbol on column 9
                ws.Cells(output_counter, 9).Value = ws.Cells(i, 1).Value
                
                'set opening number to be 0 so that in the next iteration, it will meet the first if statement condition
                opening_number_flag = False
                
                'adding the last row to get the total volume for each ticker
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                'write the total stock volume
                ws.Cells(output_counter, 12).Value = total_volume
                
                'write the greatest % increase and the greatest % decrease on column
                If output_counter = 2 Then
                    greatest_percent_increase = ws.Cells(output_counter, 11).Value
                    greatest_percent_decrease = ws.Cells(output_counter, 11).Value
                    greatest_total_volume = ws.Cells(output_counter, 12).Value
                End If
                
                'find greatest_percent_increase,greatest_percent_decrease, and associated ticker symbol
                If greatest_percent_increase < ws.Cells(output_counter, 11).Value Then
                    greatest_percent_increase = ws.Cells(output_counter, 11).Value
                    greatest_increase_ticker_symbol = ws.Cells(output_counter, 9).Value
                End If
                
                If greatest_percent_decrease > ws.Cells(output_counter, 11).Value Then
                    greatest_percent_decrease = ws.Cells(output_counter, 11).Value
                    greatest_decrease_ticker_symbol = ws.Cells(output_counter, 9).Value
                End If
                
                'write the greatest total volume on column
                If greatest_percent_volume < ws.Cells(output_counter, 12).Value Then
                    greatest_percent_volume = ws.Cells(output_counter, 12).Value
                    greatest_volume_ticker_symbol = ws.Cells(output_counter, 9).Value
                End If
                
            End If
            
            'for each row, take the current total volume and add it with the next rows volume, then assign to variable total_volume
            total_volume = total_volume + ws.Cells(i, 7).Value
              
        Next i
        
        'write greatest percent increase, greatest percent decrease, and greatest percent volume
        ws.Range("R2").Value = greatest_percent_increase
        ws.Range("R2").NumberFormat = "0.00%"
        ws.Range("R3").Value = greatest_percent_decrease
        ws.Range("R3").NumberFormat = "0.00%"
        ws.Range("R4").Value = greatest_percent_volume
        ws.Range("R4").NumberFormat = "0.00E+00"
        
        'write the ticker symbol associated for the greatest percent increase, greatest percent decrease, and greatest percent volume
        ws.Range("Q2").Value = greatest_increase_ticker_symbol
        ws.Range("Q3").Value = greatest_decrease_ticker_symbol
        ws.Range("Q4").Value = greatest_volume_ticker_symbol
        
    Next ws
End Sub



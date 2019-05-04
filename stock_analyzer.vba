Sub stock_market_analyzer()

'Assign variables for each desired result
Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim year_change As Double
Dim percent_change As Double
Dim volume As Double
Dim last_row As Long
Dim row_counter As Double
Dim results_counter As Double

'Iterate through each worksheet in workbook

For Each ws In ActiveWorkbook.Worksheets

    'Label the desired fields
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 9).Font.Bold = True
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 10).Font.Bold = True
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 11).Font.Bold = True
    ws.Cells(1, 12).Value = "Total Stock volume"
    ws.Cells(1, 12).Font.Bold = True

    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 15).Font.Bold = True
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 15).Font.Bold = True
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 15).Font.Bold = True

    'Values start at row 2
    results_counter = 2

    'Find the last row
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'ws.Cells(2, 13).Value = last_row
    volume = 0

    'Find the opening value
    year_open = ws.Cells(2, 3).Value

    'We iterate from row 2 to the last row
    For row_counter = 2 To last_row

        volume = volume + ws.Cells(row_counter, 7).Value

        'We compare the current ticker to the one above
        If (ws.Cells(row_counter - 1, 1).Value = ws.Cells(row_counter, 1).Value And ws.Cells(row_counter + 1, 1).Value <> ws.Cells(row_counter, 1).Value) Then

            'If the current ticker is the same as above and different than the one below
            year_close = ws.Cells(row_counter, 6).Value

            'Calculate the change between opening and closing values
            year_change = year_close - year_open

            'Account for divisions by 0
            If year_open = 0 And year_close <> 0 Then
                percent_change = year_close / year_close
            
            ElseIf year_open = 0 And year_close = 0 Then
                percent_change = 0
            
            Else
                percent_change = (year_close - year_open) / year_open
            
            End If

        'We move to the next year open
        year_open = ws.Cells(row_counter + 1, 3).Value

        'Add the year change for each ticker
        ws.Cells(results_counter, 10).Value = year_change

        'Conditional formattting for change: if > 0, green; if < 0, red
        If ws.Cells(results_counter, 10).Value > 0 Then
            ws.Cells(results_counter, 10).Interior.ColorIndex = 4
        
        Else
            ws.Cells(results_counter, 10).Interior.ColorIndex = 3

        End If

        'Add the percent change for each ticker
        ws.Cells(results_counter, 11).Value = percent_change
        ws.Cells(results_counter, 11).NumberFormat = "0.00%"

    End If

    'Add the total volume for each ticker
    'If the current ticker is the same as above and different than the one below
    If (ws.Cells(row_counter + 1, 1).Value <> ws.Cells(row_counter, 1).Value) Then

        'Add ticker
        ws.Cells(results_counter, 9).Value = ws.Cells(row_counter, 1).Value

        'Add total volume
        ws.Cells(results_counter, 12).Value = volume

        'Reset the value to 0
        volume = 0

        'Increase the results counter by 1 for the next iteration
        results_counter = results_counter + 1

    End If

Next row_counter
    
    'Label the desired fields

    
    'Assigns values for comparison
    ws.Cells(2, 17).Value = 0
    ws.Cells(3, 17).Value = 0
    ws.Cells(4, 17).Value = 0

    For results_counter = 2 To last_row

        If ws.Cells(results_counter, 11).Value > ws.Cells(2, 17).Value Then

            ws.Cells(2, 17).Value = ws.Cells(results_counter, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(results_counter, 9).Value
            
        End If

        'If the percent change is lower than the value in cell (3,17)
        If ws.Cells(results_counter, 11).Value < ws.Cells(3, 17).Value Then

            'Store the new percent change value in cell (3,17)
            ws.Cells(3, 17).Value = ws.Cells(results_counter, 11).Value

            'Find the lowest value and add it the results cell
            ws.Cells(3, 16).Value = ws.Cells(results_counter, 9).Value
        
        End If

        'If the volume is greater than the value in cell (4, 17)
         If ws.Cells(results_counter, 12).Value > ws.Cells(4, 17).Value Then

            'Store the larger volume in cell (4,17)
            ws.Cells(4, 17).Value = ws.Cells(results_counter, 12).Value

            'Find the highest volume and add it in the results cell
            ws.Cells(4, 16).Value = ws.Cells(results_counter, 9).Value
            
        End If
       
        Next results_counter
        
        'Format the results table
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"

Next ws

End Sub

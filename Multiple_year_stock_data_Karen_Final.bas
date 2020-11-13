Attribute VB_Name = "Module1"

Sub stockmarket2()

'loop through all the worksheets
Dim ws As Worksheet

For Each ws In Worksheets

    
    'setting variables
    Dim ticker_symbol As String
    Dim total_vol As Double
    total_vol = 0
    Dim year_open As Double
    year_open = 0
    Dim year_close As Double
    year_closed = 0
    Dim year_change As Double
    year_change = 0
    Dim percent_change As Double
    percent_change = 0

    Dim summary_row As Long
    summary_row = 2
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
    'print summary table headers
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
For I = 2 To last_row
    'loop - search ticker symbols
    If (ws.Cells(I, 1).Value <> ws.Cells(I - 1, 1).Value) Then
        year_open = ws.Cells(I, 3).Value
    End If
    
    'Total of volume for each row
        total_vol = total_vol + ws.Cells(I, 7)
    
        If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
            'ticker ? move st
            ws.Cells(summary_row, 9).Value = ws.Cells(I, 1).Value
    
            'total stock volume ? move st
            ws.Cells(summary_row, 12).Value = total_vol

            'year end price
            year_close = ws.Cells(I, 6).Value
    
            'price change for yr
            year_change = year_close - year_open
            ws.Cells(summary_row, 10).Value = year_change

            'conditional to highlight year change  - green +/red ?
        If year_change >= 0 Then
            ws.Cells(summary_row, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(summary_row, 10).Interior.ColorIndex = 3
        End If

    'percentage change for year mv to st
    If year_open = 0 And year_close = 0 Then
        percent_change = 0
        ws.Cells(summary_row, 11).Value = percent_change
        ws.Cells(summary_row, 11).NumberFormat = "0.00%"
    ElseIf year_open = 0 Then
        ws.Cells(summary_row, 11).Value = percent_change
    Else
        percent_change = year_change / year_open
        ws.Cells(summary_row, 11).Value = percent_change
        ws.Cells(summary_row, 11).NumberFormat = "0.00%"
    End If

        summary_row = summary_row + 1

        'because at the end of the row, reset to startover
        total_vol = 0
        year_open = 0
        year_close = 0
        year_change = 0
        percent_change = 0

    End If
Next I
    
'setting variables for Greatest % inc, % dec, and total vol
    Dim GreatestInc As Double
    Dim GreatestDec As Double
    Dim GreatestTotalVol As Double
  'printing greatest summary table for Greatest of things
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    
    'setting starting point (baseline) for greatest of things
    GreatestInc = 0
    GreatestDec = 0
    GreatestTotalVol = 0
    
    'count number of rows of st
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'greatest total vol
    For m = 2 To last_row
        If (ws.Cells(m, 12).Value > GreatestTotalVol) Then
            GreatestTotalVol = ws.Cells(m, 12).Value
            'printing tickers in summary table
            ws.Cells(4, 16).Value = ws.Cells(m, 9).Value
        End If
    Next m
    
    'print greatest total vol
    ws.Cells(4, 17).Value = GreatestTotalVol
    
    'greatest % inc and dec
    For m = 2 To last_row
       If (ws.Cells(m, 11).Value > GreatestInc) Then
            GreatestInc = ws.Cells(m, 11).Value
            'printing ticker inc
            ws.Cells(2, 16).Value = ws.Cells(m, 9).Value
        ElseIf (ws.Cells(m, 11).Value < GreatestDec) Then
            GreatestDec = ws.Cells(m, 11).Value
            'print ticker dec
            ws.Cells(3, 16).Value = ws.Cells(m, 9).Value
        End If
    Next m
  
    'print greatest % inc and dec in table
    ws.Cells(2, 17).Value = GreatestInc
    ws.Cells(3, 17).Value = GreatestDec

'format to percentage
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 17).NumberFormat = "#0.0000E+0"

'auto fit table
ws.Columns("I:L").EntireColumn.AutoFit
ws.Columns("O:Q").EntireColumn.AutoFit

Next ws

End Sub




Option Explicit
Sub sort()

    Dim ws As Worksheet
    Dim ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim i As Long
    Dim total_vol As Variant
    Dim last_row As Long
    Dim table_row As Long
    
    Dim max_percent_inc As Double
    Dim max_percent_inc_ticker As String
    Dim max_percent_dec As Double
    Dim max_percent_dec_ticker As String
    Dim max_total_vol As Variant
    Dim max_total_vol_ticker As String

    For Each ws In Worksheets
    
    'insert headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'set variables for summary table loop
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    year_open = ws.Cells(2, 3).Value
    table_row = 2
    
    'loop
    For i = 2 To last_row
        'add values for total vol
        total_vol = total_vol + ws.Cells(i, 7).Value
        'detect new ticker coming up next, obtain values for summary table
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            year_close = ws.Cells(i, 6).Value
            yearly_change = year_close - year_open
            
            If year_open <> 0 Then
                percent_change = yearly_change / year_open
            Else
                percent_change = 0
            End If
                        
             'insert values into summary table
            ws.Cells(table_row, 9).Value = ticker
            ws.Cells(table_row, 10).Value = yearly_change
            
            If yearly_change < 0 Then
                ws.Cells(table_row, 10).Interior.Color = vbRed

            ElseIf yearly_change > 0 Then
                ws.Cells(table_row, 10).Interior.Color = vbGreen
                
            Else
                ws.Cells(table_row, 10).Interior.Color = vbWhite
                
            End If
            
            ws.Cells(table_row, 11).Value = percent_change
            ws.Cells(table_row, 12).Value = total_vol
            'add 1 to table row to prepare summary table for next ticker
            table_row = table_row + 1
            'retrieve the next ticker's year_open value before continuing the loop
            year_open = ws.Cells(i + 1, 3).Value
            
            'reset total vol value
            total_vol = 0
        
         End If
    Next i
    
    'format numbers to %
    ws.Columns("K").NumberFormat = "0.00%"



'hard solution table
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    'define initial values to compare to
    max_percent_inc = ws.Cells(2, 11).Value
    max_percent_inc_ticker = ws.Cells(2, 9).Value
    max_percent_dec = ws.Cells(2, 11).Value
    max_percent_dec_ticker = ws.Cells(2, 9).Value
    max_total_vol = ws.Cells(2, 12).Value
    max_total_vol_ticker = ws.Cells(2, 9).Value
    
    'define last_row
    last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    'cycle through data for max values
    For i = 3 To last_row
        If max_percent_dec > ws.Cells(i, 11).Value Then
            max_percent_dec = ws.Cells(i, 11).Value
            max_percent_dec_ticker = ws.Cells(i, 9).Value
        End If
        
        If max_percent_inc < ws.Cells(i, 11).Value Then
            max_percent_inc = ws.Cells(i, 11).Value
            max_percent_inc_ticker = ws.Cells(i, 9).Value
        End If
        
        If max_total_vol < ws.Cells(i, 12).Value Then
            max_total_vol = ws.Cells(i, 12).Value
            max_total_vol_ticker = ws.Cells(i, 9).Value
        End If
    Next i
    
    
    'insert values into table
    ws.Range("O2").Value = max_percent_inc_ticker
    ws.Range("P2").Value = max_percent_inc
    ws.Range("O3").Value = max_percent_dec_ticker
    ws.Range("P3").Value = max_percent_dec
    ws.Range("O4").Value = max_total_vol_ticker
    ws.Range("P4").Value = max_total_vol
    
    Next ws

End Sub


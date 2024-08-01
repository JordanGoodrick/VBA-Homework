Attribute VB_Name = "Module1"

Sub vba_test()
Dim WS As Worksheet

'Loop for all quarters:
For Each WS In ThisWorkbook.Worksheets

    'Create variables to track data and display in a table
    Dim ticker_symbol As String
    ticker_symbol = ""
    Dim summary_table_row As Integer
    summary_table_row = 2
    Dim quarter_start As Double
    quarter_start = 0
    Dim quarter_end As Double
    quarter_end = 0
    Dim ticker_volume As Double
    ticker_volume = 0
    Dim quarterly_change As Double
    quarterly_change = 0
    Dim quarterly_percent As Double
    quarterly_percent = 0
    Dim last_row As Integer
    
    
    'isolate highest increases, decreases and volume
    Dim greatest_volume As Double
    
    Dim greatest_increase As Double
    
    Dim greatest_decrease As Double
    
    Dim max_ticker As String
    
    Dim min_ticker As String
    
    Dim high_volume_ticker As String
    
    Dim low_volume_ticker As String
    
    Dim greatest_total_ticker As String
    
    
    
    
    'Format output table
    WS.Cells(1, 9).Value = "ticker"
    WS.Cells(1, 10).Value = "Quarterly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"
    WS.Cells(1, 16).Value = "Ticker"
    WS.Cells(1, 17).Value = "Value"
    WS.Cells(2, 15).Value = "Greatest % Increase"
    WS.Cells(3, 15).Value = "Greatest % Decrease"
    WS.Cells(4, 15).Value = "Greatest total volume"
    
    
    'Find the ticker name
        For i = 2 To 96050
    
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
            '     Set the ticker symbol.
                ticker_symbol = WS.Cells(i, 1).Value
                ' add to the ticker volume
                ticker_volume = ticker_volume + WS.Cells(i, 7).Value
                ' print ticker symbol
                WS.Range("I" & summary_table_row).Value = ticker_symbol
                'find differences in quarterly earnings
                quarter_start = WS.Cells(i, 3).Value
                quarter_end = WS.Cells(i, 6).Value
                'print quarterly change
                quarterly_change = quarter_end - quarter_start
                WS.Range("J" & summary_table_row).Value = quarterly_change
                'print percent change
                quarterly_percent = (quarterly_change / quarter_start) * 100
               WS.Range("K" & summary_table_row).Value = quarterly_percent
                'print total stock volume
                WS.Range("L" & summary_table_row).Value = ticker_volume
                summary_table_row = summary_table_row + 1
                'Reset the volume total
                ticker_volume = 0
            Else
                ticker_volume = ticker_volume + WS.Cells(i, 7).Value
            End If
            'Adding conditional formating to the quarterly change row.
            If (quarterly_change > 0) Then
                WS.Range("J" & summary_table_row - 1).Interior.ColorIndex = 4
            ElseIf (quarterly_change < 0) Then
                WS.Range("J" & summary_table_row - 1).Interior.ColorIndex = 3
            End If
            If (quarterly_percent > 0) Then
                WS.Range("K" & summary_table_row - 1).Interior.ColorIndex = 4
            ElseIf (quarterly_percent < 0) Then
                WS.Range("K" & summary_table_row - 1).Interior.ColorIndex = 3
            End If
            'Find the tickers with outlier stats
            If (quarterly_change > greatest_increase) Then
                high_volume_ticker = ticker_symbol
                greatest_increase = quarterly_percent
            ElseIf (quarterly_change < greatest_decrease) Then
                low_volume_ticker = ticker_symbol
                greatest_decrease = quarterly_percent
            End If
            If (ticker_volume > greatest_volume) Then
                greatest_volume = ticker_volume
                greatest_total_ticker = ticker_symbol
            End If
            ticker_volume = 0
            quarterly_change = 0
        
        Next i
    
        WS.Range("P2").Value = high_volume_ticker
        WS.Range("P3").Value = low_volume_ticker
        WS.Range("P4").Value = greatest_total_ticker
        WS.Range("q2").Value = greatest_increase
        WS.Range("q3").Value = greatest_decrease
        WS.Range("q4").Value = greatest_volume
        
Next WS
  
  
End Sub




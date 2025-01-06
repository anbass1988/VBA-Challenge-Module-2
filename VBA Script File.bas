Attribute VB_Name = "Module1"


Sub tickerStock()

    ' Loop through each worksheet
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        
        ' Find the last row of the table
        Dim last_row As Long
        last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Add headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Dim open_price As Double
        Dim close_price As Double
        Dim quarterly_change As Double
        Dim ticker As String
        Dim percent_change As Double
        Dim volume As Double
        Dim Row As Long
        Dim Column As Long
        
        volume = 0
        Row = 2
        Column = 1
        
        ' Setting the initial price
        open_price = ws.Cells(2, Column + 2).Value
        
        ' Loop through all tickers to check for mismatch
        Dim i As Long
        For i = 2 To last_row
            If ws.Cells(i + 1, Column).Value <> ws.Cells(i, Column).Value Then
                ' Setting ticker name
                ticker = ws.Cells(i, Column).Value
                ws.Cells(Row, Column + 8).Value = ticker
                
                ' Setting closing price
                close_price = ws.Cells(i, Column + 5).Value
                
                ' Calculate quarterly change
                quarterly_change = close_price - open_price
                ws.Cells(Row, Column + 9).Value = quarterly_change
                
                ' Calculate percent change
                percent_change = quarterly_change / open_price
                ws.Cells(Row, Column + 10).Value = percent_change
                ws.Cells(Row, Column + 10).NumberFormat = "0.00%"

                ' calculate total volume per quarter
                volume = volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = volume
                
                ' Iterate to the next row
                Row = Row + 1
                
                ' Reset open price to next ticker
                open_price = ws.Cells(i + 1, Column + 2).Value
                
                ' Reset volume for next ticker
                volume = 0
            Else
                volume = volume + ws.Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Find the last row of ticker column
        Dim quarterly_change_last_row As Long
        quarterly_change_last_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        
        ' Set the cell colors
        Dim j As Long
        For j = 2 To quarterly_change_last_row
            If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 10 ' Green for positive or zero change
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3 ' Red for negative change
            End If
        Next j
        
        ' Set headers for summary
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Find the highest value of each ticker
        Dim k As Long
        For k = 2 To quarterly_change_last_row
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quarterly_change_last_row)) Then
                ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quarterly_change_last_row)) Then
                ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quarterly_change_last_row)) Then
                ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
            End If
        Next k
        
        ws.Range("I:Q").Font.Bold = True
        ws.Range("I:Q").EntireColumn.AutoFit
    Next ws

End Sub

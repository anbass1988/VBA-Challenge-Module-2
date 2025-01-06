Attribute VB_Name = "Module1"
Sub tickerStock():

'worksheet loop

Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
    ws.Activate
    
    'find the last row of the table
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
    'add headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quaterly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    Dim open_price As Double
    Dim close_price As Double
    Dim quatetly_change As Double
    Dim ticker As String
    Dim percent_change As Double
    
    volume = 0
    Row = 2
    Column = 1
    
'setting the initial price
    open_price = Cells(2, Column + 2).Value
    
      'loop through all ticker to check for mismacth
    For i = 2 To last_row
    
        If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
        
            ' setting ticker name
            ticker = Cells(i, Column).Value
            Cells(Row, Column + 8).Value = ticker
            
            ' setting closing price
            close_price = Cells(i, Column + 5).Value
            
            ' calculate quarterly change
            quaterly_Change = close_price - open_price
            Cells(Row, Column + 9).Value = quaterly_Change
            
            ' calculate percent change
                percent_change = quaterly_Change / open_price
                Cells(Row, Column + 10).Value = percent_change
                Cells(Row, Column + 10).NumberFormat = "0.00%"
                
            ' calculate total volume per quarter
                volume = volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = volume
                
            ' iterate the next row
            Row = Row + 1
            
            ' reset open price to next ticker
            open_price = Cells(i + 1, Column + 2)
            
            ' reset volume for next tcker
            volume = 0
            
        Else
            volume = volume + Cells(i, colmn + 6).Value
        End If
    Next i
    
    ' find the last row of ticker column
    Quaterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' set the Cell Colors
    For j = 2 To Quaterly_change_last_row
        If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
            Cells(j, 10).Interior.ColorIndex = 10
        ElseIf Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.ColorIndex = 3
        End If
    Next j
    
    ' set Ticker, Value, Greatest %, Increase, % Decrease, and Total volume headers
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    ' find the highest value of each ticket
    For k = 2 To Quaterly_change_last_row
    If Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("k2:K" & Quaterly_change_last_row)) Then
        Cells(2, 16).Value = Cells(k, 9).Value
        Cells(2, 17).Value = Cells(k, 11).Value
        Cells(2, 17).NumberFormat = "0.00%"
    ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Quaterly_change_last_row)) Then
        Cells(3, 16).Value = Cells(k, 9).Value
        Cells(3, 17).Value = Cells(k, 11).Value
        Cells(3, 17).NumberFormat = "0.00%"
    ElseIf Cells(k, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Quaterly_change_last_row)) Then
        Cells(4, 16).Value = Cells(k, 9).Value
        Cells(4, 17).Value = Cells(k, 12).Value
    End If
Next k
    
ActiveSheet.Range("I:Q").Font.Bold = True
ActiveSheet.Range("I:Q").EntireColumn.AutoFit
Worksheets("A").Select

Next


End Sub




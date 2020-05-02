
Sub Stocks()

    
'Create a loop

    Dim ws As Worksheet
    
'Start loop

    For Each ws In Worksheets
 
'Create column labels

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    Dim ticker_symbol As String

    Dim total_vol As Double
    total_vol = 0

    Dim rowcount As Long
    rowcount = 2
    
    Dim year_open As Double
    year_open = 0
    
    Dim year_close As Double
    year_close = 0
    
    Dim year_change As Double
    year_change = 0
    
    Dim percent_change As Double
    percent_change = 0

    
'Set variable for total rows
    Dim lastrow As Long

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Loop to search
    For i = 2 To lastrow
    
'Conditional to grab year open price
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    
    year_open = ws.Cells(i, 3).Value
    
    End If
    
'Total up the volume for each row to determine the total stock volume for the year
    total_vol = total_vol + ws.Cells(i, 7)
    
'Conditional to determine if the ticker symbol is changing
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
'Move ticker symbol
    ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
    
'Move total stock

    ws.Cells(rowcount, 12).Value = total_vol
    
'Grab year end price
    year_close = ws.Cells(i, 6).Value
    
'Calculate the price change.
    year_change = year_close - year_open
    ws.Cells(rowcount, 10).Value = year_change
    
'Conditional to format and highlight.
    If year_change >= 0 Then
    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
    End If
    
'Conditional for calculating percent change
    If year_open = 0 And year_close = 0 Then

'Starting at zero and ending at zero will be a zero increase. Cannot use a formula because
    percent_change = 0
    ws.Cells(rowcount, 11).Value = percent_change
    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
    ElseIf year_open = 0 Then

'"New Stock" as percent change.

    Dim percent_change_NA As String
    percent_change_NA = "New Stock"
    ws.Cells(rowcount, 11).Value = percent_change

    Else
    percent_change = year_change / year_open
    ws.Cells(rowcount, 11).Value = percent_change
    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
    End If

    rowcount = rowcount + 1

    
'Reset total stock volume, year open price, year close price, year change, year percent change

    total_vol = 0
    year_open = 0
    year_close = 0
    year_change = 0
    percent_change = 0
    
    End If
    
    Next i

    
'Create a best/worst performance table
'Titles

    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    
'Assign lastrow
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

    
'Set variables
    Dim best_stock As String
    Dim best_value As Double

    
'best performer
    best_value = ws.Cells(2, 11).Value
    
    Dim worst_stock As String
    Dim worst_value As Double

    
'worst performer
    worst_value = ws.Cells(2, 11).Value
    
    Dim most_vol_stock As String
    Dim most_vol_value As Double

    
'most value
    most_vol_value = ws.Cells(2, 12).Value
    
'Loop
    For j = 2 To lastrow
    
' best performer
    If ws.Cells(j, 11).Value > best_value Then
    best_value = ws.Cells(j, 11).Value
    best_stock = ws.Cells(j, 9).Value
    End If
    
' worst performer
    If ws.Cells(j, 11).Value < worst_value Then
    worst_value = ws.Cells(j, 11).Value
    worst_stock = ws.Cells(j, 9).Value
    End If
   
' greatest volume traded
    If ws.Cells(j, 12).Value > most_vol_value Then
    most_vol_value = ws.Cells(j, 12).Value
    most_vol_stock = ws.Cells(j, 9).Value
    End If
    
    Next j
    
'Move

    ws.Cells(2, 16).Value = best_stock
    ws.Cells(2, 17).Value = best_value
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = worst_stock
    ws.Cells(3, 17).Value = worst_value
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = most_vol_stock
    ws.Cells(4, 17).Value = most_vol_value

'Autofit

    ws.Columns("I:L").EntireColumn.AutoFit
    ws.Columns("O:Q").EntireColumn.AutoFit

    
    Next ws


    
End Sub
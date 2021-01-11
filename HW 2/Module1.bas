Attribute VB_Name = "Module1"
Sub Stock_Analysis()
    'Set a variable to cycle through the worksheet
Dim ws As Worksheet
    'Start a for loop
For Each ws In Worksheets

    'column labels for the table
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

    'Set variable to hold total stock volume
Dim tot_vol As Double
    'populate the variable
tot_vol = 0
  
Dim ticker_symbol As String
Dim rowcount As Long
    rowcount = 2
    
    'Set variable to hold year open price
Dim year_open As Double
    year_open = 0
    
    'variable to hold year end price
Dim year_end As Double
    year_end = 0
    
    'variable to hold change in price for the year
Dim year_change As Double
    year_change = 0
    
    'percent change in price for the year
Dim percent_change As Double
    percent_change = 0
    
    'total rows
Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'search through ticker symbols
    
For i = 2 To lastrow

    'Conditional to grab year open price
If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    year_open = ws.Cells(i, 3).Value
End If

    'Total up the volume for each row to find the total stock volume
    tot_vol = tot_vol + ws.Cells(i, 7)
    
    'Conditional to detemine if the ticker symbol is chqnging
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    'move ticker symbol to summary table
    ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
    
    'move total stock volume to the summary table
    ws.Cells(rowcount, 12).Value = tot_vol
    
    'year end price
    year_end = ws.Cells(i, 6).Value
    
    'the price change for the year
    year_change = year_end - year_open
    
    'move year_change to the summary table
    ws.Cells(rowcount, 10).Value = year_change
    
    'format to highlight positive and negative change
    
If year_change >= 0 Then
    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
Else
    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
End If

    'Calculate the percent change for the year
If year_open = 0 And year_end = 0 Then
    percent_change = 0
    
    'move to the summary table
ws.Cells(rowcount, 11).Value = percent_change
    
    'format as a percentage
ws.Cells(rowcount, 11).NumberFormat = "0.00%"

ElseIf year_open = 0 Then

    'Set a variable to hold the actual price inccrease
    
Dim percent_change_NA As String

    percent_change_NA = "New Stock"
    ws.Cells(rowcount, 11).Value = percent_change
    
Else

    percent_change = year_change / year_open
    ws.Cells(rowcount, 11).Value = perrcent_change
    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
    
End If

    'go to the next empty row in the summary table
    
    rowcount = rowcount + 1
    
    ' Reset the values
    tot_vol = 0
    year_open = 0
    year_end = 0
    year_change = 0
    percent_change = 0
    
End If
Next i

    'Create a best/worst performance table
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Assign lastrow to count the number of rows in the summary table
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    
    'Set variable to hold best performer, worst performer,and stock with the most volume
Dim best_stock As String
Dim best_value As Double
    'Set best pereformer equal to the first stock
    best_value = ws.Cells(2, 11).Value
    
Dim worst_stock As String
Dim worst_value As Double
    'Set worst performer equal to the first stock
    woest_value = ws.Cells(2, 11).Value
    
Dim most_vol_stock As String
Dim most_vol_value As Double
    'Set most volume equal to the first stock
    most_vol_value = ws.Cells(2, 12).Value
    
    'Loop through summary table
For j = 2 To lastrow

    'Conditional to determine best performer
If ws.Cells(j, 11).Value > best_value Then

    best_value = ws.Cells(j, 11).Value
    best_stock = ws.Cells(j, 9).Value
End If

    'Conditional to determine worst performer
If ws.Cells(j, 11).Value < worst_value Then
    
    worst_value = ws.Cells(j, 11).Value
    worst_stock = ws.Cells(j, 9).Value
End If

    'Conditional to determine stock with the greatest volume traded
If ws.Cells(j, 12).Value > most_vol_value Then
    
    most_vol_value = ws.Cells(j, 12).Value
    most_vol_stock = ws.Cells(j, 9).Value
End If

Next j

    'Move best performer, worst performer, and stock with the most volume items to the performer table
    ws.Cells(2, 16).Value = best_stock
    ws.Cells(2, 17).Value = best_value
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ws.Cells(3, 16).Value = worst_stock
    ws.Cells(3, 17).Value = worst_value
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ws.Cells(4, 16).Value = most_vol_stock
    ws.Cells(4, 17).Value = most_vol_value
    
    ws.Columns("I:L").EntireColumn.AutoFit
    ws.Columns("O:Q").EntireColumn.AutoFit
Next ws

End Sub

Attribute VB_Name = "Module2"
Sub analyzeStocks()
    '---------------------------------------------
    'SOLUTION START
    '---------------------------------------------
    'Loop through each worksheet
    For Each ws In Worksheets
    
        'Declare variables
        Dim i As Long
        Dim j As Long
        Dim Total As Double
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        'Initialize variables (unless initialized in a loop later on).
        Total = 0
        j = 0
        openingPrice = ws.Cells(2, 3).Value
        closingPrice = 0
    

        'print table headers
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Value"
        'Find the last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'For A2 until the last row...
        For i = 2 To lastRow
            'If A3 does not match A2...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Add A2's volume to the total
                Total = Total + ws.Cells(i, 7).Value
                'Print the ticker.
                'Note: It's "2 + j" because we want to start the new column at "J2".
                ws.Range("J" & 2 + j).Value = ws.Cells(i, 1).Value
                'Print the total.
                ws.Range("M" & 2 + j).Value = Total
                'reset total to zero for the next ticker.
                Total = 0
                'Opening price minus closing price
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                ws.Range("K" & 2 + j).Value = yearlyChange
                'Red if yearlyChange is negative, green if positive. Otherwise white.
                If yearlyChange < 0 Then
                     ws.Range("K" & 2 + j).Interior.Color = vbRed
                ElseIf yearlyChange > 0 Then
                     ws.Range("K" & 2 + j).Interior.Color = vbGreen
                ElseIf yearlyChange Then
                     ws.Range("K" & 2 + j).Interior.Color = vbWhite
                End If
                'find percent change and format as percent value
                percentChange = yearlyChange / openingPrice
                ws.Range("L" & 2 + j).Value = FormatPercent(percentChange)
                'reset opening price to next ticker's value.
                openingPrice = ws.Cells(i + 1, 3).Value
                'j moves so printing can move to the next row.
                j = j + 1
            'Otherwise (or, in other words if the ticker in A3 and A2 is the same)...
            Else
                'Add to the total value.
                Total = Total + ws.Cells(i, 7).Value
            End If
        Next i
        '---------------------------------------------
        'BONUS SOLUTION START
        '---------------------------------------------
        'define variables
        Dim x As Long
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        'print table header and first column
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        'Find greatest % increase, decrease and volume within each range
        greatestIncrease = WorksheetFunction.Max(ws.Range("L:L"))
        greatestDecrease = WorksheetFunction.Min(ws.Range("L:L"))
        'print values
        ws.Cells(2, 18).Value = FormatPercent(greatestIncrease)
        ws.Cells(3, 18).Value = FormatPercent(greatestDecrease)
        ws.Cells(4, 18).Value = WorksheetFunction.Max(ws.Range("M:M"))
        'loop through percent change column and find ticker for greatest increase and decrease
        For x = 2 To lastRow
            If ws.Range("L" & x).Value = greatestIncrease Then
                ws.Cells(2, 17).Value = ws.Cells(x, 10).Value
            ElseIf ws.Range("L" & x).Value = greatestDecrease Then
                ws.Cells(3, 17).Value = ws.Cells(x, 10).Value
            End If
        Next x
        'loop through total stock volume column and find ticker for greatest volume
        For x = 2 To lastRow
            If ws.Cells(x, 13).Value = ws.Cells(4, 18).Value Then
               ws.Cells(4, 17).Value = ws.Cells(x, 10).Value
            End If
        Next x
        '---------------------------------------------
        'AESTHETICS EDITS
        '---------------------------------------------
        'adjust column sizes
        ws.Columns("H:H").ColumnWidth = 5
        ws.Columns("I:I").ColumnWidth = 5
        ws.Columns("J:J").EntireColumn.AutoFit
        ws.Columns("K:K").EntireColumn.AutoFit
        ws.Columns("L:L").EntireColumn.AutoFit
        ws.Columns("M:M").EntireColumn.AutoFit
        ws.Columns("N:N").ColumnWidth = 5
        ws.Columns("O:O").ColumnWidth = 5
        ws.Columns("P:P").EntireColumn.AutoFit
        ws.Columns("Q:Q").EntireColumn.AutoFit
        ws.Columns("R:R").EntireColumn.AutoFit
        'bold headers
        ws.Range("J1:M1").Font.Bold = True
        ws.Range("P2:P4").Font.Bold = True
        ws.Range("Q1:R1").Font.Bold = True
    Next ws
End Sub







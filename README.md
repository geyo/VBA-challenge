# Multi-Year Stock Analysis
By Grace Yoo

**Programming Language Used:** Virtual Basic for Application (VBA)

<h2> Description</h2>
The purpose of this project is to use VBA to aggregate statistics on a stock market toy dataset with three years worth of data. The script loops through all stocks for one year and outputs:

 - The ticker symbol
 - Yearly Change from opening price to closing price
 - Percentage Change from opening price to closing price
 - Total stock volume

Additionally, the code finds and reports the stocks with the greatest percentage increase, decrease and total volume for that year.

<h3> Solution </h3>

A for loop iterates through each row fo the worksheet until the last row. 

The code adds the stock's name and initial value to the list. If the stock already exists in teh list, the code aggregates the stock's volume:

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



<h4>2018 Screenshot</h4>

![2018](/Solution/solution_2018.png)

<h4>2019 Screenshot</h4>

![2019](/Solution/solution_2019.png)

<h4>2020 Screenshot</h4>

![2020](/Solution/solution_2020.png)

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

    The code adds the stock's name and initial value to the list. If the stock already exists in teh list, the code aggregates the stock's volume.

<h4>2018</h4>

![2018](/Solution/solution_2018.png)

<h4>2019</h4>

![2019](/Solution/solution_2019.png)

<h4>2020</h4>

![2020](/Solution/solution_2020.png)

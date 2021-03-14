# stocks-analysis

## Overview

Steve has tracked data for 12 different stocks in 2017 and 2018. Each stock an entry for every business day, with each business day having the following data: opening/closing prices, highest/lowest prices, adjusted closing price, and volume. He wants to see the total daily volume and returns for each stock. He can input this information manually, but it would be time consuming, when a macro can do the work for him. Unfortunately, he isn't tech savvy with Excel. Also, he would like to expand on more than just 12 stocks, so he can further research on the stock market, helping his parents decide what to invest.

## Results

### Code

All the data are located in the [spreadsheet Steve provided](./VBA_Challenge.xlsm). The code for the macro has been written within the same file via VBA. To run the macro, Steve first has to input a year, in this case 2017 or 2018. Afterwards, a loop was performed into Steve's data for the corresponding worksheet year. Within the loop, the starting and ending prices for each ticker's respective first and last dates were noted, as well as the total daily volume for each ticker. The main difference between the original code and the refactored code is that the former uses a nested loop while the latter uses an index variable.

### Analysis

After running the macro, within the 12 stocks selected by Steve, [all but one stock have positive returns in 2017.](./resources/Stocks_2017.png) Recommending to invest in these 12 stocks alone would be great, however, [if we were to look at the analysis in 2018].](./resources/Stocks_2018.png), a majority of the stocks have lost in returns. This would most likely prompt Steve to research on a lot more stocks, if not the entire stock market, to help his parents on investing.

## Summary

### Pros and Cons of refactoring code

Refactoring code will make the code more efficient. It removes redundant code and allows for the code to run at a much faster time, which is very helpful for multiple uses. Refactoring code can be easier to maintain. It helps with the readability for other users looking at the code, making it less complex.

The main disadvantages of refactoring code are time consumption and a higher chance of making mistakes. For the former, after the code has been finished, the programmer has to spend additional time refactoring the code, which can lead to time management and/or budget issues. As for the latter, more time will be wasted trying to solve the problem, making it more likely more issues will arise due to the complexity of the code. There are a lot more advantages to refactoring code, however, so it is usually recommended to refactor, if possible.

### Pros and Cons of refactoring the VBA script

Comparing the run times for the [original code](./resources/Module_2_2018.png) versus the [refactored code](./resources/VBA_Challenge_2018.png), it is obvious that the refactored code is significantly faster. This is because the original code uses a nested loop, as stated above. This means the macro would have to loop through thousands of rows multiple times, which would make the process slower if more tickers were added. The only main disadvantage of refactoring is adding more variables and making sure to initialize their values at 0 when grabbing an analysis for a new year.

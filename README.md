# **Stock Analysis Results VBA Challenge**

## OVERVIEW: Stock Analysis in VBA
The analysis of the stock market for years 2017 and 2018, extracting total volume traded by ticker as well as return or yield. The results highlight positive and negative rates of returns.

### Purpose
The goal is to edit or refactor the script worked during module 2, with the intention of seeing an improvement in the time in terms of efficiency. 
*"When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read."* > Bootcamp Washigton University

## RESULTS: Refactoring VBA Code

### Results 2017 vs 2018
Efficiency in time shows below:
- 2017 time variance efficiency between refactored and original script: The refactored script shows improvement because the tickers were indexed before looping through all the stocks.
TIME SAVE: 0.027 seconds

![Run-time 2017 with original script] (https://github.com/ale-river/stock-analysis/blob/main/Resources/Run-time%202017_original.png)

![Run-time 2017 with refactored script] (https://github.com/ale-river/stock-analysis/blob/main/Resources/Run-time%202017_refactored.png)

- 2018 time variance efficiency between refactored and original script: The refactored script shows improvement because the tickers were indexed before looping through all the stocks.
TIME SAVE: 0.058 seconds

![Run-time 2018 with original script] (https://github.com/ale-river/stock-analysis/blob/main/Resources/Run-time_2018_original.png)

![Run-time 2018 with refactored script] (https://github.com/ale-river/stock-analysis/blob/main/Resources/Run-time%202018_refactored.png)

### Comparisson between original and refactored scripts
The first improvement there was defined arrays instead of variables in the refactored script. That reduced the time significantly because the data types used are Single which reduces the time to indicate digits in a number.

ARRAYS & DATA TYPES

![Original Script] (https://github.com/ale-river/stock-analysis/blob/main/Resources/Variables%20defined%20as%20Double.png)

![Refactored Script] (https://github.com/ale-river/stock-analysis/blob/main/Resources/Variables%20defined%20as%20Single.png)

Additionally, the refactored script make efficient to identify the tickers by index them into the for loop.

FOR LOOP INDEXED

![Original Script] (https://github.com/ale-river/stock-analysis/blob/main/Resources/original%20for%20loop.png)

![Refactored Script] (https://github.com/ale-river/stock-analysis/blob/main/Resources/refactored%20for%20loop.png)


## SUMMARY
The script enhances the importants points of the dataset by improving the nested loop and use the tickerIndex as the variable that goes through all the rows. This is makes analysis faster and achieve the goal of efficiency.

### Advantages
The refactoring script is faster than the original one.

### Disadvantages
To verify the syntax of the script can drain some time.

### Pros & Cons of the script
The Pros of writing an script in VBA is that you can be very creative by formatting the output. Microsoft Excel provides many ways to present the information in the spreadsheets.

The Cons of writing an script in VBA is that can make the work tedious because not just the semantic but the syntax in for loops and If-conditionals requires good eye to identify errors.

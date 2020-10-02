# VBA-challenge

### included:
1 - randomStocks.bas script to analyze stock market data 

2 - Multiple_year_stock_data.xlsm file with analysis results

3 - alphabetical_testing.xlsm file (used as test data) with analysis results

4 - Multiple_year_stock_data_CT.xlsm file, yearly data, ready to run; open file then click "Analyze" button 

5 - alphabetical_testing_CT.xlsm file, test data, ready to run; open file then click "Analyze" button

6 - screenshots for each year of my results on the Multi Year Stock Data file: 

    6a. 2014_result.png
    6b. 2015_result.png
    6c. 2016_result.png

7 - this README.md

### How to run:
open the Multiple Year Stock Data file and click "Analyze" button. Analyze button runs the macro randomStocks on every worksheet

### Requirements:
Analyze the stock market data and list all tickers showing Yeary Change, Percent Change, Total Stock Volume.

Based on the summary, return the stock with the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume.

### Assumption: stock data is not sorted

### Algorithm:

Use arrays to store ticker data

initialize arrays with the first record

loop through all rows

if current ticker is different from previous ticker, check if ticker already exists in the array

if ticker exists in the array, get the index for that ticker. If ticker is not in the array, add ticker to array and get the index

get the open date, open price, close date, close price for current ticker

get running total of stock volume per ticker

loop thru the arrays to calculate yearly change and percent change for each ticker

print yearly ticker summary, keeping track of the index of the stock with the greatest % increase, greatest % decrease, greatest total volume

print summary

repeat for all worksheets

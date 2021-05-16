# Stock-analysis

## Overview of Project
The purpose of this project was to perform analysis on various tickers on the Stock Market in the years 2017 and 2018 using VBA and Macros to automate and reduce errors in the calculations. The code was then refactored in order to improve calculation speeds.

## Results

In 2017 there was a positive gain for all tickers except for TERP. In 2018 there was a loss for every stock with the exception of tickers RUN and ENPH.

### 2017 Results
![2017 Results.PNG](https://github.com/crabrandoom/stock-analysis/blob/main/2017%20Results.PNG)
### 2018 Results
![2018 Results.PNG](https://github.com/crabrandoom/stock-analysis/blob/main/2018%20results.PNG)

Once the code was refactored there was about a .02 Second decrease in overall processing time as compared to the original code. Changes that were made to increase the efficiency of this code inlcuded an extra variable to index the Ticker reference in each loop as seen below in the code.

![Code Reference.PNG](https://github.com/crabrandoom/stock-analysis/blob/main/Code%20Reference.PNG)

### Refactored code speeds
![VBA_Challenge_2017.PNG](https://github.com/crabrandoom/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
![VBA_Challenge_2018.PNG](https://github.com/crabrandoom/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

### Original code speeds
![VBA_Challenge_2018' - nonRefactored](https://github.com/crabrandoom/stock-analysis/blob/main/VBA_Challenge_2018'%20-%20nonRefactored.PNG)
![VBA_Challenge_2017 - nonRefactored](https://github.com/crabrandoom/stock-analysis/blob/main/VBA_Challenge_2017%20-%20nonRefactored.PNG)


## Summary
Overall the advatanges of refactoring code are based mostly in computing time. If the data sets are fairly large any decrease in computing time will help save time when performing the analysis. Disadvantages may include extra time spent adapting the code in order to do this.

Specifically on this VBA script the advantages of refactoring were measureable at 0.02 seconds. If the data sets were much bigger than this small amount of time saved would have an even bigger impact. The disadvantages in refactoring this VBA script were miniscule as the amount of code needed did not significantly increase as compared to the time saved from the more efficient code.

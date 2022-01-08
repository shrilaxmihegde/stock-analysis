# Stock Analysis with VBA:
## Overview of Project:
Conducted a stock data analys to help Steve to decide and suggest his Parents to invest in good performing stock ticker. We will be using VBA and Excel to perform this analyss
During the analysis, one of the goal is to create a VBA script which can process the data faster and reprent the data in formatted way.

### Purpose:
 Enable macro through Excel to access the power of VBA. During the VBA scripting used:
  1. Create a ***for*** loop
  2. Using ***if-then*** statement.
  3. Using ***nested loops***
  4. Adding Debug and descriptive comment on code
  5. Using VBA ***conditional formating***
  6. Using ***arrays*** and access the content using ***array index*** 


## Results:
 ![All Stocks (2017)](/Resources/VBA_Challenge_2017.png)
 ![All Stocks (2018)](/Resources/VBA_Challenge_2018.png)
By comparing years 2017 and 2018, we can see, in year 2017, most of the stocks gave good return. In year 2018, only two stocks gave positive return. Based on this data we can conclude ***ENPH*** and ***RUN*** tickers performed well in both years and gave positive return.

With the original code, overall execution time was ~1 sec. With refactored code ( using 3 temp arrays and writing the result once to worksheet at the end), execution time reduced to 0.1 sec. 
 
## Summary

1. What are the advantages or disadvantages of refactoring code?
 One huge upside to refactoring code is making the code to run faster.This can be very easy to work with large datasets. In this case, we did not have an 
exceptional amount of data. But if we are working with a large datasets, then  making sure the program run as quick as possible would be extremely important.
The disavantage is, we should understand the dependency of the modules / code very well and be very  careful while making cahnges. 

2. How do these pros and cons apply to refactoring the original VBA script?
The  bigest benifit, refactoring code was more efficient and code run faster. It decreases the processing time.I did not see any disadvantage in refactoring code.



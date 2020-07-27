# Stock Analysis Project 

## Overview of Project  

The purpose of this project was to create VBA macros to run an analysis on 2017 and 2018 stock data and determine the stocks to invest in for our client as well as to build our client a fast running macro without nested loops that could handle large datasets.  
The data analysis was started by creating the “Stock Analysis” Macro by utilizing nested loops. 
After determining that this type of subroutine would take very long to run if our client wanted to use it on thousands of stock ticker, we decided to code a special macro for our client that uses an array of tickerIndexes instead of nested loops. This opted for a faster running macro. 


## Results 

Comparing our results between 2017 and 2018, it seems that our client should opt to invest in the “RUN” stock since per our data it had the highest return of 84%. The next best option for our client is “ENPH” was a total return of 81.9%.
After building a nested loop macro and a macro with a single variable that would only loop once. the expectation was the single variable macro. 


! [2017 Stock Return Data Results]   https://github.com/jgomez07/Stock-Analysis/blob/master/Resources/VBA_Challenge_2017.png



![2018 Stock Return Data Results]  https://github.com/jgomez07/Stock-Analysis/blob/master/Resources/VBA_Challenge_2018.png

 
## Summary

In summary, there are pros and cons between a nested loop macro and the refactored macro. In my experience, the nested macro was more straightforward and easier to follow the refactored macro. However, even though the refactored macro is more involved, it is worth the memory that the computer saves in running it. Also, the refactored macro works best with very large data sets. 
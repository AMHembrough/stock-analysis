# Stocks Analysis

## Overview of Project

### Purpose & Background 

The purpose of this analysis was to determine the performance of various stocks quickly.  To conduct this analysis, we refactored, or edited, existing code to expand our dataset to include any ticker from the stock market over the past few years, and to increase efficiency of the existing code by reducing script run time.  

### Methodology

To conduct this analysis, we refactored existing code to loop through the stock data a single time to collect all the information, and, in doing so, reducing script time.  To do this, we followed these steps: 
1.	Created a ticker variable called tickerIndex which was leveraged to access the correct tickers and the output arrays, specifically tickerVolumes, tickerStartingPrice and tickerEndingPrice. 
2.	Created a “for loop” that initialized tickerVolumes to zero.
3.	Created a “for loop” that looped over all the rows in the spreadsheet and increased tickerVolumes as based on provided if/then statements.  
	- If the current row is the first row with the tickerIndex, then we assigned the starting price to the tickerStartingPrice variable.
	- If the current row is the last row with the tickerIndex, then we assigned the closing price to the tickerEndingPrice variable. 
	- If the next row’s ticker doesn’t match the previous row’s ticker, then we increased the tickerIndex.  
4.	A “for loop” was used to loop through the four arrays, specifically tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices to return the results of the analysis, specifically, the total daily volume and the associated return for each ticker.    

### Challenges

One of the challenges that I encountered was that I activated the 2017 sheet when I was assigning column headers to the “All Stocks Analysis”.  I saw this error, and made a correction to the code, but I failed to remember this issue until my code wouldn’t work.  After much debugging, I copy/pasted the code into a new workbook and the code worked perfectly.  I believe the error messages I was receiving was because of this mistake I made earlier.  

## Results

### Outcome
Interestingly enough, the refactored code took longer to run than the original code.  I reached out to a learning assistant regarding this matter who was able to help me combine steps 3c and 3d together.  This reduced the run time slightly, but not enough to beat the speed of the original code.  The refactored code takes less than half a second longer to execute than the original code which is marginal and could be explained by other outside variables on my PC..   

### Supporting Documentation
![Run time Original-2017](https://raw.githubusercontent.com/AMHembrough/stock-analysis/main/VBA_Challenge_2017%20original.png) 
![Run time Original-2018](https://raw.githubusercontent.com/AMHembrough/stock-analysis/main/VBA_Challenge_2018%20original.png) 
![Run time Refractored-2017](https://raw.githubusercontent.com/AMHembrough/stock-analysis/main/VBA_Challenge_2017.PNG) 
![Run time Refractored-2018](https://raw.githubusercontent.com/AMHembrough/stock-analysis/main/VBA_Challenge_2018.PNG) 

## Summary

### Advantages and Disadvantages of Refactoring Code in General

Refactoring code is advantageous because it is an opportunity to make code more efficient by reducing the number of steps, using less memory, and/or improving the logic for readability purposes.  

Refactoring code is disadvantageous because it takes time to edit work that was already complete and functioning correctly

### Advantages and Disadvantages of the Original and Refactored code 

An advantage of using the original code is that the functionality worked well, and it was coded in a logical way for someone new to VBA.  A disadvantage of the original code is that it wasn’t the most concise way to reach the end point.  

An advantage of using the refactored VBA script is that is allowed us to think about the code differently, enhancing our knowledge of VBA and showing us there are several ways to get to the end point.  As a result of this challenge, I have learned to be more thoughtful in planning out my solution to ensure that I am being efficient with my time and resources! A disadvantage of using the refactored VBA script was that it took time to edit the original VBA script.  

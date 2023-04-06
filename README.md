# Stock Analysis
## Overview of Project
At the request of my friend Steve, in his effort to assist his parents with their investment portfolio, I was given an Excel file containing stock data covering the years of 2017 and 2018. Steve's parents are focused on the stock of DAQO New Energy Corporation because of its green energy focus, and they want to invest all their money into it. Steve thinks their investments should be more diversified across amongst several green energy stocks, including DAQO stock.

My original objective was to assist Steve's effort by analyzing the data on stock performances for 2017 and 2018, and provide him with the results that he could show his parents; allowing them to make more informed decisions regarding their investment choices. 

Having completed my original objective, Steve is curious to see if the code can be used to analyze the entire stock market dataset, and whether or not it would take too long to execute. I will refactor the code to loop through all the data provided and determine if refactoring the code increased the its efficiency.

## Results
- My original VBA script to capture all DAQO stock was focused on only that single stock, so therefore instead of looping and capturing all stocks the range value was set to "DAQO(Ticker:DQ)":

- ![Range Value Set to DAQO](https://github.com/Caracalla1081/stock-analysis/blob/e3d4db1ccd992b7e90c6939bce33d35bcc2be9a5/Resources/All_Stocks%20_BA_Code2.png)

- The augmented script set the range value to "All Stocks(" + yearValue + ")", this allowes us to capture all stocks in our loops:
![Range Value Set to All Stocks](https://github.com/Caracalla1081/stock-analysis/blob/e3d4db1ccd992b7e90c6939bce33d35bcc2be9a5/Resources/VBA_Challenge%201.png)


- My original VBA script used a single "for loop" with "End If" statements to parse through the stocks looking for "DAQO":
![Original VBA Single For Loop](https://github.com/Caracalla1081/stock-analysis/blob/e3d4db1ccd992b7e90c6939bce33d35bcc2be9a5/Resources/VBA_Challenge%202.png)

- The augmented VBA script established output arrays first before initiating our first "for loop" setting those arrays to "0", 
![Output Arrays](https://github.com/Caracalla1081/stock-analysis/blob/e3d4db1ccd992b7e90c6939bce33d35bcc2be9a5/Resources/VBA_Challenge%203.png)
![Output Arrays Loop](https://github.com/Caracalla1081/stock-analysis/blob/e3d4db1ccd992b7e90c6939bce33d35bcc2be9a5/Resources/VBA_Challenge%204.png)
 
   - and then initiating a second "for loop" that looped through the tickerIndex
![Loop through tickerIndex](https://github.com/Caracalla1081/stock-analysis/blob/e3d4db1ccd992b7e90c6939bce33d35bcc2be9a5/Resources/VBA_Challenge%205.png)

   - ultimately allowing a third "for loop" that returned the performance metrics for all stocks.
![Output Arrays For Loops](https://github.com/Caracalla1081/stock-analysis/blob/e3d4db1ccd992b7e90c6939bce33d35bcc2be9a5/Resources/VBA_Challenge%206.png)



## Summary

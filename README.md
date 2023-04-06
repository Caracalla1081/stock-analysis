# Stock Analysis
## Overview of Project
At the request of my friend Steve, in his effort to assist his parents with their investment portfolio, I was given an Excel file containing stock data covering the years of 2017 and 2018. Steve's parents are focused on the stock of DAQO New Energy Corporation because of its green energy focus, and they want to invest all their money into it. Steve thinks their investments should be more diversified across amongst several green energy stocks, including DAQO stock.

My original objective was to assist Steve's effort by analyzing the data on stock performances for 2017 and 2018, and provide him with the results that he could show his parents; allowing them to make more informed decisions regarding their investment choices. 

Having completed my original objective, Steve is curious to see if the code can be used to analyze the entire stock market dataset, and whether or not it would take too long to execute. I will refactor the code to loop through all the data provided and determine if refactoring the code increased the its efficiency.

## Results
- My original VBA script to capture all DAQO stock was focused on only that single stock, so therefore instead of looping and capturing all stocks the range value was   set to "DAQO(Ticker:DQ)":

![Range Value Set to DAQO](https://github.com/Caracalla1081/stock-analysis/blob/e3d4db1ccd992b7e90c6939bce33d35bcc2be9a5/Resources/All_Stocks%20_BA_Code2.png)

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

- In the end the augmented VBA script perfomed faster than the original VBA script, even though it is collecting more data:
   - Original VBA Script
   ![Original VBA Script Runtime](https://github.com/Caracalla1081/stock-analysis/blob/d049e23990da314c32616af051a015b559702d30/Resources/VBA_Challenge_All_Stocks_2018.png)
   - Augmented VBA Script
   ![Augmented 2017 VBA Script Runtime](https://github.com/Caracalla1081/stock-analysis/blob/d049e23990da314c32616af051a015b559702d30/Resources/VBA_Challenge_2017.png)
   ![Augmented 2018 VBA Script Runtime](https://github.com/Caracalla1081/stock-analysis/blob/d049e23990da314c32616af051a015b559702d30/Resources/VBA_Challenge_2018.png)



## Summary
### Advantages vs. Disadvantages of Refactoring Code
- There are several advantages and disadvantages to refactoring VBA code, and depending on ones level of expertise in working with VBA these may differ amongst users. In this experience I found that some advantages are refactoring code can make code run more efficiently by arranging it in a way that eliminates the need for operations like "for loops" to run more unnecessarily, versus setting a "for loop" to an array and then the arrays are where the "for loops" run within; it can reduce runtime for Macros; and one canan add conditional or formatting code to make the output more accessible for the intended user(s)
- Some disadvantages could be that after augmenting ones code it can be harder for others to work with, or augment themselves if not properly documented. Also, if not documented well, the creator can themselves lose track with changes implemented.

### Advantages vs. Disadvantages of this specific VBA code refactoring
- Regarding this specific project, some advantages and disadvantages that I found were the refactored code provided an advantage in the code's runtime, as the augmented code ran faster than the original, whereas a disadvantage to the refactored code is that for the level of experince I had in working on this project the code could have ran up against deadlines to be completed. More simply put, a disadvantage is one should think about the time and effort it would take to refactor the code, and still accomplish their goal in a timely manner, or is the effort worth the reward.

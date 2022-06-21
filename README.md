# Stock Analysis Using VBA
## Overview of Project 
### Purpose
Steve, my client, is overseeing the investment activities of his parents. They invested in a company called DAQO and Steve wants to compare this stock with others on the market. He is particularly interested in assessing the total daily volume and yearly return of the DAQO stock in 2017 and 2018. He would also like to expand his analysis to include the entire stock mrket over the last few years. 
## Results
### Stock Performance for 2017
To compare and contrast the total daily volume and yearly retutn of DAQO's stock with others, I created a new worksheet called "All Stocks Analysis" to output the data for multiple stocks. I named the new subroutine "AllStocksAnalysis" and formatted the output worksheet to create a title and header. To ensure greater functionality, I created a "yearValue" variable that held an InputBox that asked the user for the year they would like to analyze a given stock's performance. I then initialized an array to hold the 12 tickers of different stocks. I made variables for the starting and ending prices of all stocks. I created a program flow to loop through the array using the index "i" to access each element in the array. 
![all stock](https://user-images.githubusercontent.com/106560739/174704207-001153bb-8500-4450-82e5-30ced25ebb9e.png)
Specifically, I utilized a nested for loop that firstly looped through the tickers and secondly looped through the rows of data to get the total volume, starting price, and ending price for the current ticker. 
![All stock 2](https://user-images.githubusercontent.com/106560739/174704228-3e5d6a26-8d72-481c-a4cd-0890950aa083.png)
I ran my program and entered 2017 as the date in the InputBox. Upon running the code, I found that the DAQO stock (ticker name DQ) had a total daily volume of 35,796,200 and a return of 199.4%. The high volume indicates the stock in 2017 performed with great stability, yet the high return indicates the stock is rather volatile and investment in it can come with higher risks than others. 
![2017 excel snip](https://user-images.githubusercontent.com/106560739/174711075-9601bf29-3659-4289-9614-eb5aedbc0f4a.png)
### Stock Performance for 2018
To analyze the total daily volume and yearly retutn of DAQO's stock with others, I ran the same code and subroutine ""AllStocksAnalysis" as I did for 2017. To ensure I ran the code correctly, I entered the year 2018 when prompted by the InputBox. Upon running the code, I found that the DAQO stock (ticker name DQ) had a total daily volume of 107,873,900 and a return of -62.6%. In 2018, the stock was traded more as displayed by a higher total daily volume. However, the investment shrunk as shown by a negative return value. This indicates that DAQO may not be a suitable company for Steve's parents to invest in.
![2018 excel snip](https://user-images.githubusercontent.com/106560739/174712289-ec58fd5b-0279-4b42-821a-86303fadae79.png)
### Execution Time of Original Script
g
### Execution Time of Refactored Script
## Summary
- What are the advantages or disadvantages of refactoring code?
- How do these pros and cons apply to refactoring the original VBA script?

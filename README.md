# Developing a VBA Script for Stock Analysis
## Overview
Steven’s parents want to invest in green energy, and they choose DAQO New Energy Corp. To help his parents to analyze the stock market and make a sound choice, we help him to develop a VBA module to analyze green energy stocks from 2017 to 2018. Further, to help him use the script to analyze an entire stock market over the last few years, we refactor the script to make it more efficient.
## Results
### Analyzing Green Stock performances
We obtained the stock excel file from Steve for ten green energy stocks showing daily stock opening and closing prices and volume of each company from 2017 and 2018. To analyze and compile thousands of data points, we wrote a VBA code (AllStockAnalysis) to go over all the daily data to create a table showing ticker name of the company, total volume and return percentage of the stocks for both years as shown below,

![2017stockanalysis_table](/Resources/2017_stockanalysis.png)
![2018stockanalysis_table](/Resources/2018_stockanalysis.png)

To show performance of a stock, we calculate the return value by percentage increase of starting close price over ending close price within a year as following,

![Return_calculation_code](/Resources/Return.png)

To visualize and compare stock performances, we formatted return lower than zero as red and over zero as green, therefore, green indicates increase in stock prices and red shows decreased. 

![Format_code](/Resources/Format.png)
2017 Stock analysis showed that majority of the green energy stocks were performing very well, with a increase of stock price from 5.5% to 184.5%, especially SEDB and DQ almost doubled the prices. Only one company has decreased stock price (TERP). However, the stock prices for green energy dropped significantly in 2018 compared to 2017. Stock prices of 80% of the companies decreased by -.5% to -62.6%. Only two stocks with increased prices are ENPH and RUN. It suggests that Steve should help with his parents to diversify stock options instead of investing all in green energy. 

### Refactoring the script
To calculate run time for the stock analysis scrip, we use code (Timer) to calculate run time of each year. The original script takes 0.6796875 seconds to run for 2017 and 0.7226562 seconds for 2018. 

![Runtime_Original_2017](/Resources/Original_2017.png)
![Runtime_Original_2018](/Resources/Original_2018.png)

As Steven wants to expand dataset to the entire stock market over the last few years, he needs a VBA script to run efficiently with large amount of data. To make it more efficient, we refactored the VBA script to make it execute faster. Rather than looping the dataset 12 times,  we use tickerIndex as a variant to create arrays so the script will loop through the whole data only one time to collect the information. The original vs refactored codes for looping through rows are shown below.

![Original_code](/Resources/Original_code.png)
![Refactored_code](/Resources/Refactored_code.png)

The refactored script runs 0.140625 seconds for 2017 and 0.148475 seconds for 2018, almost 5 times faster than the original script. 

![Runtime_Refactored_2017](/Resources/VBA_Challenge_2017.png)
![Runtime_Refactored_2018](/Resources/VBA_Challenge_2018.png)

## Summary
Refactoring code will optimize the script to help it run more efficiently, execute faster, use less memory, and make the code logical and easier to read. However, refactoring code takes manpower and extra time to edit the code. Refactoring process is also prone to making mistakes and takes extra time to debug. For stock analysis, refactored VBA script runs faster than original one, which is beneficial. At the same time, it took us extra time to think of ways to refactor it and debugging the script for it to work, so cons are designing and debugging time. 



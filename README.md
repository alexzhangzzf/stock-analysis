# Developing a VBA Script for Stock Analysis
## Overview
Steven’s parents want to invest in green energy stocks, and they choose DAQO New Energy Corp out of interest. To help his parents to analyze the stock market and make a sound decision, we help him to develop a VBA module to analyze green energy stocks from 2017 to 2018. Further, to help him use the script to analyze an entire stock market over the last few years, we refactor the script to make it more efficient.
## Results
### Analyzing Green Stock Performances
We obtained the stock excel file from Steve for 12 green energy stocks showing daily stock opening and closing prices and volume of each company from 2017 and 2018. To analyze and compile thousands of data points, we wrote a VBA code (AllStockAnalysis) to go over all the daily data and created a table showing ticker name of the stock, total volume and return percentage of the stocks for both years as shown below,

![2017stockanalysis_table](/Resources/2017_stockanalysis.png)
![2018stockanalysis_table](/Resources/2018_stockanalysis.png)

To evaluate stock performances, we calculated the return percentage by comparing stock starting price and ending price (using closing prices)  within a year as following,

![Return_calculation_code](/Resources/Return.png)

To visualize and compare stock performances, we formatted return lower than zero as red and higher than zero as green, therefore, green indicates increase in stock prices and red shows decreased prices. 

![Format_code](/Resources/Format.png)

2017 Stock analysis showed that majority of the green energy stocks were performing very well, with a increase of stock price from 5.5% to 184.5%, especially SEDB and DQ almost doubled the prices. Only one company has decreased stock price (TERP). However, the stock prices for green energy dropped significantly in 2018 compared to 2017. Stock prices for most of the companies decreased by -3.5% to -62.6%. Only two stocks (ENPH and RUN) have increased prices. It suggests that overall stocks for green energy are going down from 2017 to 2018. Steve should help with his parents to diversify stock options instead of investing all in green energy. 

### Refactoring the script
To calculate the run time for the stock analysis script, we used  `Timer` to calculate run time of each year. The original script takes 0.6796875 seconds to run for 2017 and 0.7226562 seconds for 2018. 

![Runtime_Original_2017](/Resources/Original_2017.png)
![Runtime_Original_2018](/Resources/Original_2018.png)

Because Steven wants to expand dataset to the entire stock market over the last few years, he needs a VBA script to run efficiently with large amount of data. To make it more efficient, we refactored the VBA script to make it execute faster. Rather than looping the dataset 12 times,  we use tickerIndex as a variant to create arrays so the script will loop through the whole data only one time to collect the information. The original vs refactored codes for looping through rows are shown below.

![Original_code](/Resources/Original_code.png)
![Refactored_code](/Resources/Refactored_code.png)

The refactored script runs 0.140625 seconds for 2017 and 0.148475 seconds for 2018, approximately 5 times faster than the original script. 

![Runtime_Refactored_2017](/Resources/VBA_Challenge_2017.png)
![Runtime_Refactored_2018](/Resources/VBA_Challenge_2018.png)

## Summary
- Refactoring code will optimize the script to help it run more efficiently, execute faster, use less computer memory, and make the code logical and easier to read. However, refactoring code takes manpower and extra time to pefect the code. Refactoring process also has potential risk for making mistakes, disrupting the original strucurre which will require extra time to trouble shoot and debug. 
- For this stock analysis script, refactored VBA script runs faster than original one, which is beneficial. However, it took us extra time to think of ways to refactor it and during the process it needed debugging script when it didn't work, so cons are extra time for desinging, writing and debugging. 



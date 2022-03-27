# VBA_Challenge
# VBA Challenge: The VBA of Wall Street
##  Background
  This project is to use VBA scripting to analyze generated stock market data. There are two datasets to be analyzed: one is the Test Data which is used while developing the scripts, another is the Stock Data which is the main data to generate the final report. Both of the datasets are real stock market data and are sourced into Excel. The Test Data has five sheets (A-F), and the data set is smaller and can be run faster. The Stock Data has three sheets(2018,2019,2020), and the data size is larger and takes time to run.
## Solution: 
  A loop was created through all the stock data for one year to output the ticker symbol, yearly change, the percent change, the volume of the stock and the greatest of increase change, decrease change and the total volume. 
## Solution 1: The Ticker Symbol
  The distinct Ticker name was sorted and output to Column "I". The conditional script should be used to find the different Ticker symbol.
## Solution 2: Yearly Change
  Find the opening price at the beginning of a given year for each Ticker, and then find its closing price at the end of that year. The difference between the two prices is the value of yearly change. And then display the value on "J" column. 
## Solution 3: The Percent Change
  Calculate the percentage of the value of yearly change for each Ticker in the opening price at the beginning of a given year, and put the percentage into the column "K". For the column of percent change, conditional formatting was used to highlight positive change in green and negative change in red.
##Solution 4: The Greatest of Increase Change, Decrease Change and the Total Volume
  Find the maximum and minimum values of percent change and the maximum value of the total volume, and find their Ticker name at the same time.
## The final solutions for Stock Data (2018-2020) are given below:

<img width="960" alt="screenshot2018bonus_stockdata" src="https://user-images.githubusercontent.com/100816322/160286991-b1410c99-5afa-4109-beaf-49298dab0194.png">

<img width="960" alt="screenshot2019bonus_stockdata" src="https://user-images.githubusercontent.com/100816322/160287000-e6433317-c460-4aab-93da-4366af38b964.png">


<img width="960" alt="Screenshot2020bonus_stockdata" src="https://user-images.githubusercontent.com/100816322/160287003-2caaa86f-7984-4ce3-b74b-3f93e6c8492e.png">


This VBA Code is compiled and written for the homework of Data Analytics Bootcamp class at Georgia Institute of Technology. The dataset generated by Trilogy Education Services, LLC. @2020 Trilogy Education Services, a 2U, Inc, brand. All Rights Reserved.


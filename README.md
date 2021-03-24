# VBA-Challenge - Stock Market Analysis

![Stock_Market](Images/stockmarket.jpg)

## Background

The goal of this project is to summarize the stock market for 2014-2016 using VBA.

The stock market data is sourced into Microsoft Excel and includes three (3) sheets with each year.

## Solution

The script looped through each worksheet and found the number of rows and columns in each worksheet.  Then, we create a summary table as additional columns on each worksheet for the desired outcomes.

###### Ticker Symbol

The script loops through each row in the worksheet and finds each unique ticker value and adds it to the summary table.
	
##### Total Stock Volume

For each ticker symbol, we calculate the total stock volume and add it to Column L.

##### Yearly Change

We calculate both the total change and the percent change in opening stock price at the beginning of a given year to the closing price at the end of the year and add it to Columns J and K respectively.

The script uses conditional formatting to highlight the cells in the 'Yearly Change' column (Column I) in green for positive changes and red for negative changes.

##### Ticker Symbol Overall Summary

For each year, the script loops through each ticker in the summary table and identifies the ticker symbol with the greatest % change, greatest % decrease, and the greatest total stock volume and adds them to Columns O and P.


## Summary Tables

##### 2014

![2014 Stock Summary Tables](Images/Year_2014)

##### 2015

![2015 Stock Summary Tables](Images/Year_2015)

##### 2016

![2016 Stock Summary Tables](Images/Year_2016)






	
	
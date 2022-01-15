# "Green" Stock Analysis
Project designed to allow for quick analysis of green energey stocks. Data provided covers daily results of 12 different green stocks - starting and end prices, daily high and low, and daily traded volume.

Stock analysis tabulates the total daily volume traded, and provides an end of year return. Analysis tools provided to compare individual years.

# Initial Analysis
Based on the data provided, 2018 was a difficult year for green energy companies. Through the year, 10 of 12 stocks lost value, with a deficit net return between -3% and -60%.

There were two shining stars in 2018 however, both netting over 80% return on investment. These two companies also posted a positive return in 2017.

|Ticker|2017 Return|2018 Return|
|---|---|---|
|AY|**8.9%**|_-7.3%_|
|CSIQ|**33.1%**|_-16.3%_|
|DQ|**199.4%**|_-62.6%_|
|ENPH|**129.5%**|**81.9%**|
|FSLR|**101.3%**|_-39.7%_|
|HASI|**25.8%**|_-20.7%_|
|JKS|**53.9%**|_-60.5%_|
|RUN|**5.5%**|**84.0%**|
|SEDG|**184.5%**|_-7.8%_|
|SPWR|**23.1%**|_-44.6%_|
|TERP|_-7.2%_|_-5.0%_|
|VSLR|**50.0%**|_-3.5%_|

### Recommendations
Based on these findings, both stocks **ENPH** and **RUN** show promise going into 2019.  Further investigation into the company profiles, products, and financials would be prudent for investment.

# Behind the Scenes
## Data Overview
Data provided needed little transforming for analysis desired. Data well sorted.

## Pulling the Data
Coded to pull data for individual years.  Code loops through entire worksheet using stock ticker names provided.  Total traded volume is aggretate data from individual days.

Code also finds the first and last instance of a ticker to calculate yearly end return.
```vba
If Cells(i - 1, 1).Value <> tickers(j) And Cells(i, 1).Value = tickers(j) Then
    tickerStartingPrices(j) = Cells(i,6).value
End if

If Cells(i + 1, 1).Value <> tickers(j) And Cells(i, 1).Value = tickers(j) Then
    tickerEndingPrices(j) = Cells(i,6).value
End if
```

# Refactored Code
Refactored code is the idea that initial coding can be reworked and adjusted to simplify the process and make the code quicker to run.

The initial coding process is often a quick and dirty code, a way to "get it down on paper".  Refactoring goes over the process to find updated ways to get the same desired results.

### The Advantages
The biggest advantage of properly refactored code is that it should require less memory and less time to provide the same results, either by eliminating extraneous lines of code, or by finding better flow through the code.

### The Disadvantages
There's no real "disadvantage" to looking over code for possible timesaving areas though the amount of time spent refactoring should be weighed against the possible benefit.

### Real-World Example
In this project, refactoring the code had the advantage of shaving a few microseconds from the process time of the original project. This is likely attributed to switching between worksheets during the run.

The original code found each ticker individually, calculated return and volume, and printed to a new worksheet, the started with the next ticker. The refactored coded stored all of the data as part of an array, only printing to the new worksheet after the full run.

While the outcomes are the same, memory and processing were saved by allowing the code to hold the values, rather than forcing Excel to switch between worksheets multiple times, printing in between each switch.

![2017 Refactored Timesaving](https://i.postimg.cc/FR83FkjL/VBA-Challenge-2017.png) ![2018 Refactored Timesaving](https://i.postimg.cc/Df41j5Td/VBA-Challenge-2018.png)

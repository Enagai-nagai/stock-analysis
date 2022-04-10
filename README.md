# Stock-analysis report


## Overview of project
### Background
Steves's parents are interested in investing in green energy such as hydro, geothermal, wind, and bio energy as they believe that traditional fuels will be less popular in the future.
They are specifically interested in Daqo New Energy Corp (Ticker DQ) and Steve needs to investigate about green energy market for his parents, which is his client.

### Purpose
The purpose of this project is to analyze and compare the green energy stock market, especially the volume of trade, and the performance of Daqo New Energy Corp and other green energy stocks.

### Used Dataset
We used the trade information of the following 12 stocks during 2017 and 2018, which includes the daily volume, and opening and closing prices.

|Ticker|Name|
| :---: | :---: |
|AY|Atlantica Sustainable Infrastructure PLC|
|CSIQ|Canadian Solar Inc|
|DQ|Daqo New Energy Corp|
|ENPH|Enphase Energy Inc|
|FSLR|First Solar Inc|
|HASI|Hannon Armstrong Sustainable Infrastructure Capital Inc|
|JKS|JinkoSolar Holding Co Ltd|
|RUN|Sunrun Inc|
|SEDG|SolarEdge|
|SPWR|SunPower|
|TERP|TerraForm Power, Inc|
|VSLR|Vivint Solar, Inc|

## Results
### Stock Analysis
Here is the overview of the trade amount and performance of each green energy stock.  
Overall, we can see the following trends:
1. In general, most green energy stocks had a positive return in 2017, however, they faced a negative return in 2018.
2. The value of DQ (Steve's parents' interest) tripled its value in 2017, however it decreased by 60% in 2018.
3. In 2017, DQ traded at $3.5 million, the lowest of Green Energy's 12 company shares, but increased to $10 million in 2018  
**2017**  
![image](https://user-images.githubusercontent.com/99149443/162596609-eabf505f-fa4d-466e-ac5d-711cbc9c0893.png)

**2018**  
![image](https://user-images.githubusercontent.com/99149443/162596616-499bc739-5207-45c5-9b34-13b33363f8eb.png)

**Chart: Trade amount of each stock through 2017 and 2018**  
Looking at the trade amount, 
* Many stocks increased their trade volume in June 2018.
* ENPH encountered a huge growth in its trade volume in 2018.
![image](https://user-images.githubusercontent.com/99149443/162596709-a61911a5-5ecf-4495-86d4-b1cbee7af8a6.png)


**Chart: Stock price through 2017 and 2018**  
Looking at the stock price,
* The value of many stocks increased gradually in 2017.
* DQ, along with two other companies, FSLR and SEDG, increased their share prices sharply between September 2017 and January 2018.
* Many companies faced a decrease in their value between May and June 2018.
![image](https://user-images.githubusercontent.com/99149443/162596773-907b4bb0-94b1-4b05-90eb-fdaa529cabc1.png)

### Refactoring related
The VBA code was refactored to improve the visibility and speed of the analysis program.
The speed of the program was shortened as the following table shows.

|Speed|2017|2018|
| :---: | :---: | :---: |
|Original|1.109375|1.171875|
|Refactored|0.265625|0.2539063|

Basically, the difference between the original and the refactored code was the order of extracting the volume, starting price, and ending price of each ticker.
* The original code used a nested loop so the program was operated for all the rows in the 2017/2018 files and was repeated for 12 tickers.
* The refactored code used only one loop by assigning Ticker Index to ticker arrays.  


**Original code**  
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       '5) loop through rows in the data
       Sheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
           '5b) get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                StartingPrice = Cells(j, 6).Value
            End If
           '5c) get ending price for current ticker
            If Cells(j + 1, 1) <> ticker And Cells(j, 1).Value = ticker Then
                EndingPrice = Cells(j, 6).Value
            End If
        Next j
        
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = EndingPrice / StartingPrice - 1  
       
**Refactored code**  
 ' 2a) Create a for loop to initialize the tickerVolumes to Zero
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    
    ' 2b) Loop over all the rows in the spreadsheet
    For i = 2 To RowCount
        ' 3a) Increase volume for curren ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
    
        ' 3b) Check if the current row is the first row with the selected tickerIndex.
        ' If so, assign the closing amount as ticker starting price
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        ' 3c) Check if the current row is the last row with the selected ticker
        ' If so, assign the closing amount as ticker ending price
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d) Increase the ticker Index
            tickerIndex = tickerIndex + 1
        End If
        
    Next i
## Summary
Refactoring code is the common practice of rebuilding code structure for the same purpose to make it cleaner, easier, and more efficient.  
There are several benefits of code refactoring:  
* Reduce less memory to operate the code
* Improve the speed of operating the code
* Simplify the structure and improve the readability of the code to other people  

There are a few disadvantages of refactoring code, however:
* Refactoring requires extra time and work to review the code, which may cause man-hours in real business occasions.

In the case of this stock analysis, 
* The original code has the advantage of (1) not having to assign an index to arrays (2) that you can easily understand the process you are going through (there are too loop so you can see that you are repeating the same process for all the rows and all the tickers.
* The original code, on the other hand, has a disadvantage that (1) It consumes more memory and thus (2) the operation takes longer.
* The advantage of the refactored code is (1) It consumes less memory and thus (2) the operation is faster
* The disadvantage of this code is that (1) it is probably not the simplest code that you can come up so it may be difficult for those who are not used to programming.


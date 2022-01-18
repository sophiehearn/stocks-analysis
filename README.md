# Stocks Analysis

## Overview of the Project 

Steve's parents are looking for good investments in the stock market. Steve has data for 12 green stocks, which we helped him analyze. He's interested in looking at the greater stock market, however, so he's wondering if there's a more efficient way to run the code than what we have already compiled. To help with this, we are refactoring the code and timing how long it takes to run an analysis with the different codes on the same data. 
This will help us understand whether the new refactored code is superior to the original code.

## Results

To attempt to improve the efficiency of the processing, the code was edited so that the program only loops through the dataset once, rather than individually for every stock.

The original code to check and output the total volume, starting price, and ending price of a stock looked like this: 
,,,
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
,,,


## Summary


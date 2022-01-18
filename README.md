# Stocks Analysis

## Overview of the Project 

Steve's parents are looking for good investments in the stock market. Steve has data for 12 green stocks, which we helped him analyze. He's interested in looking at the greater stock market, however, so he's wondering if there's a more efficient way to run the code than what we have already compiled. To help with this, we are refactoring the code and timing how long it takes to run an analysis with the different codes on the same data. 
This will help us understand whether the new refactored code is superior to the original code.

## Results

To attempt to improve the efficiency of the processing, the code was edited so that the program uses fewer loops and is able to track each stock in an array. 

The original code used the below nested for loop to check and output the total volume, starting price, and ending price of a stock looked like this: 
```
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0

	'Loop through rows in the data

       Worksheets(yearValue).Activate
       For j = 2 To RowCount
	

	'Get total volume for current ticker
         
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
          
	'Get starting price for current ticker
         
	 If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

         'Get ending price for current ticker
           
	If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       
	'Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
```

The refactored code instead used output arrays to store the data from each ticker. Initially the output arrays were created: 
``` 
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
Then a for loop was created to set the tickerVolumes to 0 initially: 
```
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
```
Finally, the for loop was created to have ticker volumes, ticker start price, and ticket end price change their values and update in the array.

```
  For i = 2 To RowCount
    
        ' Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        ' Check if the current row is the first row with the selected tickerIndex.
            
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         
         End If
        
        ' Check if the current row is the last row with the selected ticker
        ' If the next row’s ticker doesn’t match, increase the tickerIndex.
        
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
         End If

         ' Increase the tickerIndex.
            
         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                   
         tickerIndex = tickerIndex + 1
         
         End If
    
    Next i
    
    ' Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```

This refactored code was tested against the original code for both the 2017 and 2018 data sets. 

Original code: 
![Original2017](https://raw.githubusercontent.com/sophiehearn/stocks-analysis/main/Resources/original2017.png)
![Original2018](https://raw.githubusercontent.com/sophiehearn/stocks-analysis/main/Resources/original2018.png)

Refactored code:
![Refactored2017](https://raw.githubusercontent.com/sophiehearn/stocks-analysis/main/Resources/refactored2017.png)
![Refactored2018](https://raw.githubusercontent.com/sophiehearn/stocks-analysis/main/Resources/refactored2018.png)

## Summary


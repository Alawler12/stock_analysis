# VBA of Wallstreet

## Project Overview
### Refactoring previously written code
Throughout the lesson on VBA, we were tasked with writing a VBA script that would analyze two years of data on green stocks available for trading on Wallstreet in the green_stocks.xlsm file.  The script searched though the data by year and returned the volume of each stock traded and determined the percentage of return or loss over the course of trading. The purpose of this final challenge was to take that previous script and to refactor it to be more efficient.  The hope is that this code would run more quickly than the original and could be applied to a larger dataset in future.

## Results 
### Stock Performance
In terms of stock performance, almost all of them had a better year in 2017 than 2018.  TERP had a 2% increase in 2018, but it was still in the negative overall.  RUN and ENPH were the only stocks to have positive returns in 2018, however ENPH did worse than the previous year, while RUN went from 6% to 84%.  DQ fell from a 199% return to a -63% return.  So, RUN looks to be the stock with the best return rate, and DQ the worst.

![VBA_Challenge_2017.PNG](https://github.com/Alawler12/stock_analysis/blob/master/VBA_Challenge_2017.PNG)

![VBA_Challenge_2018.PNG](https://github.com/Alawler12/stock_analysis/blob/master/VBA_Challenge_2018.PNG)

### Code Performance
In relation to run times for the old code versus the refactored code, the refactored code also performed better.  The refactored code ran about 4 -5 times faster than the old code.  The refactored code ran at .094 seconds for 2017 data analysis and .086 seconds for 2018, as opposed to .417s and .410s respectively. 

The old code looped through each individual row of the worksheet and returned the volume, start price, and end price for each stock ticker and then output it to the output worksheet before moving on to loop through each row again for the next ticker: 
```
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
            
            'loop through rows of data
            Worksheets(yearValue).Activate
            
            For j = 2 To RowCount
            
                'find total volume for current ticker
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
                
                'find starting price for current ticker
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
                
                'find ending price for current ticker
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
                
            Next j
            
            'output data for current ticker
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
        Next i
```


The refactored code, on the other hand, assigned these same data outputs to defined arrays, seen below:
```
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
Then a separate loop was used to output the data from the array to the output worksheet:
```
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```
**This is more efficient because the code stays within itself until completion instead of jumping back and forth to the output worksheet for each run of a loop.**  The more complete code can be seen below:
```
  '1a) Create a ticker Index
    Dim tickerIndex As Integer
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

  '2a) Initialize ticker volumes to zero
    For i = 0 To 11
        tickerIndex = i
        tickerVolumes(tickerIndex) = 0
    
    Next i
        
        '2b) loop over all the rows
        tickerIndex = 0
        For j = 2 To RowCount
        
    
            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1) = tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
            End If
            
            '3c) check if the current row is the last row with the selected ticker
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
        
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            
            End If
        Next j
           
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```

### Summary of Findings 
This challenge project shows that there are advantages and disadvantages to refactoring code.  An advantage is that it greatly increases the efficiency of the code and uses fewer computational resources.  A disadvantage of refactoring code may be the old adage of “if it ain’t broke, don’t fix it.”  

This applies to the refactoring of the VBA script as well.  The efficiency of the refactored code is an obvious improvement. The project dataset is relatively small, and the extra steps present in the old code would add up to much higher run times if the dataset were larger.  The refactored code also requires less of the computer and therefore reduces the probability of computational errors or equipment failure.  But on the other hand, the old code worked in less than half a second, so it was hardly a failure.

Overall, refactoring the code was a valuable exercise in seeing many solutions to the same problem, and provided a greater understanding of both the old and new code by comparing them to each other. 

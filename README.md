# Excel VBA Stock Analysis
## Overview of Project
### Purpose
A client named Steve had requested a stock analysis to be done on 12 green energy stocks with which he provided the corresponding datasets for. The Excel VBA script created was able to accomplish the task; however, Steve would then request to see if the script could be run on a larger dataset such as the entire stock market. Since the script was not written to be optimized for such a large database, the script would need to be refractored.
### Background
The original Excel VBA script that was made for the stock analysis requested only took around 0.7 seconds to analyze each year for the 12 stocks requested. However, if a large dataset such as the entire stock market needed to be analyzed, this would lead to the script taking an extended amount of time to run. By refractoring the original script to become more efficient in memory usage and thus reducing the run time of the script, the analysis of a dataset as large as the entire stock market would run much quicker.
### Results
```
#### Original Script
For i = 0 To 11
        
        
        ticker = tickers(i)
        
        totalVolume = 0
        
        Worksheets(yearValue).Activate
        
        For j = 2 To RowCount
        If Cells(j, 1).Value = ticker Then
        
        
            totalVolume = totalVolume + Cells(j, 8).Value
            
        End If
        
        'Checking starting price
        If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
            
            'set starting price
            startingPrice = Cells(j, 6).Value
            
            
            
        End If
        
        'Checking ending price
        If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            
            'set ending price
            endingPrice = Cells(j, 6).Value
            
        End If
```
#### Refractored Script
```
'1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row ticker doesn't match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerIndex = tickerIndex + 1
            End If
            
        'End If
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
``` 
### Summary
#### Advantages and Disadvantages
![VBA_Challenge_2017.PNG](https://github.com/tommy-chin/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
![VBA_Challenge_2018.PNG](https://github.com/tommy-chin/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

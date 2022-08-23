# Excel VBA Stock Analysis
## Overview of Project
### Purpose
A client named Steve had requested a stock analysis to be done on 12 green energy stocks with which he provided the corresponding datasets for. The Excel VBA script created was able to accomplish the task; however, Steve would then request to see if the script could be run on a larger dataset such as the entire stock market. Since the script was not written to be optimized for such a large database, the script would need to be refactored.
### Background
The original Excel VBA script that was made for the stock analysis requested only took around 0.7 seconds to analyze each year for the 12 stocks requested. However, if a large dataset such as the entire stock market needed to be analyzed, this would lead to the script taking an extended amount of time to run. By refactoring the original script to become more efficient in memory usage and thus reducing the run time of the script, the analysis of a dataset as large as the entire stock market would run much quicker.
### Results
In the original script, a nested for loop alongside if statements were used to calculate the total volume, starting prices, and ending prices of the stocks. These values were then individually outputted into the "All Stocks Analysis" worksheet as it went through the entire nested for loop. In the refactored script, a nested for loop was not used. Instead, three output arrays were created which would hold the total volumes, starting prices, and ending prices of each stock. Then, a single for loop with if statements was used to loop all of the rows in the dataset which stored all of the values needed in each individual array. Another for loop which looped through each array was used to output the values into the "All Stocks Analysis" worksheet. With the refactored script, the run time was able to go from around 0.7 seconds for each year analyzed to 0.08984 seconds for 2017 and 0.09766 seconds for 2018. 
#### Original Script
```
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
        
        Next j
    
    'Output results
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    
    Cells(4 + i, 2).Value = totalVolume
    
    Cells(4 + i, 3).Value = startingPrice
    
    Cells(4 + i, 4).Value = endingPrice
    
    Cells(4 + i, 5).Value = endingPrice / startingPrice - 1
    
    Next i
    
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
#### Refractored Run Times for 2017 and 2018
![VBA_Challenge_2017.PNG](https://github.com/tommy-chin/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
![VBA_Challenge_2018.PNG](https://github.com/tommy-chin/stock-analysis/blob/main/VBA_Challenge_2018.PNG)
### Summary
#### Advantages and Disadvantages of Refactoring
Refactoring a script may lead to a much cleaner script not just visually but also practically. By refactoring, a script can be improved by reducing the amount of actions that need be performed which in turn would lead to less memory usage for the script. An original script can often times be a rough draft that is just written to accomplish a task on a small scale. However, the performance of the script may not perform as well on a larger scale so refactoring can lead to the script being able to do so more efficiently. However, refactoring a script may not always be optimal. It is possible that the original script is fairly optimal for the task it was trying to accomplish. Attempting to refactor the script may lead to unexpected bugs which would lead to time wasted in debugging. 
#### Advantages and Disadvantages of the Original and Refactored VBA scripts
As seen in the refactored run times, the refactored script was able to drastically improve the run times of the original script. These improved run times were due to less memory being used in the refactored script. If a dataset as large as the entire stock market was analyzed using the original script, it is possible that due to the amount of memory required to run the script, Excel could potentially crash during analysis. Since the refactored script was able to run the analysis nearly 10x quicker than the original, the chance of Excel crashing would be smaller. 

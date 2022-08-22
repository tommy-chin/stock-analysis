# Excel VBA Stock Analysis
## Overview of Project
### Purpose
A client named Steve had requested a stock analysis to be done on 12 green energy stocks with which he provided the corresponding datasets for. The Excel VBA script created was able to accomplish the task; however, Steve would then request to see if the script could be run on a larger dataset such as the entire stock market. Since the script was not written to be optimized for such a large database, the script would need to be refractored.
### Background
The original Excel VBA script that was made for the stock analysis requested only took around 0.7 seconds to analyze each year for the 12 stocks requested. However, if a large dataset such as the entire stock market needed to be analyzed, this would lead to the script taking an extended amount of time to run. By refractoring the original script to become more efficient in memory usage and thus reducing the run time of the script, the analysis of a dataset as large as the entire stock market would run much quicker.
### Results
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
```
### Summary
#### Advantages and Disadvantages
![VBA_Challenge_2017.PNG](https://github.com/tommy-chin/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
![VBA_Challenge_2018.PNG](https://github.com/tommy-chin/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

# Stock Analysis

# Overview of Project
The purpose of our analysis was to create a VBA script capable of analyzing a set of stock data. We wanted this script to be easy to use for our client. We ran our analysis using two VBA scripts, one that we created in our module and another that was refactored to enhance performance and accuracy. Having two scripts gives us a chance to compare the pros and cons of the codes.

# Results
The stock performance for 2017 and 2018 were much different. 

![stocksresults2017](https://user-images.githubusercontent.com/82848585/131198623-06b1dc05-4165-4ffd-8905-c7680d2240c2.png)![stockresults2018](https://user-images.githubusercontent.com/82848585/131198633-200a6683-352e-44f5-bd39-b800b8733bad.png)


In 2017 almost all of the stocks had a positive return. The best performing stock was DQ with near 200% return and the worst was TERP with a negative 7.2% return. The same positive notes cannot be said for 2018 where all of the stocks except two showed a negative return. RUN was the best performing stock with an 84% return and JKS had a large negative return of -60.5%. ENPH performed the best over both years with returns of 130% and 82% respectively. 

# Script Comparisons
We ran our analysis through two VBA scripts, one that was created during the module and another that we 'refactored' for the challenge. The goal with the refactored script was to increase efficiency by decreasing processing time.


In this original script it goes through all the data 12 times using nested loops with the variables 'j' and 'i'. 
```
'Loop over the tickers array
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        'Loop over the data
        Worksheets(yearValue).Activate
        For j = 2 To rowEnd
            'totalVolume for the current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            'startingPrice for the current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            'endingPrice for the current ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j
        'Output results
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
```

In this refactored script we used arrays for the results allowing the script to go through all of our data rows only one time. 
```
'1a) Create a ticker Index
    Dim tickerIndex As Integer

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Initialize ticker volumes to zero
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
    Next tickerIndex
    're-initialize tickerIndex to zero before looping over all rows
    tickerIndex = 0
        
    '2b) loop over all the rows
    For i = 2 To RowCount
         
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            'starting price for the current ticker
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            'ending price for the current ticker
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
        End If
        
        '3d) Increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```
The runtime for the original code and refactored code differed by almost 0.7 seconds. This is quite a signiciant boost in efficiency. 
![originalcoderuntime](https://user-images.githubusercontent.com/82848585/131199329-b638b1b3-4ff8-42c4-91d8-9a1ad5e3eaa9.png)![refactoredcoderuntime](https://user-images.githubusercontent.com/82848585/131199333-0c766353-4886-438e-8d18-860fc51cc550.png)

# Summary
1. advantages and disadvantages of refactored code
Refactoring code is in general a positive. Obviously, it makes our code faster, but it also provides a cleaner code for future editing and reduces the chance of bugs. Some potential negatives would be that it could take more man power to go through and edit a code so it could be important to decide if the time spent in refactoring a code would be made up later by the codes efficiency or ease of editing. 
2. advantages and disadvantages of the VBA script modification
Our refactoring of this VBA code was positive. The time it takes to run has improved and the code is easier to understand. The disadvantages of the code have to do with both the original and the refactored code in that the functionality is still the same and comes with the same limitations - it only analyzes these twelve stocks. The VBA scripts we worked with would take extensive editing to apply to any set of stocks or to be generally usable for a wider range of stock analysis.

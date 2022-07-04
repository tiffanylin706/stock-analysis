# Stock-Analysis

## Overview of Project
### Purpose
The purpose of this project was to refactor a VBA code to collect certain stock information in both year 2017 and 2018 and determine whether or not the stocks are worth investing.

### The Data
The data that is presented includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

## Results
### Analysis
I created the input box, chart headers, ticker array, and to activate the specific worksheet. The steps were then listed out in order to set the structure for the refactoring. The instruction and comments were written in the file below. 

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
        For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(j, 1).Value = tickerIndex Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
            
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
    End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
     End If

            '3d Increase the tickerIndex.
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If
         
        Next j
            
        'End If
    
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For x = 0 To 11
       Worksheets("All Stocks Analysis").Activate
        tickerIndex = x
        
        Cells(4 + x, 1).Value = tickers
        Cells(4 + x, 2).Value = tickerVolumes(x)
        Cells(4 + x, 3).Value = tickerEndingPrices(x) / tickerStartingPrices(x) - 1
     
        
    Next x

In 2017, the stocks mostly generated a positive return, with DQ having the highest Return Rate. Only TERP generated a negative return.

![VBA 2017 Screenshot](https://github.com/tiffanylin706/stock-analysis/blob/c5681e82b6a90e8d141c382521ae4750e0fcf22b/Resources/VBA_Challenge_2017.png)

![VBA 2018 Screenshot](https://github.com/tiffanylin706/stock-analysis/blob/c5681e82b6a90e8d141c382521ae4750e0fcf22b/Resources/VBA_Challenge_2018.png)

## Summary
### Pros and Cons of Refactoring Code in General
**Pros:** After refactoring, the code is fresher, easier to understand or read, less complex and easier to maintain. Refactoring helps make our code more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming. 


**Cons:** Time consuming. Sometimes, applications are too large and it took more time than we expected. 

### Pros and Cons application to refactoring the original VBA script
 After refactoring, I got a simplified version of the original VBA script, and making it more presentable and organized.
 However, a disadvantage could be that the steps were too lengthy, and it could increase the complexity.

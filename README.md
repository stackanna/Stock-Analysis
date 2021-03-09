# Stock-Analysis with Excel VBA

Click [HERE](https://github.com/stackanna/Stock-Analysis/blob/main/VBA_Challenge.xlsm) to view [VBA Challenge Worksheet](https://github.com/stackanna/Stock-Analysis/blob/main/VBA_Challenge.xlsm)

## Overview of Project

### Purpose

The purpose of this analysis was to refactor a Microsoft Excel VBA code to retrieve specific information pertaining to stocks. Make a more efficient VBA code to analyze data from the stock market for specific years. This information and stock analysis was to be used to determine the potential of investment. The VBA code was given in a simplified form and the intention of this analysis was to increase the effectiveness when running its program. 

### Background

The data presented below is information on 12 separate stocks during the years of 2017 & 2018. 
The information on the stocks analysis include: the ticker value, opening/closing price, lowest/highest price, and the volume of the stock. The overall analysis performed was to compute the ticker, total daily volume and financial return on each of the dozen stocks.
![alt text](https://github.com/stackanna/Stock-Analysis/blob/main/2017..png)
![alt text](https://github.com/stackanna/Stock-Analysis/blob/main/2018..png)

## Results Analysis

Here I have provided an example of the copied code with the required alterations to run the VBA

    1a) Create a ticker Index
    tickerIndex = 0

    1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    2a) Create a for loop to initialize the tickerVolumes to zero.
    If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    Next i

    2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

    3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    3b) Check if the current row is the first row with the selected tickerIndex.
    If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    3c) check if the current row is the last row with the selected ticker
    If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

    Next i

    4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
   
    Next i

 By comparing the performances from 2017 & 2018 stocks we are able to conclude that 2018 would have been a negative return for most of the stocks except ENPH & RUN. The significantly better financially fiscal year would be for 2017 as all but for TERP had a positive return.
 The returns over 100% on the investors money would have been in year 2017 for stocks DQ, ENPH, FLSR & SEDG. 
 Overall the only two stocks that had a positive return for both years included ENPH & RUN. 
 If I was the investor I would suggest RUN & ENPH as their growth seems steady and return remains positive.

## The advantages of refactoring Stock Analysis.

    The advantages of refactoring the VBA code for stock analysis led us to a significantly faster macro code run time. We are able to efficiently include information that allows us to more effectively analyze the stocks value. 

## Pros & Cons of Refactoring Code

    Refactoring the original VBA script is an absolute advantage while making it easier for the individual analyzing the code to interpret and understand what they are trying to accomplish. 
![alt text](https://github.com/stackanna/Stock-Analysis/blob/main/2017.png)
![alt text](https://github.com/stackanna/Stock-Analysis/blob/main/2018.png)

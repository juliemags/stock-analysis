# stock-analysis

## Overview
The purpose of our project is to refactor our VBA Code to cycle through data in an attempt to collect our data faster than our original method (not refactoring). We are trying to determine the value of 12 stocks from 2017 and 2018 by looking at a large amount of data at one time, rather than looking at each data point, which would be inefficient and slow. 

## Results
We analyzed 12 stocks in 2 different sheets of data, each tied to a different year (2017 and 2018). Each stock had the high value, low value, closing value, adjusted closing value, and volume for each day in the year. This information all comes together to help determine whether Steve should buy any of these stocks or not. 

Using this data, we refactored VBA code to hopefully create quicker results when looking at larger sets of data. After the user inputs the desired year of analysis, the intended output of our data would be in a separate "All Stocks Analysis" tab that posts the Ticker, the Total Daily Volume, and the Return % for the indicated year. We also timed the code to determine how quickly the program would run from start to finish. The code is included below: 
```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
      
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
           tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
           End If
        
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

## Summary
From this analysis, we determined that only one ticker (**TERP**, -7.2%) had a negative return in 2017. Each of the other tickers had a positive return in 2017. In the next year, the only tickers with a positive return in 2018 were **ENPH** (81.9%) and **RUN** (84.0%). Every other ticker had a negative return on the year.  

### Advantages & Disadvantages
One huge advantage to this method is the speed for our current data table. Each year had a run time of less than 0.5 seconds. The disadvantage of this method is that the run time can take a lot longer with larger datasets. See below for screenshots of the runtimes. 
![2017 Runtime](https://github.com/juliemags/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
![2018 Runtime](https://github.com/juliemags/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

### Pros and Cons to Refactoring VBA
One pro to refactoring is that the runtime is shorter than any more manual coding method. An experienced coder will be able to loop through the data quickly without many hiccups. One con is that the code is difficult for people with little organization skills or experience coding, as the for and if loops can cause a novice coder to be very confused. 

# Module 2: Stock Analysis - VBA Challenge

## Project Overview

The purpose of this project was to perform an analysis on a stock market excel workbook using VBA. We had previously written out code to analyze the data for any year, including finding the total volume and returns and conditionally formatting them. Now for this part, we refactored our previous code to be more efficient and timely. While the results looked the same (which means it was correct), the code runs faster. I also found this code was easier to understand and follow process wise and made more sense logically. 
## Results

As previously mentioned, the end results between the original and refactored were the same. However, we were able to get faster results using the following. 

> Sub AllStocksAnalysisRefactored()

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
    
    tickers(0) = “AY”
    tickers(1) = “CSIQ”
    tickers(2) = “DQ”
    tickers(3) = “ENPH”
    tickers(4) = “FSLR”
    tickers(5) = “HASI”
    tickers(6) = “JKS”
    tickers(7) = “RUN”
    tickers(8) = “SEDG”
    tickers(9) = “SPWR”
    tickers(10) = “TERP”
    tickers(11) = “VSLR”
    
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
    
     '2a) Create a for loop to initialize the tickerVolumes to zero.

        For i = 0 To 11
        ticker = tickers(i)
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    
        Next i
   
     '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
    
       '3a) Increase volume for current ticker
    
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
    End If
    
    
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            'If  Then
         
       If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
           End If
            
        '3d Increase the tickerIndex.
        
        tickerIndex = tickerIndex + 1
        
        'End If
     
        Next i
     
     '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i).Value = tickers(i)
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

## Summary

The benefits of refactoring is having cleaner, more organized code that is easier to follow and contains fewer individual steps. I have found that the fewer the steps, the more I can limit potential mistakes. And the refactored code runs more efficiently and therefore finishes faster than the previous code.
I would say the main disadvantage of refactoring is time. While I had plenty of time to finish this particular project, circumstances are different in a real professional setting. You might not always have time to refactor if you are hitting up against a deadline. There is also a risk that if your code is already working correctly, you might get new errors while refactoring. 

# VBA_Challenge

Overview of Project: The purpose of this analysis was to present dtate of the stoick market for the last few years to allow Steve and his parents to easily track volume in the stock market as an indicator of accurate trading prices.

Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

By using the developer function in excel and VBA, I created an all stock analysis to analyze the entry price (EP), sale price (SP), and Volume using the For Loop construct

For i = 0 To 11
        tickerVol(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVol(tickerIndex) = tickerVol(tickerIndex) + Cells(i, "H").Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i - 1, "A").Value <> tickers(tickerIndex) Then
            tickerSP(tickerIndex) = Cells(i, "F").Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        If Cells(i + 1, "A").Value <> tickers(tickerIndex) Then
            tickerEP(tickerIndex) = Cells(i, "F").Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    Next i


Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, "A").Value = tickers(i)
        Cells(4 + i, "B").Value = tickerVol(i)
        Cells(4 + i, "C").Value = tickerEP(i) / tickerSP(i) - 1
        





Summary: The advantages or disadvantages of refactoring code are that it eliminates sorting and confusing forumlas on one sheet and allows analysis on a seperate sheet. It also places a simple button to easily move back and forth. Basically, it simplifies the data presentation.However, the coding is time consuming and must one must be diligent in building the code. 

How do these pros and cons apply to refactoring the original VBA script? One must take care not to alter the origianl VBA script in a manner that alters the data.

# Stock_Analysis

# OVERVIEW

	Steve was impressed with our work in the module that he wanted to complete some additional work on stocks for 2017 and 2018. Our objective is to refactor our code to loop through the data in order to determine if the VBA script runs faster than it did in the module. Steve wants a written analysis of our findings.

# THE DATA

	The data used in this analysis is stock information for 12 different stocks. The stock information included the ticker, date, open & close price, highest & lowest price, adjusted closing price, and volume of the stock. The results we are looking for are the ticker, the total daily volume and the return on each of the 12 stocks.

# THE RESULTS

	The goal is to refactor the code, but before completing that I had to copy the starter code in order to create the ticker array, produce chart headers and activate the correct worksheet. The steps to complete this are listed below:

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
     
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 3).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            

        '3d) Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
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

After running the code the results were as follows:

![alt text](VBA_Challenge_2017.png)

![alt text](VBA_Challenge_2018.png)



#SUMMARY

The completion of refacting our code makes the data cleaner and more organized. This can result in faster debugging and faster programming. Another advantage is that it makes our clients more likely to understand the data clearer and be able to translate what is happening within the code. Unfortunately there are possible disadvatages. There can be errors in the data that can make it difficult to refactor the code. There might be applications or data sets that are too large to retrieve a concise result of the data.

Refactoring the code resulted in a decrease in macro run time. The analysis during the module took about a minute to run, whereas the macro run time during the refactor took about 1/4 of the time. The results of each refactor macro run time are shown below:
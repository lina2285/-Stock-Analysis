# Stock-Analysis
Excel file of Analysis [VBA Challenge - Stock Analysis](https://github.com/lina2285/-Stock-Analysis/blob/main/VBA_Challenge.xlsm.xlsm)

##Overview of Project
###Purpose

The most recent stock analysis project aimed to allow Steve to collect information on all stocks over 2017 and 2018. The information provided will assist in deciding which stocks his parents should consider for investment. There were some limitations in accessing all the data in the previous dataset provided, but by refactoring the original code, we can now offer more efficiency in the analysis.  

###Results
##Data anlysis result
The results of the data collected showed the results for both 2017 and 2018. The charts illustrates Ticker name, total Daily Volume and the yearly return for each. on the Return column, the data was specifically conditioned to highlight the negative returns in red and the positive returns in green. Having such data displayed gives Steve a detailed summary to easily separate the stocks that are doing better. In 2017, most stocks had great performace with the exception of "TERP". In 2018, the stocks "ENPH" and "RUN" were the only two stocks that had a positive a return.  

##Analysis process

When refactoring the code, I was able to resuse some of the previous code to set up the sheet similarly to the previous dataset provided. Once the table was set up I moved on to refactoring.  Below are the steps with detail of what what being achieved in every step. 

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
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
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
    
###Summary
##Advantages of refactoring code

Refactoring code has advantages that mainly assist in making the code cleaner, more organized and easier to follow for oneself and for others looking at the code. Refactoring improves the desing, makes it easier to understand, and easier to maintain. In the case of this project, it improved the macro run time. The new code took about 0.27 seconds to run, where the old code took about one full second. 
[VBA_Challenge_2017](https://github.com/lina2285/-Stock-Analysis/blob/main/VBA_Challenge_2017.png)

[VBA_Challenge_2018](https://github.com/lina2285/-Stock-Analysis/blob/main/VBA_Challenge_2018.png)


##Disadvatages of refactoring code

Refactoring code also has its disadvantages, which include it being very time consuming. It is also very easy to make mistakes and it takes very long to find and fix them. 

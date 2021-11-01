# Stock Analysis with VBA

## Overview of Project
Steve's parents are passionate about green energy. Therefore, they decided to invest in Daqo New Energy Corporation, which makes silicon wafers for solar panels. Steve is going to look into Daqo for his parents. He also wants to analyze a handful of green energy stocks in addition to Daqo's stock.

### Purpose
The purpose of this project is to refactor the code we have built. This is because Steve wants to expand the dataset to include the entire stock market over the last few years. Although the code works, it might not work as well for thousands of stocks. We are going to refactor the code and present an analysis and findings.<br>
<br>
The excel workbook is located [here](https://github.com/Takomochi/stock-analysis/blob/main/VBA_Challenge.xlsm). 

## Results
### Analysis
1. Refactoring Code<br>
To refactor code, we loop through the data one time and collect all the information. First set tickerIndex as zero. Then, created three output arrays, tickerVolumes, tickerStartingPrices and tickerEndingPrices. Loop through all the rows and store values (tickerVolumes, tickerStartingPrices, and tickerEndingPrices) for each ticker. We used the IF-THEN statement to get tickerStartingPrices and tickerEndingPrices. Finally, loop through all the arrays(tickers, tickerVolumes,tickerStartingPrices, and tickerEndingPrices) to output the ticker, total daily volume, and return. <br>

- Create arrays
'''
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
'''
<br>

- Store tickerVolumes inside the loop
'''
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
'''
<br>

- Store tickerStartingPrices and inside the loop
'''
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
End If
'''
<br>

- Store tickerEndingPrices and Increase tickerIndex inside the loop
'''
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

tickerIndex = tickerIndex + 1
            
End If
'''

<br>

- Store tiPrices and inside the loop
'''
For i = 0 To 11
        
Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = tickers(i)
Cells(4 + i, 2).Value = tickerVolumes(i)
Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

Next i
'''
<br>

2. Stock performance between 2017 and 2018<br>
Between 2017 and 2018, the stock performance was much better in 2017. <br>
In 2018, most of the stocks performed negatively. The only stocks that kept positive returns were ENPH and RUN in 2018.

<br>

3. Execution times of the original code and the refactored code<br>
Run time became much faster for both years with the refactored code, as shown in the images.<br>



## Summary

- What are the advantages or disadvantages of refactoring code?<br>
    - One of the advantages of refactoring code is making code more straightforward to understand. Furthermore, the code runs much faster, which is suitable for the more extensive dataset.<br>
    <br>
    - The disadvantage of refactoring code is time-consuming. It requires reconstructing the code. It needs to be appropriately planed before refactoring the code.<br>
    
- How do these pros and cons apply to refactoring the original VBA script?<br>




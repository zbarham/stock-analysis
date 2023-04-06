# Stock Analysis with VBA

## Overview of Project
The client Steve is happe with the workbook that has been prepared for him but he wants to expand the dataset to from a small dataset to a large one. In addition to building a script to process more data, the ask is to refactor the code and ensure that it is as efficient as possible.

### Purpose
The purpose of this analysis is to refactor the VBA code given to loop through all of the data at one time, collecting the same information from the stock market data based on the year input into the script. This exercise was done to ditermine if refactoring the code would reduce the time taken for the VBA script to run.

## Results
The refactored code successfully produced the same output as the original script, with the ability to analyze the entire stock market. The refactored code was able to run faster than the original script, the execution time was significantly reduced being less than a quarter of the original time. The original runtime for analyzing the 2017 stocks was 0.9228516 seconds while the refactored code was only 0.1132813, 2018 was similarly improved with the original 0.9228516 seconds being reduced to 0.1152344 seconds after refactoring. An analysis of the stocks between 2017 and 2018 shows that while 2017 was a very strong year for almost every stock the trend did not follow into 2018 with all but two having a loss in the end, the same result we had prior to refactoring but found more quickly.
### Original VBA Script
![Original Module 2 2017 Analysis](https://github.com/zbarham/stock-analysis/blob/main/Resources/Module_2_2017.png) & ![Original Module 2 2018 Analysis](https://github.com/zbarham/stock-analysis/blob/main/Resources/Module_2_2018.png)
### Refactored VBA script
![Refactored VBA Challenge 2 2017](https://github.com/zbarham/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png) & ![Refactored VBA Challenge 2 2018](https://github.com/zbarham/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
### Code that helped reach these Achievements
Refactoring the code to make use of a ticker index to count through the full stock sheet at once rather than look for each stock individually was a significant improvement in execution time and efficiency. Going through the full sheet at once and counting as we go and only increasing the tickerIndex when the code does not see the current ticker means that we have a more efficient process.
### Refactored loop
	For i = 2 To RowCount
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        'End If
        End If
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d) Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        'End If
        End If
    Next i

## Summary
### Advantages or Disadvantages of Refactoring code
The advantages of refactoring code are improved efficiency, better readability, and the ability to handle larger datasets. The disadvantage to refactoring code is errors can be introduced in your efforts to improve the already working code, causing the need for more time or resources to fix the code. In general refactoring code is a great positive in efficiency and the larger the dataset the more impact the time reduction will have in a positive way on server resources and time.

### Pros and Cons of Refactoring the Original VBA script
The pros to refactoring the original vba script is it can now handle a larger dataset and produce the same output as the original script but in a significantly faster time. The improved efficiency should be easier for others to read and understand. Easier time reading the code means easier for others to update or fix your code if needed.

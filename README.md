# Stock Analysis

## Overview of Project
The overall essence of the project was to create a VBA module that could filter through a given dataset and analyze its contents.

### Purpose
The purpose of the challenge was to use VBA to create a general code that would filter through various stocks in a given year. The code was then expected to give an output of each stock's daily volume and percent increase or decrease in said year. The final output was required to be formatted appropriately for ease in reading.

## Results
### Stock Performance Between 2017 and 2018
Upon initial observation, it can be seen that stocks in 2017 had a better performance, as eleven out of twelve of the stocks being analyzed had a positive return (see image below). In addition, four of the twelve stocks gave a return of greater than 100%. 

<img width="253" alt="Screen Shot 2022-03-19 at 10 37 45 AM" src="https://user-images.githubusercontent.com/86126331/159125427-fcda0fed-6f30-4ca1-b027-a9e9909e1f1f.png">

When looking at the data for 2018, it can be seen that only two out of the twelve stocks being analyzed had a positive return, with the remaining ten having a negative return. In addition, the stocks with a positive return did not excede 100%, unlike the stocks in 2017 (image of 2018 stocks below).

<img width="250" alt="Screen Shot 2022-03-19 at 10 49 27 AM" src="https://user-images.githubusercontent.com/86126331/159125876-fa2af725-d924-47c0-b671-1f4726bf2c69.png">

By analyzing the performance of the same twelve stocks in both 2017 and 2018, it can be said that the specified stocks had a better performance in 2017.


### Execution Times of Original vs. Refractored Script
When executing the code for both 2017 and 2018, a slight difference in the execution time can be observed. For 2017, the execution time was 0.105 seconds while the execution time for 2018 was 0.113 seconds (images below). 

<img width="251" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/86126331/159126313-a72e2d69-cd08-47a7-90d8-974734b8f13e.png">
<img width="257" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/86126331/159126330-367bb01a-308b-444d-bc01-4927a12dd90f.png">

The lack of a significant difference between the execution times can be attributed to the fact that both data sets involve the same twelve stocks, and the same analyses. The minute difference may be attributed to the time it took the program to format the results, in terms of whether the respective 'Return' cells should be green or red.

        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If

When looking at the above formatting code, it can be seen that all 'Return' cells with a positive value will require the program to run through fewer lines of code. Therefore, the sample with the fewest negative values in the 'Return' column will have the shortest execution time. 

## Summary
An advantage of refactoring code is the convenience of using the same code for multiple similar datasets. Ultimately, this would mean less work for the developer as each they would not have to start from scratch each time a new dataset, similar to a previous one, arrives. A disadvantage of refactoring code is the multiple variables that require the raw data to be presented in a certain way. 

    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i

For instance, the above lines of code, taken from my 'AllStocksAnalysisRefactored', will run smoothly if the raw data has values starting from the second row and first column. If the raw data does not start from there, the refactored code will not run and result in an error.

# Refactor VBA Stock Analysis 
Click here to view the spreadsheet: [VBA_Challenge.xlsm](Kickstarter_Challenge.xlsm)

## Overview of Project
Steve’s parents are interested in green energy but have not done a lot of research and decided to invest solely in DAQO New Energy Corp. As a new financially analyst, Steve wants to analyze green energy stock in addition to DAQO to help his parents diversify their investment. 

Steve has now expanded the dataset from a dozen stocks to now include the entire stock market over the last few years. The larger dataset may cause a long time to execute using my original code. The purpose of the project is to edit or refactor my original code to make the coding process more efficient. My goal is to edit my original code to now take fewer steps, use less memory, and improve the logic of the code to make it easier for future users to read. The new refactor code should have a faster execution time for the larger dataset. 

## Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### Analysis 
To begin my analysis, I used my original code as a starting point to then edit and refractor. I took the following steps to improve and refractor my original code. 
1.	1a) I created a “tickerIndex” variable and set it to zero before iterating over all the rows. I used this variable to access the correct index across the arrays I will be using in my code. 
1.	1b) Next, I created three output arrays for the following variables “tickerVolumes” as a Long data type and “tickerStartingPrices” and “tickerEndingPrices” as a Single data type. 
2.	2a) Then, I created a for loop to initialize the “tickerVolumes to zero. 
2.	2b) I also created another for loop to loop over all the rows in the spreadsheet
3.	3a) Inside the loop in Step 2b, I wrote a script that increases the current “tickerVolumes”, which represents the stock ticker volume. The script adds the ticker volume for the current stock sticker. 
3.	3b) With an if-then statement, I check if the current row is the first row with the selected “tickerIndex”. If yes, then I assign the current starting price to the “tickerStartingPrices” Variable. 
3.	3c) Again, with an if-then statement, I check if the current row is the last row with the selected “tickerIndex”. If yes, then I assign the current starting price to the “tickerEndingPrices” Variable. 
3.	3d) Next, I write a script to increase the tickerIndex if the next row’s ticker does not match the previous row’s ticker. 
4.	4) Finally, I use a for loop to loop through my arrays of “ticker”, “tickerVolume”, “tickerStartingPrices”,  and “tickerEndingPrices” to output the columns “Ticker”, “Total Daily Volume”, and “Return” in my spreadsheet. 

### VBA Script
Please see my script that represent my analysis steps. 

```
'1a) Create a ticker Index
    
    tickerIndex = 0
    

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For j = 0 To 11
        
        tickerVolumes(j) = 0
    
    Next j
        
    ''2b) Loop over all the rows in the spreadsheet.

    For i = 2 To RowCount
    
    
        '3a) Increase volume for current ticker
        
        'increase tickerVolumes
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
                
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
           tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d) Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        
        End If


        'End If
        
    Next i
    

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
    Next i
```

### Execution Times Comparison 
The benefit of refactoring is more efficient code that has faster run time that overall increases productivity. Please see the following execution time and results for the refactor code as well as the execution time for the original code. For both 2017 and 2018, the refactor code has a faster execution time. In 2017, the difference in execution time is 0.945313 = 1.132813 – 0.1875 If we extrapolate the difference in run time for 12 stocks to say 120,000 stocks the difference in execution time is now around 16 minutes. The difference in execution time demonstrates the importance of refactoring when coding for larger datasets and more complicated analysis. 

#### 2017 
![VBA_Challenge_2017] (Resources/ VBA_Challenge_2017.png)
![ VBA_Challenge_2017_Output] (Resources/ VBA_Challenge_2017_Output.png)

#### 2018 
![VBA_Challenge_2018] (Resources/ VBA_Challenge_2018.png)
![ VBA_Challenge_2018_Output] (Resources/ VBA_Challenge_2018_Output.png)

#### Original Execution Time 
![ VBA_Challenge_Original_2017] (Resources/ VBA_Challenge_Original_2017.png)
![ VBA_Challenge_Original_2018] (Resources/ VBA_Challenge_Original_2018.png)

## Summary 

### Advantages and Disadvantages of Refactoring Code
The largest advantages of refactoring code are cleaner code and faster execution time. The cleaner your code is, the easier it is for everyone to understand, debug, and improve your code in the future. The faster your code runs, the faster you can produce in depth analysis. 

The main disadvantage of refactoring code is that it typically comes after a first attempt at the code and you much have the luxury of time. When there are deadlines and other projects, you may not have the time to edit and improve the code to be more efficient. Another risk of refactoring is breaking the analysis in the process and taking more time to debug code that was working before. 

### Original and Refactored VBA Script Advantages and Disadvantages 
The main advantage of the refactored VBA script is the faster execution time. Please see the subsection “Execution Times Comparison” for the decreased execution time of the refactored script and the benefits. 

An advantage of the original code, which is also a disadvantage for the refactored code, is it might be slightly easier a more intuitive to understand the for loop that goes through 0 to 11 for the ticker in the original code than the ticker index that we create in the refactor code. 

Overall, the advantages of efficiency and reduced execution time of the refactored scripts is main reason to update and improve our code.  


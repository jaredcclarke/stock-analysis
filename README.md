# Stock Analysis with VBA
## Overview of Project
### Purpose
The purpose of this excercise was to refactor a solution code for Steve that analyzes the total daily voulume and the stock's percent return of several stocks. The code needed to be refactored in order to run faster, easy to use, and be able to handle a large data set. To do this, the code would need to be more effecient by using fewer steps and use less memmory. Here we are testing the code with two different years of stck data with twelve different stock tickers that were provided in the [green_stocks.xlsm] (https://github.com/jaredcclarke/stock-analysis/blob/master/green_stocks.xlsm) file.
## Results
A `tickerIndex` was created and set it equal to zero before iterating all the rows and it woull allow Steve to access the correct index across the four different arrays that were created: `tickerIndex`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`. A `For` was also created. 
```    For I = 0 To 11
       tickerIndex = tickers(I)
 ```
The `tickersVolumes` was listed as a `Long` data type while `tickerStartingPrices`, and `tickerEndingPrices` were listed as `single` data types
    ``` 
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
    ```
After the `For` loop was created, the `tickersVolume` was initialized to zero
```tickerVolumes = 0```
Then, another `For` loop was created. this loop will loop over all the rows in the spreadsheet.  
```For j = 2 To RowCount```
For this inner loop  we increase the current tickerVolume variable and adds the ticker volume for the current stock ticker by using the tickerIndex variable as the index with the following script: 
          ``` 
            If Cells(j, 1).Value = tickerIndex Then
           
           tickerVolumes = tickerVolumes + Cells(j, 8).Value 
          ```
To calculate the yearly percent return of the stocks, we need to know the starting price of the stocks at the beginning of the year and also the price at the end of the year. That means we need to look at the first row of a specific stock for the starting price and the last row for its ending price. `IF-Then` statemetnts were used script a way to find the starting and ending prices. They are as follows: 
  
  ```
   If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerStartingPrices = Cells(j, 6).Value
    
    If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerEndingPrices = Cells(j, 6).Value
  ```
   Next, we used another `For` loop to loop over the arrays to the output sheet `All Stocks Analysis`. 
    
        ```
           Cells(4 + I, 1).Value = tickerIndex
           Cells(4 + I, 2).Value = tickerVolumes
           Cells(4 + I, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
        ```
 ### Images
 After running the code, we get the following results:
  ![VBA_Challenge_2017.png] (https://github.com/jaredcclarke/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)
  ![VBA_Challenge_2017.png] (https://github.com/jaredcclarke/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)
## Summary
### Advantages of Refactoring Code
The advantages of refactoring code are readability, with cleaner code that has no repeats, making it easier to understand. It also helps finding errors and will help the code run faster. 
### Distadvantages of Refactoring Code
Te disadvantages with refactoring code with respect to this specific data was mainly that it was time consuming. Going line by line to pick out and refactor the code is tedious work. It also lead to me having to debug the script several times because of errors. 
  
    

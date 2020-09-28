# Stock-Analysis with VBA 

## Overview of Project
The Purpose of this project was to refactor a useful VBA function so that it can be effective on a much larger scale. Although the original function serves its purpose and runs relativley quickly; It only has to iterate over a dataset of 12 stocks and contains a nested `for` loop which is considered malpractice. We will want to run it as quickly as possible for thousands of stocks therefore we need to ensure its ability to do this. Luckily in VBA there is a `Timer` function we can ustilize to measure the speed of our scripts. Our Goal will be to take the VBA function we had previously created and modify/refractor it to run at a faster rate which is testable by the previously mentioned `Timer`. We will be using most of the original code as is but will find a way to improve the blocks relying on the nested `for` loop.  
## Results and Analysis

### Data Analysis
The output of our function displays the data aranged into 3 columns. The Ticker Column represents the name of the stock, Total Daily Volume Column represents the total daily volume for the corresponding stock, and the final Return Column displays the return percentage or the stocks ending price divided by its starting price.
Finnally the Return percentage is code red for a negative and green for a positive return percentage.

#### 2017
![alt text](https://github.com/sebcampos/stock-analysis/blob/master/other_pngs/2017.png)

In the above image we can all of the stocks except for TERP giving a positive return percentage and their coresponding Total Daily Volume. It appears that some ove the stock is returning over 100 percent ! now lets compare these numbers from 2017 to the graph image below depicting data from 2018

#### 2018
![alt text](https://github.com/sebcampos/stock-analysis/blob/master/other_pngs/2018.png)

When we compare our 2017 image with the above 2018 image one difference is immediatley clear thanks to our formating of the data. There are many more red cells than before. Our script utilized a condition `if` statement to format the negative returns as red and positive as green. We can see that almost all of our return percentage has dropped, and although some did increase most of the data set had slipped into the negative return range.


### VBA Analysis
#### Original Function
![alt text](https://github.com/sebcampos/stock-analysis/blob/master/other_pngs/VBA_function.png)

The above png is a image of the portion of our Original VBA script specifically the portion that iterates over the cells. The iteration portion relys on a nested `for` loop. 
We begin with 

`For i = 0 To 11`

This creates 12 loops or iterations represented by the variable `i`, this loop only loops through the integers  in range 0 To 11; Within this loop we create another loop defined by the variable `j` which iterates from the lowest row applicable `2` to the lowest row defined as `lastRow`, again this only loops through the range of integer values in the range of 2 To the number represented by the variable `lastRow`

    For i = 0 To 11
    
        For j = 2 To lastRow
the above code will begin iteration at 0 then within that iteration it will loop over every row, the at iteration index 1 it will again loop over every row, it will continue to do this untill i reaches the last value of 11. This loop combination iterates over all the rows in the excel 12 times and collects the aporoiate values with conditionals. For only 12 items to iter over this process works in under 1 second. Below is the resulting pop up displayed by the `timer` function giving us the time it took to run

#### Refactored Function
![alt text](https://github.com/sebcampos/stock-analysis/blob/master/other_pngs/VBA_Functionrefactored.png)

## Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
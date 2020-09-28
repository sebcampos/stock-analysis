# Stock-Analysis with VBA
 
## Overview of Project
The Purpose of this project was to refactor a useful VBA function so that it can be effective on a much larger scale. Although the original function serves its purpose and runs relatively quickly; It only has to iterate over a dataset of 12 stocks and contains a nested `for` loop which is considered malpractice. We will want to run it as quickly as possible for thousands of stocks therefore we need to ensure its ability to do this. Luckily in VBA there is a `Timer` function we can utilize to measure the speed of our scripts. Our Goal will be to take the VBA function we had previously created and modify/refactor it to run at a faster rate which is testable by the previously mentioned `Timer`. We will be using most of the original code as is but will find a way to improve the blocks relying on the nested `for` loop. 
## Results and Analysis
 
### Data Analysis
The output of our function displays the data arranged into 3 columns. The Ticker Column represents the name of the stock, Total Daily Volume Column represents the total daily volume for the corresponding stock, and the final Return Column displays the return percentage or the stocks ending price divided by its starting price.
Finally the Return percentage is code red for a negative and green for a positive return percentage.
 
#### 2017
![alt text](https://github.com/sebcampos/stock-analysis/blob/master/other_pngs/2017.png?raw=True)
 
In the above image we can see all of the stocks except for TERP giving a positive return percentage and their corresponding Total Daily Volume. It appears that some over the stock is returning over 100 percent ! now let's compare these numbers from 2017 to the graph image below depicting data from 2018
 
#### 2018
![alt text](https://github.com/sebcampos/stock-analysis/blob/master/other_pngs/2018.png?raw=True)
 
When we compare our 2017 image with the above 2018 image one difference is immediately clear thanks to our formating of the data. There are many more red cells than before. Our script utilized a condition `if` statement to format the negative returns as red and positive as green. We can see that almost all of our return percentage has dropped, and although some did increase most of the data set had slipped into the negative return range.
 
 
### VBA Analysis
#### Original Function
![alt text](https://github.com/sebcampos/stock-analysis/blob/master/other_pngs/VBA_function.png?raw=True)
 
The above png is an image of the portion of our Original VBA script specifically the portion that iterates over the cells. The iteration portion relies on a nested `for` loop.
We begin with
 
`For i = 0 To 11`
 
This creates 12 loops or iterations represented by the variable `i`, this loop only loops through the integers  in range 0 To 11; Within this loop we create another loop defined by the variable `j` which iterates from the lowest row applicable `2` to the lowest row defined as `lastRow`, again this only loops through the range of integer values in the range of 2 To the number represented by the variable `lastRow`
 
   For i = 0 To 11
  
       For j = 2 To lastRow
The above code will begin iteration at 0 then within that iteration it will loop over every row, at iteration index 1 it will again loop over every row, it will continue to do this until variable `i` reaches the last value of 11. This loop combination iterates over all the rows in the excel worksheet 12 times and collects the appropriate values with conditionals, `if` statements. For only 12 items to iterate over this process works in under 1 second. Below is the resulting pop up displayed by the `timer` function giving us the time it took to run for the year 2017.
 
![alt text](https://github.com/sebcampos/stock-analysis/blob/master/other_pngs/Original_function_timer_2017.png?raw=True)
 
 
 
#### Refactored Function
![alt text](https://github.com/sebcampos/stock-analysis/blob/master/other_pngs/VBA_Functionrefactored.png?raw=True)
 
Above is the refactored function, this new refactored version is almost identical, but instead of using a nested `for` loop it uses 2 different `for` loops. This is much more efficient because one `for` loop will loop to set the values for one array, then the other loop will iterate over every row in the sheet once grabbing and storing all relevant data in only one iteration as opposed to the previous 12. A screenshot of our timer function for the refactored function shown below reveals that the runtime for this refactored function was about 4 times faster!
 
![alt text](https://github.com/sebcampos/stock-analysis/blob/master/resources/VBA_Challenge_2017.png?raw=True)
 
 
 
## Summary
 
We create functions to serve a purpose. Sometimes the function is needed immediately. Sometimes we take out time and build it as efficiently as possible. But as we accumulate more knowledge and find new ways to accomplish different tasks or goals one could greatly benefit from editing or refactoring an already existing function especially if it will be run or applied in a different environment. I will be taking a look at all my python functions after this final commit! As for our VBA functions we were able to significantly cut down the run time and although either time runs under one second which could seem insignificant to us, we could imagine that if applied to a much larger set of data one could find it extremely useful.

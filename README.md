# Stock-Analysis with VBA 

## Overview of Project
The Purpose of this project was to refactor a useful VBA function so that it can be effective on a much larger scale. Although the original function serves its purpose and runs relativley quickly; It only has to iterate over a dataset of 12 stocks and contains a nested `for` loop which is considered malpractice. We will want to run it as quickly as possible for thousands of stocks therefore we need to ensure its ability to do this. Luckily in VBA there is a `Timer` function we can ustilize to measure the speed of our scripts. Our Goal will be to take the VBA function we had previously created and modify/refractor it to run at a faster rate which is testable by the previously mentioned `Timer`. We will be using most of the original code as is but will find a way to improve the blocks relying on the nested `for` loop.  
## Results and Analysis
The analysis is well described with screenshots and code (4 pt).
Using images and examples of your code
### Data Analysis
The output of our function displays the data aranged into 3 columns. The Ticker Column represents the name of the stock, Total Daily Volume Column represents the total daily volume for the corresponding stock, and the final Return Column displays the return percentage or the stocks ending price divided by its starting price.

![alt text](https://github.com/sebcampos/stock-analysis/blob/master/other_pngs/2017.png)



### VBA Analysis
___execution times of the original script and the refactored script.
## Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
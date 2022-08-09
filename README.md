# Analyzing Stocks with VBA

## Overview of Project

### Purpose
By using Excel's Visual Basic for Applications, or VBA, we are able to write scripts to automate simple tasks that allow Steve analyze stocks, which will help him direct his parents on how to invest their money. By the click of a button, Steve will be able to choose what year he would like to analyze, then the tickers, total daily volumes, and yearly returns for each stock will be generated in a neatly organized table. With the ease of this process, Steve can quickly and clearly let his parents know that DQ may not be a profitable stock to invest in and guide them in the right direction.

## Analysis and Results

### Analysis of Total Daily Volume

Using my knowledge of VBA and the starter code provided in this Challenge, I was able to refactor the scipt so that I looped through the data one time and collected all of the information. 

A helpful tool that was used to streamline Steve's analysis was including an InputBox. This allowed Steve to enter the year that he wanted to analyze. In the future, he can insert new sheets each year and use the same workbook to make his analysis.

To help with this desciption, I will use the concrete example of the ticker "AY". 

Steve requested that I collected the Total Daily Volume and Returns for each stock. To do this, I created nested for loops. The outer loop focused on each ticker one at a time, using a tickerIndex variable. Since "AY" was our first ticker, it would have a tickerIndex of 0. This meant that the inner loop would only focus on finding the ticker "AY" and collect and store the information associated with it. Once all information is stored for "AY" then the outer loop will move onto the next ticker, "CSIQ", and so forth. The inner loop cycles through all of the rows, in column 1, in the selected sheet to search for the current tickerIndex, "AY". Once the code indentified that the cell contains "AY" we used the code `tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value` to increase the Total Daily Volume for "AY". Notice the Total Daily Volume is being stored in an array called tickerVolumes.

Using an array was very beneficial in this process because the tickerVolumes array will be able to hold the Total Daily Volume for each ticker, based on the tickerIndex. I like to think of an array as a table like the one shown below. Since we are focusing on "AY", which has a tickerIndex of 0,  we are storing the Total Daily Volume for "AY" in the first, left-most, cell in the table (array). Once that information is collected, our inner loop will continue to cycle through the rows until it finds the next "AY". Then it will use the same code as above to update and store the new Total Daily Volume in the array, in it's correct tickerIndex position. This will continue until the inner loop has cycled through all rows containing "AY", updating the Total Daily Volume each time, at which point the outer loop will move to the next tickerIndex and the same process will begin for the next ticker "CSIQ". Arrays for tickerStartingPrices and tickerEndingPrices will be used in the same way. 

tickerVolumes array

|tickerIndex       |  0  |   1  |  2  | ... |
|------------------|-----|------|-----|-----|
|Total Daily Volume|  AY | CSIQ |  DQ | ... |


### Analysis of Returns

At the same time that the inner loop is cycling through each row to collect the Total Daily Volume for "AY", the code is also determining the starting price and ending price. To determine the starting price for each ticker, an if statement is used. If the cell contains "AY" and the previous cell above it does not, then the closing price (in column 6) associated with the it is stored in the tickerStartingPrices array in the tickerIndex that corresponds to "AY". 'i' corresponds to the current row number. 
```
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value 
    
End If
```
To determine the ending price for each ticker, a similar if statement is used but now if the cell contains "AY" and the next cell below it does not, then the closing price (in column 6) associated with it is stored in the tickerEndingPrices array in the tickerIndex that corresponds to "AY"

***Note: The way that the code is written for starting price and ending price assumes that the stocks are in cronological order.*** 

### Results

After sorting each years returns from largest to smallest, we see that Steve's parents stock DQ did excellent in 2017, but poorly in 2018. Based on consistent returns, I would suggest Steve look further into ENPH and RUN, as those were the two stocks that had positive returns in both years. 

![2017_All_Stocks_Analysis_Returns](https://user-images.githubusercontent.com/109091887/183675282-135b713f-dac4-4363-9051-4dae364ee42e.PNG)          ![2018_All_Stocks_Analysis_Returns](https://user-images.githubusercontent.com/109091887/183675308-540de1a9-6af6-4798-9875-905b4521d002.PNG)

## Summary

Refactoring code is essential to truly understanding the script-writing process. The key to refactoring is efficiency. Instead of attempting to add new loops or statements, we want to take what is already given and simplify it. By refactoring code, we want to take less steps by eliminating redundancies, shorten the process with respect to run time, and solidify that the structure is logically sound. As we have discussed in class, the first time you write code it may work but as we gain experience we should revisit the code to determine if we can improve it through refactoring. Although refactoring code is beneficial to understanding the process, it has its disadvantages. It's possible that you take code that is working just fine, and by changing it you cause the script to no longer run. To avoid this, it is important to be very careful with your adjustments. Refactoring can be tedious and time consuming work, which could be considered a disadvantage. 

Specific to this Challenge, refactoring the script by using arrays to store Total Daily Volume, Opening Price, and Closing Price based on the tickerIndex allowed us to loop through the data one time and have all the information stored, then presented all together at the end. This has the advantage of making the code more efficient and display more data. The original script was well designed so Steve could analyze the stock his parents were interested but it limited him to only viewing once stock at a time. By refactoring the code, he has the ability to view many stocks at once, making it easier to compare them. The original script had its limitations but by only focusing on one stock at a time, the script ran faster. That being said, I believe that even though our refactored code will take longer to run, it allows for more analysis to be done on the stocks. 

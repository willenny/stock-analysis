# Analyzing Stocks with VBA

## Overview of Project

### Purpose
By using Excel's Visual Basic for Applications, or VBA, we are able to write scripts to automate simple tasks that help Steve analyze stocks to direct his parents on how to invest their money. By the click of a button, Steve will be able to choose what year he would like to analyze, then the tickers, total daily volumes, and yearly returns for each stock will be generated in a neatly organized table. With the ease of this process, Steve can quickly and clearly let his parents know that DQ may not be a profitable stock to invest in and guide them in the right direction.

## Analysis and Results

### Analysis of Total Daily Volume and Returns

Using my knowledge of VBA and the starter code provided in this Challenge, I was able to refactor the scipt so that I looped through the data one time and collected all of the information. 

A helpful tool that was used to streamline Steve's analysis was including an InputBox. This allowed Steve to enter the year that he wanted to analyze. 

To help with this desciption, I will use the concrete example of the ticker "AY". 

The first piece of information that we collected was the Total Daily Volume. To do this, I created nested for loops. The outer loop focused on each ticker one at a time, using a tickerIndex variable. Since "AY" was our first ticker, it would have a tickerIndex of 0. This meant that the inner loop would only focus on finding the ticker "AY", collect all of the information associated with "AY", then move onto the next ticker, "CSIQ", and so forth. The inner loop cycled through all of the rows, in column 1, in the sheet to search for the current tickerIndex, "AY". Once the code indentified that the cell contained "AY" we used 'tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value' to increase 






## Summary

Using arrays to store Total Daily Volume, Opening Price, and Closing Price based on the tickerIndex allowed us to loop through the data one time and have all the information stored, then presented all together at the end. 

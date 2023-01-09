# VBA-challenge

This code group is for an assignment in excel that generally takes multiple sheets with alphabetized, repetitive discrete data and collects characteristics for each of the discrete values. 

In the specific example it was developed for, there are lists of tickers with opening, closing, and volumes of stock sold on multiple days. 
The code is written to collect ticker names by checking for changes in values in column 1 on each sheet. 
Once a change has been detected, I used an if then statement to collect the following:
  The ticker name
  The opening value of the next ticker
  The closing value of the previous ticker
  Document the yearly change (closing - opening) for the last ticker
  Document the total stock volume for the last ticker
  
Using a for loop, until the ticker has changed, the total will add each day's volume, to be documented in a summary when the ticker changes. 

After all of the sheets have had the discrete ticker names listed and characteristics found, the extrema can be searched. 
This is initiated with a button press, and the code checks down the columns of percent change and total volume to find the max or min and record the ticker and value in a separate summary area. 
To accomplish this, the code records a new max or min value as it finds one larger or smaller, respectively, than it's previous value. 
Once it has checked all rows, it records the final max or min.

This code can be applied to other similar data by just changing the column names to whatever you want to find, and the column numbers in the code to where the data is stored. It is already written to search all worksheets in a workbook, no matter how many there are. 

# Stock Performance Analysis

## Overview of Project
The purpose of this project is to be able to look at a set of identified stocks and determine their total volume and overall performance over the course of an identified year.  The initial project was designed to focus on 12 specific stocks whereas this updated code has been designed to be more flexible to allow for a much larger set of data and to operate in a quicker/more time efficient manner.  It includes modified formatting for the output to make the resulting data more easily referenced as well as added code to display the amount of time it took to run the given request so that as the datasets become larger, potential performance degradation can be monitored.  The spreadsheet was designed with 2 buttons:  one used to run the VBA script to generate the analysis, requesting the specific year from the end user, and the other to clear out the data and formatting so that a separate run of the data could be performed.

##  Results
The analysis of this data produced 4 key results:  the total daily volume for the identified year, the stocks' starting price for the year, the stocks' ending price for the year, and then a calculated return for the year.  For comparison purposes, the VBS in this project includes both the original coding as well as the refactored coding.

[VBA_Challenge](https://github.com/purvishjd/stock-analysis/blob/main/VBA_Challenge.xlsm)

### 2017 Results
Using the code "yearValue = Inputbox("What year would you like to run the analysis on?)", the user is able to specify the desired year for the analysis.  Once this input is received, the code is then executed to obtain the total volume, starting price, and ending price for the array of pre-defined stocks.  The overall return for the year for each stock is then calculated and added to the output for the worksheet.  The resulting data is then formulated to show which stocks showed a positive return and which showed a negative return, differentiated by the color highlight for the result:

[2017_Results](https://github.com/purvisjd/stock-analysis/blob/main/2017_Results.png)

Based on this information, only one stock did poorly in 2017 (TERP) while three others (DQ, ENPH, SEDG) did exceptionally well for that year.

### 2018 Results
The same input code was used to obtain the results for the same set of stocks for 2018 as well.  These results showed a drastic change in outcome over the year as only ENPH and RUN generated positive returns:

[2018_Results](https://github/purvisjd/stock-analysis/blob/main/2018_Results.png)

Compairing the two years against each other, it is observed that only ENPH and RUN had positive returns in both years, with ENPH showing significant returns for both years identifying as the most promising investment option.

##  Summary

###  Speed Test Comparison
When comparing the refactored code to the original analysis code, multiple observations can be made:  First, the formatting that was updated for the new code makes the resulting data much easier to review at a glance and easier to understand.  Second, the speed test on running the code is significantly improved with the refactored code.  The code as become more efficient.  Rather than having to loop through nested For loops, the code goes through one loop creating 3 separate arrays to hold the resulting data, eliminating the need to constantly reset variables.  For comparison, here are the times based on the original code as well as the refactored code:

[Old_code_speed_test_2017](https://github.com/purvisjd/stock-analysis/blob/main/old_code_speed_test_2017.png)
[VBA_Challenge_2017](https://github.com/purvisjd/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

[Old_code_speed_test_2018](https://github.com/purvisjd/stock-analysis/blob/main/Old_code_speed_test_2018.png)
[VBA_Challenge_2018](https://github.com/purvisjd/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

With the addition of variables as opposed to hard coding in line counts, the code is more adaptable to larger datasets without having to make significant changes.  While hard coding is easier because you don't need to worry about the potential for information being stored incorrectly as much, it is more cumbersome and harder to follow.  Refactoring the code allowed me to generalize it and make it more versatile; able to interact with larger datasets with minimal further input and modification.  



# Stock Analysis

## Overview of Project

### Purpose
The purpose of this analysis is to look at a subset of stocks and figure out which stocks are most likely to be profitable.

## Results

### Comparison of Stock Performance
All stocks except one had a positive return in 2017.

![2017 Stock Performance](/Resources/Stock_Performance_2017.png)


2018 was the opposite, with only two stocks having positive returns.  ENPH and RUN had returns of over 80% with high volume, so they look like the best stocks to purchase right now.

![2018 Stock Performance](/Resources/Stock_Performance_2018.png)

### Comparison of Script Performance
The original VBA script took 1.25 seconds to run for 2017, and 1.15 seconds to run for 2018.

![Original Script Performance for 2017](/Resources/VBA_Challenge_2017_Original_Code.png)
![Original Script Performance for 2018](/Resources/VBA_Challenge_2018_Original_Code.png)


The refactored script was much faster, running in 0.11 seconds for 2017, and 0.09 seconds for 2018.

![Refactored Script Performance for 2017](/Resources/VBA_Challenge_2017.png)
![Refactored Script Performance for 2018](/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and Disadvantages of Refactoring Code
Refactoring code can improve it by making it run faster.  Also, by condensing the code into reusable parts, it can be edited more easily in the future since it doesn't have to be changed in multiple places.  The disadvantages of refactoring code are the possibility of adding bugs to it, and potentially making it more complex and harder to read.

### Advantages and Disadvantages of Refactoring This VBA Script
Refactoring this script definitely improved its speed.  This is because the old script looped over all the rows 12 times (once for each ticker), while the new script only loops over the rows once.  It also adds simplicity by gathering all the data at once and then printing all the results out at once, instead of repeatedly switching between worksheets.  The disadvantage of refactoring the script is that the array code is more confusing to read than the code that just makes a calculation and puts it on the analysis worksheet right away.
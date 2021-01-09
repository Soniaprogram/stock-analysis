# stock-analysis

## Overview of Project
Steve wants to analyze different green energy stocks similar to Daqo (DQ) to help diversify his parent’s stock portfolio. I used VBA code to automate the analysis to allow Steve to reuse the code to look into different sets of stocks in the future.
I created a worksheet to analyze a list of stocks and report the total daily volume (indicating how actively a stock is traded) and yearly return (percentage difference in price from the beginning to the end of the year) for each stock. Using conditional formatting, I highlighted positive returns as green and negative returns as red to make it easier to see which stocks performed well. 
Steve also wants to see how fast the VBA code will compile the results so I programmed a script to calculate how long the code will take to execute and output the elapsed time in a message box. 

### Purpose
Steve wants to expand the dataset to include the entire stock market over the last few years. Thus, the purpose of this project is to refactor the code to allow for a large dataset to be analyzed in an efficient amount of time. The refactored code will loop through all the data only one time to output the total volume and yearly return.  

## Analysis

### Deliverable 1: Refactor VBA Code and Measure Performance

#### Compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script
The execution time of the original script for the 2017 dataset was 0.854s and for the 2018 dataset was 0.839s as seen below. 
![image1: 2017 original code runtime](https://github.com/Soniaprogram/stock-analysis/blob/main/Additional%20Screenshots/Runtimes%20of%20original%20code/Original2017.PNG)
![image2: 2018 original code runtime](https://github.com/Soniaprogram/stock-analysis/blob/main/Additional%20Screenshots/Runtimes%20of%20original%20code/Original2018.PNG)

The refactored script execution time for the 2017 dataset was 0.227s and for the 2018 dataset was 0.226s as seen below.
![image3: 2017 refactored code runtime](https://github.com/Soniaprogram/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
![image4: 2018 refactored code runtime](https://github.com/Soniaprogram/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

In both instances, the 2017 dataset took a little longer to execute. Analyzing the same dataset, the refactored script execution time is less than the original script execution time, averaging a difference of 0.6s. I believe this difference would be more apparent with an even larger dataset.
In terms of stock performance, there were more positive returns in 2017 compared to 2018. The total daily volumes varied a bit. 5 stocks had a decreased total daily volume in 2018, whereas the remaining 7 had an increased total daily volume in 2018. However, it is difficult to see a direct correlation between the total daily volume change and positive/negative return of each stock. 

![image5: 2017 chart](https://github.com/Soniaprogram/stock-analysis/blob/main/Additional%20Screenshots/Results/2017chart.png)
![image6: 2018 chart](https://github.com/Soniaprogram/stock-analysis/blob/main/Additional%20Screenshots/Results/2018chart.PNG)

## Summary

#### What are the advantages or disadvantages of refactoring code?

The advantages of refactoring code are a faster runtime, requiring less steps, less memory, and easier code readability for future users since it only loops through all the data one time. 
The disadvantages of refactoring code would be the time spent having to go back to the original code to make these changes. 

#### How do these pros and cons apply to refactoring the original VBA script?

The refactored code allowed for a faster runtime as seen by the calculated execution times. The execution time of the original script for the 2017 dataset was 0.854s and for the 2018 dataset was 0.839s. However, the refactored script execution time for the 2017 dataset was 0.227s and for the 2018 dataset was 0.226s.

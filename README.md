# Stock-analysis
Analyzing stock data with VBA 

## Overview of Project
The purpose of the project is to analyze and observe the benefits of refactoring code. The analysis was performed on two datasets which include stock data from the years 2017 and 2018. Each dataset contained 3012 items and recorded select daily activity of 12 stocks. The stock data collected was divided into eight columns, including ticker, date, high, low, open, close, adjusted close, and volume. Code ran that analyzed and extracted specfic data from our dataset. The challenge is to refactor code in effort to make the program run more efficiently. Code performance was objectively measured by a timer, which measured the total time it took to run the macro and store data in our workbook.

## Results
The analysis extracted and categorized data into three groups. The first group displays the ticker, or stock. The second group contains total daily volume for each stock, while the third category displays percentage return from the years data for each stock. The percentage return was found by establishing starting price at the beginning of the year for each stock and ending price at the end of the year for the same stock. This was performed on all 12 stocks and enables the viewer to identify which stocks improved their value over the course of the year and which decreased their value. The analysis found generally superior stock performance during the 2017 year with only one stock, TERP, finishing with a negative return. The 2018 performances for the same stock generally dropped and saw only two stocks, ENPH and RUN, with positive returns. The code used to perform these tasks also utilized a timer and captured the run time for each script. Original VBA stock analysis scripts for 2017 and 2018 data ran in roughly 1.5 seconds while the refactored scripts ran significantly faster, coming in around 0.2 seconds.

**2017 Stock Performance**<br> 
<img src= "https://github.com/ChrisBarton107/Stock-analysis/blob/master/Resources/Resources/Stock_Performance_2017.png" width= "1000"><br>
**2017 Original Timer**<br> 
<img src= "https://github.com/ChrisBarton107/Stock-analysis/blob/master/Resources/Timer_Original_2017.png" height="300" width="700"><br>
**2017 Refactored Timer**<br> 
<img src= "https://github.com/ChrisBarton107/Stock-analysis/blob/master/Resources/VBA_Challenge_2017.png" height="300" width="700"><br>
**2018 Stock Performance**<br> 
<img src= "https://github.com/ChrisBarton107/Stock-analysis/blob/master/Resources/Resources/Stock_Performance_2018.png" width="1000"><br>
**2018 Original Timer**<br> 
<img src= "https://github.com/ChrisBarton107/Stock-analysis/blob/master/Resources/Timer_Original_2018.png" height="300" width="700"><br>
**2018 Refactored Timer**<br> 
<img src= "https://github.com/ChrisBarton107/Stock-analysis/blob/master/Resources/VBA_Challenge_2018.png" height="300" width="700"><br>

## Summary
**Refactoring: Advantages & Disadvantages**
Code performance is often a critical component in machine based learning systems and improving aspects of code can make significant differences in the functionality of programs and their applications in the world. Refactoring code has many advantages and enables the coder to optimize and improve the code in many ways. Bugs are often detected and bad patterns, including tight coupling and duplicates, are removed. However, refactoring typically requires an experienced coder and can be risky with larger programs. Modifying code can cascade into serious problems and unexpected changes in output can and often arise.

The original stock analysis macro performed significantly slower than its refactored counterpart. The original macro used more conditional statements while the refactored version used arrays and cut down on conditional statements. The original VBA scripts ran in approximately 1.5 seconds while the refactored versions of the 2017 and 2018 data ran in approximately 0.2 seconds. This 1.3 second difference would prove critical in real-time applications and demonstrates the positive and often necessary aspects of refactoring.

# An Analysis of Green Energy Stocks

## Overview of Project

Analyze and automate reporting for Green Energy stock data on behalf of Steve and his parents 

### Purpose

Perform historical annual return analysis on Green Energy stocks and prepare reporting that can be automated and refactored to analyze the entire stock market    

## Results

Stock performance in 2017 was better overall compared to 2018 with 11 of 12 stocks posting positive returns for the year versus only two stocks with positive returns in 2018. Additionally, only two stocks performed better in 2018 than in 2017.

Automating the historical annual return analysis and reporting using VBA led to significant time saving over performing manual filtering and calculations on eight columns and 3,013 rows of data, with the entire macro-driven process taking only 0.8750 seconds for 2017 and 0.8516 seconds for 2018 to complete.

Refactoring the original VBA script led to additional time saving with the entire process now only taking 0.1875 seconds for 2017 and 0.1797 seconds for 2018 to complete, or 4.7x faster for each year.

[VBA_Challenge](https://github.com/JemiShieh/stock-analysis/VBA_Challenge.xlsm)

![VBA_PreChallenge_2017](https://github.com/JemiShieh/stock-analysis/Resources/VBA_PreChallenge_2017.png)

![VBA_PreChallenge_2018](https://github.com/JemiShieh/stock-analysis/Resources/VBA_PreChallenge_2018.png)

![VBA_Challenge_2017](https://github.com/JemiShieh/stock-analysis/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018](https://github.com/JemiShieh/stock-analysis/Resources/VBA_Challenge_2018.png)

## Summary

* What are the advantages or disadvantages of refactoring code?

  1. Potential advantages of refactoring code include better organization, readability, understandability, and efficiency, easier maintenance, updating and debugging, and faster run times.
 
  2. Potential disadvantages of refactoring code include risk of introducing bugs, and opportunity cost of time and money spent versus time and money saved without introducing any additional functionality.

* How do these pros and cons apply to refactoring the original VBA script?

  1. Advantages of refactoring the original Module2_VBA_Script include using indexing and output arrays to allow for easier updating, more efficiency, and faster run times.

  2. Disadvantages of refactoring the original Module2_VBA_Script include the opportunity cost of five hours of time spent versus the less than 0.69 seconds in saved run time, without introducing any additional functionality.

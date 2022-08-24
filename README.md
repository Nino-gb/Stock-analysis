 Stock Analysis Challenge

## Overview of Project 

Steve's first client want to invest all their money in a green energy stocks called DAQO. Before they make any investment Steve want to analyze a handful of green energy stock data in addition to DAQO stocks, and  compare current financial statement to others stocks to give the investors a sense of whether DAQO is growing, stable, or deteriorating. In order to analyze the stocks data he need to know how often the stock is been traded.  This analysis included stocks name, total daily volume and the year return during 2017 and 2018. To perform this analysis we used Vision Basics for Application (VBA), a programing language that interacts directly with excel. Although the codes created in VBA work well it may not work well with a thousands of stock data and may take a long time to execute.

### Purpose

The purpose of this project is to edit, or refactor the existing All Stocks Analysis VBA codes to improve their design, structure and implementation without altering the code's external behavior. We will add new functionality to make the codes more efficient by taking fewer steps, using less memory, and improving structure behavior. After refactoring we will be able to determinate if the codes successfully made VBA script run faster, and analyze stock performance between 2017 and 2018.


## Results

![Stock Analysis Refactor](images/ScreenShot2022-08-23at7.50.25PM.png)

### Codes Performance

Refactoring the Original Script we were able to;

- Eliminate useless code, reduce that amount of code use in the script.

- Simplifies codes by breaking it up into a chunk of loops (2a,2b,4).

- Reduce the amount of if-then statement from 3 to 2 (3b,3c).

- Loop through all the data one time and collect the same information of the original script (2b,3a,3b.3c,3d).

- Eliminated nest loop. 

- Refactoring structure improve consistent indentation, thus the refactoring script look more clean.


### Compare Stock performance 

![Stock Analysis 2017 and 2018](/Users/ninotshkacruz/Desktop/ScreenShot2022-08-23at7.32.59PM.png)

To perform this valuation we analyze the financial performance of twelve green energy stock during 2017 and 2018. We calculate the total volume and return for each year. Using the data in Worksheet All Stock Analysis we can visualize the following;

Stocks 2017 -  11 out of 12 tickers had a positive year return with a high total volume.

Stocks 2018- 10 out of 12 tickers had a negative year return with a high total volume.

Only ENPH and RUN tickers have a high volume total and a positive return for 2017 and 2018 period. We think ENPH and Run represent a low risk investment due to their patterns of large number of outstanding shares,  active stocks and yearly return. Investing in ENPH and RUN stock may offer several benefits, including the potential to earn dividends or an average annualized return.


### Execution times between Original Script and the Refactored script

![Stock 2017, Original Script vs Refactored Script](/Users/ninotshkacruz/Desktop/Challenge 2/Screen Shot 2022-08-23 at 7.06.53 PM.png)
2017 Refactored Script code run 0.2421875 less than Original Script, therefore Refactoring code run faster than the Original Script


![Stock 2018, Original Script vs Refactored Script](/Users/ninotshkacruz/Desktop/Challenge 2/Screen Shot 2022-08-23 at 7.08.23 PM.png)
2017 Refactored Script code run 0.234375 less than Original Script, therefore Refactoring code run faster than the Original Script


## Summary

1. What are the advantage and disadvantages of refactoring code?

### Advantage 

- Improve the design of exiting code

- Restructuring the source code is possible without altering the functionality

- Removed redundancy and duplications improve the effectiveness of the effectiveness of the code

- Clean code with shorter, self contained methods and classes is is characterized by better testability

### Disadvantage  

- Imprecise refactoring could introduce new bugs end errors into the code.

- In the case of larger teams working on refactoring, the coordination effort required could be high.

- Is risky when developers do not understand the project.

- May take a long time to complete the process.


2. How do these pros and cons apply to refactoring the original VBA?

### Pro 

- Refactoring improve the design of the existing code, the refactored script have more consistent indentation and avoid nesting loops. 

- Reduce redundancy and duplication by creating a code that loop through all data one time.

- Restructuring the source code was possible and run faster without altering the code external behavior.

- Clean codes creating a chunk of loops instead using a long loop with a nest loop.

- Having a good understanding of the original script avoid the risk of failure during refactoring.

### Con

- Inexperience led to create several imprecise codes and introduce several error during the process of refactoring.

- The refactoring task was assigned to one person, therefore coordination through teams was not necessary.

- Refactoring took several hours to complete





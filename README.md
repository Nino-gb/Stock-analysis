# Stocks Analysis Challenge

## Overview of Project 

Steve's first client want to invest all their money in a green energy stocks called DAQO. Before they make any investment Steve want to analyze a handful of green energy stock data in addition to DAQO stocks, and  compare current financial statement to others stocks to give the investors a sense of whether DAQO stocks are growing, stable, or deteriorating. Before Steve make a recommendation, he need to know how often the stock is been traded. To perform this analysis we used Vision Basics for Application (VBA), a programing language that interacts directly with excel. This analysis included the stocks name, total daily volume and the return during 2017 and 2018. Although the codes created in VBA work well it may not work well with a thousands of stock data and it may take a long time to execute.

### Purpose

The purpose of this project is to edit, or refactor the existing All Stocks Analysis VBA codes to improve their design, structure and implementation without altering the code's external behavior. We will add new functionality to make the codes more efficient by taking fewer steps, using less memory, and improving structure behavior. After refactoring we will be able to determinate if the codes successfully made VBA script run faster, and analyze stock performance between 2017 and 2018.


## Results

> *All Stocks Analysis Refactored VBA Codes*

![Screen Shot 2022-08-23 at 7 50 25 PM](https://user-images.githubusercontent.com/110786136/186437641-a204dadc-99f0-441a-8906-819716fdc2f7.png)


### Codes Performance

When we refactored the Original Script we were able to;

- Eliminate useless code, reduce that amount of code use in the script.

- Simplifies codes by breaking it up into a chunk of loops (2a,2b,4).

- Reduce the amount of if-then statement from 3 to 2 (3b,3c).

- Loop through all the data one time and collect the same information of the original script (2b,3a,3b.3c,3d).

- Eliminated nest loop. 

- Improve consistent indentation, thus the refactored script look more clean.


### Compare Stocks performance 

> *Stocks Analysis 2017 and 2018*

![Screen Shot 2022-08-23 at 7 32 59 PM](https://user-images.githubusercontent.com/110786136/186449499-13425354-d183-4f4a-9f00-e45c8483a3b8.png)

To perform this valuation we analyze the financial performance of twelve green energy stock during 2017 and 2018. We calculate the total volume and return for each year. Using the data in Worksheet All Stock Analysis we can visualize the following;

Stocks 2017 -  11 out of 12 tickers had a positive return with a high total volume.

Stocks 2018- 10 out of 12 tickers had a negative return with a high total volume.

Only ENPH and RUN tickers have a high volume total and a positive return for 2017 and 2018 period. We think ENPH and Run represent a low risk investment due to their patterns of large number of outstanding shares,  active stocks and returns. Investing in ENPH and RUN stock may offer several benefits, including the potential to earn dividends or an average annualized return.


### Execution times between Original Script and the Refactored script

> *2017 Refactored Script code ran 0.2421875 less seconds than Original Script, therefore refactored codes run faster than the original script*

![Screen Shot 2022-08-23 at 7 06 53 PM](https://user-images.githubusercontent.com/110786136/186449834-0958c74b-cf3f-4c44-8c45-c29b3a874c03.png)



> *2018 Refactored Script code ran 0.234375 less seconds than Original Script, therefore refactored codes run faster than the original script*

![Screen Shot 2022-08-23 at 7 08 23 PM](https://user-images.githubusercontent.com/110786136/186449961-76a1fde3-fb96-4b17-8b48-2443a5d31107.png)

## Summary

### What are the advantages and disadvantages of refactoring code?

#### *Advantages* 

- Improve the design of exiting code

- Restructuring the source code is possible without altering the functionality

- Removed redundancy and duplications improve the effectiveness of the effectiveness of the code

- Clean code with shorter, self contained methods and classes is is characterized by better testability

#### *Disadvantages*  

- Imprecise refactoring could introduce new bugs end errors into the code.

- In the case of larger teams working on refactoring, the coordination effort required could be high.

- Is risky when developers do not understand the project.

- May take a long time to complete the process.


### How do these pros and cons apply to refactoring the original VBA?

#### *Pros*

- Refactoring improve the design of the existing code, the refactored script have more consistent indentation and avoid nesting loops. 

- Reduce redundancy and duplication by creating a code that loop through all data one time.

- Restructuring the source code was possible and run faster without altering the code external behavior.

- Clean codes creating a chunk of loops instead using a long loop with a nest loop.

- Having a good understanding of the original script avoid the risk of failure during refactoring.

#### *Cons*

- Inexperience led to create several imprecise codes and introduce several error during the process of refactoring.

- The refactoring task was assigned to one person, therefore coordination through teams was not necessary.

- Refactoring took several hours to complete





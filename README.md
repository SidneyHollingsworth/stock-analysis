# VBA Stock Analysis

## Overview of Project
I've been given the task of reviewing stock performance of twelve green energy stocks. Rather than tabulating x and y manually, I've been provided starter code for a VBA script. However, in anticipation of increasing the number of stocks to analysing from a dozen to thousands, I've been additionally asked to refactor the code in order to reduce run time. 

- This project came about from a need to review stock performance and further, refactor the code to reduce run time. 

## Results

### 2017 Stock Performance
![2017 Stock Performance](https://github.com/SidneyHollingsworth/stock-analysis/blob/0ca17f30b627400ccf14fad89a18d47471a4f2f6/Resources/VBA_Challenge_2017_Output.png)

### 2018 Stock Performance
![2018 Stock Performance](https://github.com/SidneyHollingsworth/stock-analysis/blob/0ca17f30b627400ccf14fad89a18d47471a4f2f6/Resources/VBA_Challenge_2018_Output.png)

## Summary 

#### Year to Year Comparison
![Year to Year Comparison](https://github.com/SidneyHollingsworth/stock-analysis/blob/0ca17f30b627400ccf14fad89a18d47471a4f2f6/Resources/Change_Over_Year.png)

- [Summarize comparison of stock performance between 2017 and 2018.]

#### Execution times of original script vs refactored script

Refactoring the VBA All Stocks Analysis script sped up run time for 2017 more than 2018.
- Original run time for 2017 was 0.1328125 seconds. Refactoring sped up 2017 run time by 0.0548125 seconds. (Screenshots below)
- Orginal run time for 2018 was 0.1015625 seconds. Refactoring sped up 2018 run time by 0.0155625 seconds. (Screenshots below)

##### 2017 Starting Run Time
![2017 Starting Run Time](https://github.com/SidneyHollingsworth/stock-analysis/blob/0ca17f30b627400ccf14fad89a18d47471a4f2f6/Resources/2017_Starting_Run_Time.png)

##### 2018 Starting Run Time
![2018 Starting Run Time](https://github.com/SidneyHollingsworth/stock-analysis/blob/0ca17f30b627400ccf14fad89a18d47471a4f2f6/Resources/2018_Starting_run_time.png)

##### 2017 Refactored Run Time
![2017 Refactored Run Time](https://github.com/SidneyHollingsworth/stock-analysis/blob/0ca17f30b627400ccf14fad89a18d47471a4f2f6/Resources/VBA_Challenge_2017.png)

##### 2018 Refactored Run Time
![2018 Refactored Run Time](https://github.com/SidneyHollingsworth/stock-analysis/blob/0ca17f30b627400ccf14fad89a18d47471a4f2f6/Resources/VBA_Challenge_2018.png)

### Advantages and Disadvantages of Refactoring Code
- [Explination in general]
- [How pro/cons apply to project. Example of advantages and dissavantages of refactoring original and refactored VBA script.]
- I was able to take the VBA code run time from 1.18 seconds down to 0.1 seconds. 
- The main gain was by using manual calculation rather than automatic. The other piece is doing away with worksheet.activate.



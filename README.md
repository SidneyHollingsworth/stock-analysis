# VBA Stock Analysis
 

## Overview of Project
This project came about from a need to review stock performance and further, refactor the code to reduce run time. I've been given the task of reviewing stock performance of twelve green energy stocks. Rather than tabulating Total Daily Volume and Return manually, I've been provided starter code for a VBA script. In anticipation of increasing the number of stocks to analysing from a dozen to thousands, I've been additionally asked to refactor the code in a way which would reduce run time. 

## Results

### 2017 Stock Performance
![2017 Stock Performance](https://github.com/SidneyHollingsworth/stock-analysis/blob/0ca17f30b627400ccf14fad89a18d47471a4f2f6/Resources/VBA_Challenge_2017_Output.png)

### 2018 Stock Performance
![2018 Stock Performance](https://github.com/SidneyHollingsworth/stock-analysis/blob/0ca17f30b627400ccf14fad89a18d47471a4f2f6/Resources/VBA_Challenge_2018_Output.png)

## Summary 

#### Year to Year Comparison
![Year to Year Comparison](https://github.com/SidneyHollingsworth/stock-analysis/blob/5386966c36c3546338fca4205506c2897c8285e2/Resources/Change_Over_Year.png)

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
There are advantages and disadvantages to refactoring code. Generalized examples include X and Y. In regard to this project, the hours I spent refactoring the All Stock Analysis VBA code only saved a partial second of run time.

- The main gain was by using manual calculation rather than automatic. The other piece is doing away with worksheet.activate.



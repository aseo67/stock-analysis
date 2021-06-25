# VBA of Wall Street
Module 2 Challenge
Data Analysis File Link: 
  ![VBA Analysis](https://github.com/aseo67/stocks-analysis/blob/main/VBA_Challenge.xlsm.zip)

## Overview of Project

### Background/Purpose
Given Steve's background and recently acquired finance degree, his first clients - his parents - are seeking advice on their stock investments. Currently, they have invested all their money into DAQO (ticker: DQ), though they had not conducted much research heading into the investment. In addition to analyzing the performance of DAQO stock, they'd like to find ways to diversify their funds. Through this project, we will be automating the analysis of yearly stock data to create quick, accurate analyses for Steve and his parents, reproducible for any given set of stocks. 

## Results

### Stock Performance 2017 vs. 2018
While DAQO performed well in 2017 with nearly 200% return, its current performance is poor, dropping over 60% in 2018 alone. This confirms that reallocation and diversification of Steve's parents' investment is needed. A variety of other stock options were also analyzed for performance in 2017 and 2018. Of these, the stocks that demonstrate potential are tickers "ENPH" and "RUN". While most stocks had performed well in 2017, this didn't carry through in 2018 for any except the two aforementioned stocks. 
For "ENPH", they achieved strong return of ~130% in 2017, and this continued with ~80% growth in 2018, standing out as the stock with the highest total daily volume as of 2018.
For "RUN", although they only achieved a return of ~5% in 2017, this has improved to almost ~85% return in 2018. 
  
  **2017 Stock Performance Analysis Results**
  ![Screenshot](https://github.com/aseo67/stocks-analysis/blob/main/VBA_Challenge_Table%202017.png)
  
  **2018 Stock Performance Analysis Results**
  ![Screenshot](https://github.com/aseo67/stocks-analysis/blob/main/VBA_Challenge_Table%202018.png)
  

### Execution Time
In the course of this analysis, the code was refactored from its original script to help optimize the runtime and alleviate burden of running this analysis on larger selections of stocks. 

  #### Pre-Refactor Code
  Prior to the refactoring, the analysis was run then outputted, on a row-by-row (or stock-by-stock) basis. This resulted in runtimes ~0.5seconds.  
  
  ![Screenshot](https://github.com/aseo67/stocks-analysis/blob/main/VBA_PreChallenge%20Refactor%20Code.png)
  
  **2017 Original VBA Script**
  ![Screenshot](https://github.com/aseo67/stocks-analysis/blob/main/VBA_PreChallenge%20Refactor_2017.png)
  
  **2018 Original VBA Script**
  ![Screenshot](https://github.com/aseo67/stocks-analysis/blob/main/VBA_PreChallenge%20Refactor_2018.png)
  
  #### Refactored Code
  The code was then refactored to help optimize the calculation/output process by first collecting and storing all the array values, then outputting the stored data into the table. This resulted in runtimes ~0.1seconds, expediating the analysis considerably vs. the original script. 
  
  ![Screenshot](https://github.com/aseo67/stocks-analysis/blob/main/VBA_Challenge%20Refactored%20Code.png)
 
  **2017 Refactored VBA Script**
  ![Screenshot](https://github.com/aseo67/stocks-analysis/blob/main/VBA_Challenge_2017.png)
  
   **2018 Refactored VBA Script**
  ![Screenshot](https://github.com/aseo67/stocks-analysis/blob/main/VBA_Challenge_2018.png)


## Summary

### What are the advantages or disadvantages of refactoring code?


### How do these pros and cons apply to refactoring the original VBA script?


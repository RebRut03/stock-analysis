# **Title of stock-analysis**

## **Overview of Project**
### Purpose - Explain the purpose of this analysis. The purpose of this project was to create some code for Steve do some stock analysis on behalf of his parents.  Although he originally looked at green stocks, he wants the option to anlysis the entire stock market over the last few years to assist his parents in determining their investments.

Specifically for this portion of the project, I will refactor the code with the intention of making the VBA script run faster but keeping the same analysis functionality.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

## **Results**
### Stock Performance Comparison - Steve wants to find the total daily volume and yearly return for each stock. The yearly return is the percentage difference in price from the beginning of the year to the end of the year.  Using images and examples of your code, compare the stock performance between 2017 and 2018, 

- Total Daily Volume - Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. 
- To calculate DQ's total daily volume, we need to loop through all of the stocks, so we've typed the number of rows into the code itself. What would be even better, though, is to use VBA to find the number of rows to loop over. Unfortunately, VBA doesn't have a nice function or method to figure that out. But we can't be the first person to have this problem; someone must have found a solution. 
- 
- Yearly return - Steve wants to know how DQ performed in 2018. One way to measure this is to calculate the yearly return for DQ. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. In other words, if you invested in DQ at the beginning of the year and never sold, the yearly return is how much your investment grew or shrunk by the end of the year. Daqo dropped over 63% in 2018—yikes! Steve will definitely want to offer some better stocks to his parents.


**2017 Stock Analysis:**

  ![Stock_Table_2017](Resources/Stock_Table_2017.PNG)
  
**2018 Stock Analysis:**

  ![Stock_Table_2018](Resources/Stock_Table_2018.PNG)


- As a whole, the green stocks included in Steve's list droppe
- I would suggest that he steer his family in a different direction.
### Script Execution Time - In the future, Steve may want to perform his analysis on larger datasets, and he wants to know how fast his VBA code will compile the results.

  - Refactored Script
    - Using images and examples of your code, compare the execution times of the original script and the refactored script make sure links are working!!!!
    ![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)
      
  
  - Original Script 
    -  blah blah 
  
      ![Original_2017](Resources/Original_2017.PNG)
      ![Original_2018](Resources/Original_2018.PNG)
  
## **Summary**
### Advantages or Disadvantages of Refactoring Code
- One advantage of refactoring might be that the coder is able to improve the efficiency of the code.  
    -  Three examples given in Module 2 of how code can be made more efficient include:
      
      > taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.
      
    -  In this project, blah blah
 - Possible disadvantages might be the opposite of what was just mentioned.  If you are not the original author of the code, you may spend more time attempting to become familiar with the code or you may not have the background.  Perhaps the original coder might have scripted the code in what appears to be a more efficient way, and it didn't create the same output.
### Application of Refactoring Original Script
- How do these pros and cons apply to refactoring the original VBA script?

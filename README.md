# **Steve's Stock Analysis Tool**

## **Overview of Project**
  - ### Purpose - 
    The purpose of this project is to create a tool for Steve to analyze stock daily volume and yearly return.  He will use this information to assist his parents with their investment strategy.  Although Steve originally focused his efforts on green stocks, he wants the option to anlyze all types of stocks over the last few years.  Specifically for this portion of the project, I will refactor the original code I provided Steve with the intention of making the VBA script run faster and more efficiently while keeping the same analysis functionality and output.

## **Results**
  - ### Stock Performance Comparison - 
    Steve would like this tool to calculate the total daily volume and yearly return for each stock to get a better idea of what stocks he will advise his parents to invest in. According to Module 2 material:
  
   > Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. And the yearly return is the percentage difference in price from the beginning of the year to the end of the year.
 
   
   **2017 Stock Analysis:**
  
   ![Stock_Table_2017](Resources/Stock_Table_2017.PNG)
  
  **2018 Stock Analysis:**
  
   ![Stock_Table_2018](Resources/Stock_Table_2018.PNG)


   - Referring to the tables above, you'll see the total daily volumes and yearly returns of the green stocks that Steve originally wanted analyzed.  In 2017, the returns for these stocks, with the exception of TERP (-7.2%), were positive; however, in 2018, many of the stocks' returns drop drastically, with the exception of ENPH (+81%) and RUN (+84%).  This obviously led to Steve's desire to diversify his anlysis to include other types of stock to find a wiser investment direction for his parents.

  - ### Script Execution Time - 
    In the future, Steve may want to perform his analysis on larger datasets so he wanted to know how fast his VBA code will compile the results. I refactored the orignal code to make it run more efficiently.  Also, as you can see by the following screen shots, the refactored script's execution time is lower than the original script.  As I continue to run the refactored code, the execution time seems to become quicker and more efficient.

    - **Refactored Code Execution Time**
          
        ![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)
    
        ![VBA_Challenge_2018](Resources/VBA_Challenge_2018.PNG)
      
      - Some examples of how I refactored the code to make it run more efficiently include:
        1) creating a ticker index: `Dim tickerIndex As String tickerIndex = tickers(0)`,
        2) creating output arrays for ticker volume, starting price, and ending prices:  `Dim tickerVolumes(12) As Long Dim tickerStartingPrices(12) As Single Dim tickerEndingPrices(12) As Single`
        3) creating a For loop to initialize the arrays to zero: `For i = 0 To 11 tickerIndex = tickers(i) tickerVolumes(i) = 0 tickerStartingPrices(i) = 0 tickerEndingPrices(i) = 0`,
        4) creating a for loop to loop through all rows ` For j = 2 To RowCount`  One example in this for loop of how I increaded the volume for the current ticker is `If Cells(j, 1).Value = tickerIndex Then tickerVolumes(i) = tickerVolumes(i) + Cells(j, 8).Value End If`, and 
        5) creating a For loop to loop through the arrays and output the ticker, total daily volume, and return. `For i = 0 To 11 Worksheets("All Stocks Analysis").Activate Cells(4 + i, 1).Value = tickers(i) Cells(4 + i, 2).Value = tickerVolumes(i) Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1`
  
    - **Original Code Exection Time** 
        
       ![Original_2017](Resources/Original_2017.PNG)
       
       ![Original_2018](Resources/Original_2018.PNG)
  
## **Summary**
  - ### Advantages or Disadvantages of Refactoring Code in General
    - An advantage of refactoring could be that the coder is able to improve the efficiency of the code.  
      -  Three examples given in Module 2 of how code can be made more efficient include:
      > taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.
    - Possible disadvantages of refactoring code might include spending an inordinate amount of time familiarizing yourself with someone elses code and not knowing the history behind why it was coded in the way that it is, could be an obstacle.  Perhaps, the original coder tried it the way you are refactoring and it created a snowball effect or more obstacles.  If the original code is ridden with bugs or issues, it might be more fruitful to simply rewrite the code instead of refactoring it.
  - ### Pros and Cons of Refactoring Original Script
    - As displayed in the screenshots above, the efficiency of the script was improved.  The addtions to the script made it run faster, and the addtion of the arrays make it more flexible for future analysis of larger data sets.
    - I don't feel the disadvantages listed above apply to this particular script because it was working without errors, and I was the original coder so I was familiar with the inerworkings of it.




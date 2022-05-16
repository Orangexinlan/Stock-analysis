# Stock-analysis

## Overview of Project

### Purpose and Background
Steve just graduated with his finance degree,his parents wanted to invest in Green Energy. Without doing lots of research, they were going to put all the money in a company call Daqo. Steven is going to look into the stock for his parents, but he's concerned about diversifying their funds. He wants to analyst a handful Green EnergyStocks in additions. He created an Excel with all the information.

So we created a code with VBA to automate analyses allows Streve to reuse it with any stock, and reduces the chance of accidents and errors.I'll refactor, the code to loop through all the data one time in order to make the VBA script run faster.

## Analysis and Results

## Results
The two images below showed how efficient this refactor code could be.
Here's the code on how this image pop up.
   
   endTime = Timer
    
  MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


![this is an image](https://github.com/Orangexinlan/Stock-analysis/blob/47576c9fe175d255348ab553c0d4ded27872ee79/Resources/VBA_Challenge_2017.png)
![this is an image](https://github.com/Orangexinlan/Stock-analysis/blob/47576c9fe175d255348ab553c0d4ded27872ee79/Resources/VBA_Challenge_2018.png)

### Analysis
From two images listed below, we can tell that in 2017, only one company had negative year return. But in 2018, nearly 83% of the company has negative year return.

![this is an image](https://github.com/Orangexinlan/Stock-analysis/blob/92dce54850e1d202a4f453109ba85baaee3ad573/Resources/2017%20Return.png)
![this is an image](https://github.com/Orangexinlan/Stock-analysis/blob/92dce54850e1d202a4f453109ba85baaee3ad573/Resources/2018%20Reurn.png)

## Summary

### the advantages or disadvantages of refactoring code

#### the advantages
Efficiency-- With the Refactor Code, Excel is taking less time to excute the code. Since the functionality stays the same.

#### the disadvantages
Frustrating-- The code is often difficult fot costomers to recgnize

### How do these pros and cons apply to refactoring the original VBA script
Determine the refactoring successfully made is to see if the VBA script run faster. 
By applying the refactor code, the VBA runs around Five times faster than the original one. That means the refactor code we wrote is actually working. 
But in the meantime it's also frustrating. We need to figure out what exatly each code means in the steps. It is so easy to confuse with the original code.

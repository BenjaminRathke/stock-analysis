# VBA of Wall Street

## Overview of Project
This project utilizes Visual Basic for Applications (VBA) to refactor code, mostly already written, for Module 2.  Ultimately, the purpose of the code is to output data to show the total daily volume and single-year rate of return for 12 different stocks over the years 2017 and 2018.  With this information, end users may evaluate which stock may be a more sound investment.  This project illustrated and taught the use of for loops, nested loops, if/then and if/elseif/else coding, creating and defining variables, indices and arrays, amongst myriad other VBA tools and tricks.

## Results

### Stock Performance

With the exception of stocks ENPH and RUN, all stocks' had a negative rate of return for the year 2018.  However, it should be noted that the decrease in return in 2018 for most stocks would not have necessarily completely erased the returns from 2017.  Instead, one could look at the data to evaluate the most stable stocks (had the smallest loss rates in 2017 such as VSLR and TERP) or the stocks for companies that are in a multi-year growth pattern, such as ENPH and RUN.

[INSERT PHOTOS OF STOCK PERFORMANCE]

### Script Execution Times

Upon refactoring, there was a significant reduction in script execution times.  Refactoring resulted in making the code far more efficient.  

Utilizing a tickerIndex variable in several formulae allowed us to use this variable for multiple outputs instead of creating a new variable for each output (ticker string value, total daily volume, and rate of return).

[INSERT CODE FOR tickerIndex VARIABLE CREATION AND USE]

Perhaps the most important change to the code during refactoring was the utilization of a simple if/then statement that increased the ticker index and started the loop with the next tickerIndex before the entire data set was looped over unnecessarily.  This was doable since the data already had the ticker symbols conveniently grouped together, but any data set can be reorganized to do this prior to writing the VBA script.

[INSERT CODE FOR ENDING FOR LOOP]

## Summary

### Advantages and Disadvantages of Refactoring Code

The advantages of refactoring code are relatively straightforward.  When code is initially created, it may not have been coded as efficiently as it could be.  There are multiple ways of getting a certain output from a script or program, but some ways may be more efficient than others.  While modern computers are powerful devices, those who create these scripts and programs should strive for the most efficient code possible to minimize resource usage.  Refactoring can also make code more accessible and easier to read and follow by others in a team environment.  Finally, refactoring will always challenge the programmer to learn new tips, tricks, and resources for any project.
The disadvantages of refactoring include the time it takes to refactor in a meaningful and impactful way within a reasonable amount of time.  Exact placement of lines of code, exact syntax, and unexpected results when the programmer is certain of a different outcome can be (and were) all experienced, though this is not unique to VBA or this project.  These experiences can be extremely frustrating.  

### Advantages and Disadvantages of Refactoring for This Specific Project

The advantages of refactoring in this project include learning a tolerance for trial and error, self-directed investigation into why a code does or does not work, creating a more efficient way of looking through and sorting data to conserve CPU resources, and an opportunity to format and edit for clean and easily readable and followable code.  Utilizing the “outside the box” tricks to stop unnecessary processes or remove unnecessary code was a powerful lesson learned.
The disadvantages are primarily time.  It seemed far easier and faster to create the code as it existed in the original module.  However, this code was not efficient as it looked over the entire data set when that was objectively unnecessary and led to longer program run times.  Another disadvantage is that code that is essentially working can be broken in a refactoring process.  Care and attention to the original code must be utilized in refactoring and backing up and saving multiple iterations of the code is a must.

# stock-analysis
## Overview
The purpose of this analysis is to provide capabilities to visualize the yearly return on a wide array of company stocks. In turn, we can then help clients visualize stock data and help them determine which companies they might want to invest in.
We pulled data for twelve separate companies, all identifiable by unique tickers. The data provided to us was formatted in rows. Each row provided the ticker, the date corresponding to when data was collected, the opening, high, low, and closing prices, as well as the adjusted close price, and finally the sheer volume of each stock traded on that exact date.
## Results
Apparently, 2017 was generally a better year for the twelve stocks we kept tabs on. All but one ticker ended up providing positive returns by the end of the year. However, in 2018, all stocks but two ended with negative returns. So, we can take away a good lesson when it comes to the stock market: movement is unpredictable - even two years can end up looking totally different from one another.
Using the *startTime = Timer* line of code in the start of my code, and following up using *endTime = Timer*, I was able to set variables to keep time for me as my code performed its analysis. I could return the actual amount of time it took to run by completing my code with the following line:

'MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)'

Provided here are the two runtimes for analysis performed on the data.
### Running Code on 2017 Data
![Duration for Running Code on 2017 Data](VBA_Challenge_2017.png)
### Running Code on 2018 Data
![Duration for Running Code on 2018 Data](VBA_Challenge_2018.png)

## Summary
### Refactoring Code: Advantages and Disadvantages
The advantages of refactoring code are clear. Refactored code is moulded into a state in which it becomes easier to read and simple to follow. Refactored code includes all the bits that were created along the way, from the first successful output to the final desired graphics. In essence, refactored code is a *final draft* as opposed to the messier nature of a *rough draft*; the lines of code are synthesized into one, coherent conglomerate of code after having chipped away little ideas bit by bit in one's original code.
Some disadvantages of refactoring code is that the "original train of thought" isn't necessarily being present any more. Even though refactored code becomes more clear-cut and concise, a non-author reader of the code isn't going to see what the author(s) were originally thinking along the lines of when they began their code. This may become disadvantageous if the refactored code is actually inadequate. Along this same vein, the author who refactors the code might unknowingly cause a reader more confusion by omitting critical steps along the way of writing the finished product.
### Refactoring Code: Applied
Specifically regarding the refactoring of code I had my code undergo in VBA, I believe the same ups and downs I elaborated on above apply. Though easier to follow and read, and only filled with absolutely relevant ideas, I could agree to an extent that my refactored code is scant around how I got to think about things, along the way. Though, I do believe what I finished with is entirely comprehensive for seeing my thought-process as I completed my project.
## Overview of Stock Analysis Project

The premise for this project was to assist a client in conducting stock analysis in scalable, efficient manner. The objective of this challenge was to build VBA Macros to efficiently analyze stock data from 2017, and 2018 for 12 stocks. The analysis included the Total Volume, and Total Return based on open and close pricing for each individual stock, on a specified year. During this challenge, we refactored the code to run more efficently utilizing a TickerIndex for the individual stocks, as well as three output arrays for TotalVolume, TickerStartingPrice and TickerEndingPrice to remove the need for duplicative code. Lastly, we applied code to clearly format the code to easily identify performance and total volume by individual ticker in an easily digestible manner.

## Review of Results

### Stock Performance
The following analysis was conducted on the 12 indiviual stocks: AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, VSLR. Generally speaking, 2017 saw strong performance for the majority of stocks in our sample set, recognizing an annualized return of of 5% to nearly 200%, with the exception of one stock, $TERP which realized a -7.2% return for the year. The average annual trading volume and average return were 263M Shares, and 67.3% (respectively). By contrast, 2018 proved to be a far more challenging year from a performance perspective (average performance: -8.5%), with only two tickers generating a positive return, $ENPH and $RUN. Broadly speaking, trading volumes of the data set as a whole, was relatively comparable to 2017. Average Trading Volume in 2018 was 275M shares. The belwo chart depicts a side by side comparison of results for each ticker on a given year:

<img width="270" alt="2017 Results" src="https://user-images.githubusercontent.com/80016496/112729203-02ce2f80-8ef9-11eb-8333-83e8ee2c3147.png">

<img width="269" alt="2018 Results" src="https://user-images.githubusercontent.com/80016496/112729208-0792e380-8ef9-11eb-94bf-97b52c4d8da8.png">

### VBA Code-Related Results
In order to conduct this analysis more efficiently, we undertook a refactoring exercise to streamline the code, and run our analysis more expiditiously. The below screenshots show the efficiency differience in the original analysis conducted in the course work, and the refactored version. 

Original 2017

<img width="422" alt="Original 2017" src="https://user-images.githubusercontent.com/80016496/112735206-88ada300-8f18-11eb-80f9-bbb9c893a55e.png">

Refactored 2017

<img width="424" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/80016496/112735228-9f53fa00-8f18-11eb-9487-3c004c7c8d1b.png">

Original 2018

<img width="418" alt="Original 2018" src="https://user-images.githubusercontent.com/80016496/112735299-ffe33700-8f18-11eb-815b-7aec3d7cad03.png">

Refactored 2018

<img width="418" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/80016496/112735306-0a9dcc00-8f19-11eb-8c59-02947d8b54c0.png">

### Discussion of Refactoring Code
Refactoring code provides several benefits. The most notable benefits of refactoring code include improving performance/run-time (as is the case with this analysis), but also provide developers the opportunity to improve/streamline the overall efficiency of the code. Often times, as coding languages are improved, and subsequent releases/patches are made, new logic and functions are available. Refactoring code gives us the ability to apply those new and improved functions and logic in a more streamlined manner. Refactoring also affords developers a "fresh set of eyes" on code that may have been written in the past. Of course, with benefits, come some risks associated with refactoring code. Most notably, if done in a production environement without detailed commenting and strong version controls developers run the risk of user error and implementing new bugs to the code. In addition to presenting new errors in the code, there is also an opportunity cost to consider. Refactoring code allocates development resources to problems that in theory are already solved. Strategically it is important to consider the value of refactoring intiatives, verse addressing new projects that may add value to the firm. Refactoring initiatives vs. new projects should be evaluated at a senior level.

### Benefits of Refactoring Code for this Analysis
The most important component of this refactoring project is the implementation of a TickerIndex and three subsequent output arrays created for TickerVolume, TickerStartingPrice, and TickerEndingPrice. These elements allow for a more concise analysis for the conditionals and for loops. Instead of completing the entire Sub Process one ticker at a time, or refactored version allows the code to reference the tickers, and underlying data associated using the TickerIndex in the conditionals/for loop further down in the process. In a sense, these arrays offer the logic and "express lane". The only notable downside to this refactored analysis is that the initial code already ran in a reasonable amount of time (< 1 second), so the efficiency gain (other than for educational purposes) is nominal. 

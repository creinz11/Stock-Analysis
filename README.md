## Overview of Stock Analysis Project

The purpose of this challenge was to build VBA Macros to efficiently analyze stock data from 2017, and 2018 for 12 stocks. The analysis included the Total Volume, and Total return based on open and close pricing for each individual stock, on a specified year. During this challenge, we refactored the code to run more efficently utilizing a TickerIndex for the indivudal stocks, as well as three output arrays for TotalVolume, TickerStartingPrice and TickerEndingPrice to remove the need for duplicative code. Lastly, we applied code to clearly format the code to easily identify performance and total volume by individual ticker in an easily digestible manner.

## Review of Results

### Stock Performance
The following analysis was conducted on the 12 indiviual stocks: AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, VSLR. Generally speaking, 2017 saw strong performance for the majority of stocks in our sample set, recognizing an annualized return of of 5% to nearly 200%, with the exception of one stock, $TERP which realized a -7.2% return for the year. The average annual trading volume and average return were 263M Shares, and 67.3% (respectively). By contrast, 2018 proved to be a far more challenging year from a performance perspective (average performance: -8.5%), with only two tickers generating a positive return, $ENPH and $RUN. Broadly speaking, trading volumes of the data set as a whole, was relatively comparable to 2017. Average Trading Volume in 2018 was 275M shares. The belwo chart depicts a side by side comparison of results for each ticker on a given year:

<img width="270" alt="2017 Results" src="https://user-images.githubusercontent.com/80016496/112729203-02ce2f80-8ef9-11eb-8333-83e8ee2c3147.png">

<img width="269" alt="2018 Results" src="https://user-images.githubusercontent.com/80016496/112729208-0792e380-8ef9-11eb-94bf-97b52c4d8da8.png">

### VBA Code-Related Results
In order to conduct this analysis more efficiently, we undertook a refactoring exercise to streamline the code, and run our analysis more expiditiously.

### Advantages and Disadvantages of Refactoring Code

### Discrete Refactoring Benefits of This Analysis

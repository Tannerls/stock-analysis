
> # An Analysis of 2017 and 2018 Stocks
> 
## Overview of Project
Steve is advising his parents on stock options and wants to make sure they are going to get a good return on their investments. An analysis was performed on stock data from 2017 through 2018 to help Steve advise his parents by looking at the daily volume and the yearly return from the twelve different stocks options. In order to assist Steve, VBA language needed to be learned and applied. I then refactored the code in order to speed up the run time and clean up the code.

## Results
Comparing the results from 2017 and 2018 there are two stocks, ENPH and RUN, that remain positive on return going in to 2018.


![2017 Results](https://user-images.githubusercontent.com/83738699/124402381-9633af80-dcf5-11eb-8e69-850691073459.PNG)![2018 Results](https://user-images.githubusercontent.com/83738699/124402388-9fbd1780-dcf5-11eb-9a68-cc8c8ca6b217.PNG)

When you look at 2017 Run only had 5.5% on return with 267,681,300 in daily trade volume, while ENPH had a very high return on 129.5% and only had 221,772,100 in daily trade volume. Based on 2018 data, ENPH would be the best option to invest in and get the most in return.
#### Execution times for VBA script
The original script was refactored to make it run faster when pulling different years. The original script ran 2017 code in .836 seconds, while the refactored code ran in .211 seconds.

![2017 Module code](https://user-images.githubusercontent.com/83738699/124402859-879ac780-dcf8-11eb-8477-9afd185c2da1.PNG)  ![2017 refactored](https://user-images.githubusercontent.com/83738699/124402883-9ed9b500-dcf8-11eb-8cdf-c1cc57abac7e.PNG)

The same was true for 2018 results going from a run time of .843 to .211.


![2018 Module code](https://user-images.githubusercontent.com/83738699/124402909-c6c91880-dcf8-11eb-8a34-cce6bb6baba4.PNG)  ![2018 refactored](https://user-images.githubusercontent.com/83738699/124402915-cd579000-dcf8-11eb-9692-c8640372f2b4.PNG)

By speeding up the run time for the script more data could be analyzed faster on future projects with more data.  
## Summary
Some advantages to refactoring code is that you already know the code works and how it works. So by refactoring it you are able to improve the efficiency of the output along with adjusting any visual aspects displayed that you didn't like on the first code. By improving the structure of the code and how it runs, you are able to decrease the run time and possibly adapt it to include future data not yet available. There are several benefits to refactoring code, but some disadvantages is that you could possibly break an already functioning code and waste time trying. 
#### Effects of refactoring code for this analysis
When I was refactoring the code for this analysis I ended up spending a lot of time figuring out how to use the tickerIndex in the If statements as well as increasing the tickerVolume. There was a lot of trial and error as well as broken code. Some advantage's to refactoring the code was that I could eliminate the formatting code I used in the original script to create grid lines for the checkers board. Also the script runs faster and looks a lot cleaner than the original, I was also able to adjust a variable from using j back to i since i understood how its used better.

> Written with [StackEdit](https://stackedit.io/).

# Module 2 Challenge: VBA Stock Analysis
---
## Overview 
The purpose of this challenge was to refactor code that was previously used during the lessons during the module to see if there was any main differences between the code we used before versus the refactored code. Doing this we were able to see trends from the data collected for both the stocks in the years 2017 and 2018. Using the refactored and orginal code both yielded the same results.
---
### Analysis 
The overall analysis during this challenge was to see and compare the refactored code with the original code that we used. Some key takeaways from the analysis itself was the difference in percent changes between the two years of each stock. We saw that the year 2017 was a good year for the majority of the stocks and 2018 was a bad year. We were able to highlight this using formatting tools, showing red for negative percentage changes and green for positive percentage changes. Putting together the refactored code we saw that in many cases it was very similar to the original but instead all of the tools we wanted to use were summarized in one long script making the insert button run all the macros we wanted to previously in one push. Our For loops and If statements to return the data from each year is shown below. 
![Code](https://github.com/mckjack/stock-analysis/blob/main/Code.png)
--- 
We see that both are very similar but in this case we use a tickersIndex for i = 0 to 11 which then we were able to set our three output arrays and store them as data types. 
---
Running this macro were able to see that the data given to us was the same as it was in the module for both 2017 and 2018 data. 
![VBA_Challenge_2017](https://github.com/mckjack/stock-analysis/blob/main/VBA_Challenge_2017.png)
---
![VBA_Challenge_2018](https://github.com/mckjack/stock-analysis/blob/main/VBA_Challenge_2018.png)
> And here are the pictures of the script timed for each respective year
![Timed](https://github.com/mckjack/stock-analysis/blob/main/VBA_Challenge_2017_timed.png)
---
![Timed_2](https://github.com/mckjack/stock-analysis/blob/main/VBA_Challenge_2018_timed.png)

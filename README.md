# Module 2 Challenge: VBA Stock Analysis
---
## Overview 
---
#### The purpose of this challenge was to refactor code that was previously used during the lessons during the module to see if there was any main differences between the code we used before versus the refactored code. Doing this we were able to see trends from the data collected for both the stocks in the years 2017 and 2018. Using the refactored and orginal code both yielded the same results.

## Analysis 
---
#### The overall analysis during this challenge was to see and compare the refactored code with the original code that we used. Some key takeaways from the analysis itself was the difference in percent changes between the two years of each stock. We saw that the year 2017 was a good year for the majority of the stocks and 2018 was a bad year. We were able to highlight this using formatting tools, showing red for negative percentage changes and green for positive percentage changes. Putting together the refactored code we saw that in many cases it was very similar to the original but instead all of the tools we wanted to use were summarized in one long script making the insert button run all the macros we wanted to previously in one push. Our For loops and If statements to return the data from each year is shown below. 
---
![Code](https://github.com/mckjack/stock-analysis/blob/main/Code.png)
--- 
> We see that both are very similar but in this case we use a tickersIndex for i = 0 to 11 which then we were able to set our three output arrays and store them as data types. 
---
> Running this macro were able to see that the data given to us was the same as it was in the module for both 2017 and 2018 data. 
---
![VBA_Challenge_2017](https://github.com/mckjack/stock-analysis/blob/main/VBA_Challenge_2017.png)
---
![VBA_Challenge_2018](https://github.com/mckjack/stock-analysis/blob/main/VBA_Challenge_2018.png)
> And here are the pictures of the script timed for each respective year
---
![Timed](https://github.com/mckjack/stock-analysis/blob/main/VBA_Challenge_2017_timed.png)
---
![Timed_2](https://github.com/mckjack/stock-analysis/blob/main/VBA_Challenge_2018_timed.png)
## Summary
---
#### To summarize our findings we can look at the advantages and disadvantages of refactored code used in this challenge and the idea of refactored code as a whole. Some advantages of refactored code as a whole are
 - deals with bug fixes very well meaning that the refactored code will be free of most bugs than the original.
 - It is subject to peer review meaning that others have looked over the code to ensure that it is right in every line and can run smoothly.
 - eliminates duplicate code meaning that it is a way more efficient way to run tasks since it avoids looking at certain scripts twice. 
 ---
 Some disadvantages of refactored code as a whole are
 - introducing new bugs that could make it difficult to run
 - very time consuming, instead of writing original, could take much longer to formulate a whole code to run tasks.
---
We saw some these advantages and dissadvantages during this challenge to give us a better viewpoint on refactored code. We saw that it made the task of getting the data and formatting it into an excel table much quicker since it took all of the original and put it into one sub routine. This meant that we could run the code much faster due to the button being inserted was assigned to the one macro. However while writing the new script, we ran into some isssues involved the variables since we had to rename them from the original and keep that consistent throughout the script. 

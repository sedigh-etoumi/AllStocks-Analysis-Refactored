# AllStocks-Analysis-Refactored
Develop enhanced code for stocks analysis
Overview.
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.
In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.
Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.
So, we successfully developed new code to reduce the run time for significantly better performance as attached below for both years 2017 and 2018.







 
	Stocks Analysis 2017


 
Stocks Analysis 2017
As we’ve seen in screenshots above 2017 was better than 2018 in general for investments in stocks, just one stock “TERP” was red it’s mean down in 2018.
And the running time was reduced because instead of doing nested loop, we used arrays and tickerIndex and the ticker were working as first loop to held the stock name, which is significantly reduced the calculations and operations about 3 times better this enhanced code it will help specially when we have thousands of stocks.

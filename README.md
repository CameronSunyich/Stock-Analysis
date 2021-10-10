# Stock-Analysis

## Overview of Project
  Our client Steve loved the Green_Stocks_Analysis Workbook so much he asked us to refactor the code to find more stock information for other stocks in the 2017 and 2018 for potential investment. The improved workbook will now include twelve potential stocks, thier ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The code created in this Workbook then retrieves this information, finds the total daily volume, and the return on each stock.
  By refactoring this code, my goal is to make it more effieciant and be able to retrieve data for multiple tickers. This will then give Steve the needed information to help his parents invest in more beneficial stocks. This refactored code will also have the potential to add more stock to the dataset for future stock investments.

## Results and Analysis
  When refactoring the code, I followed the instructions given.    
https://2u-data-curriculum-team.s3.amazonaws.com/dataviz-online/module_2/challenge_starter_code.vbs

   After following these instructions in VBA I was able to refactor the code to analize mulitple stocks much more efficiantly. When running stocks for 2017, the code finished in 0.19531125 seconds. The original code ran at a speed of 0.9765625 seconds. The refactored code was just as efficiant when running stocks for 2018 at 0.1953125 and the original running at, 0.9765625 seconds. Refer to image below for pop-ups showing the refactor code run time.
   
![new code](https://user-images.githubusercontent.com/90425412/136683372-5dcddf5f-6df0-4d31-aa60-538a4ca6659b.png)
![new code 2018](https://user-images.githubusercontent.com/90425412/136683593-d036d7cb-6795-489a-9242-e1e6bb876d83.png)

   As it can be seen the refactored code runs much more quickly than before. This can be beneficial if Steve decides to include more stocks to the data and more easiely edited with the code being more organized.

  ## Summary
  Overall, refactored code helps organize the code, to be easier to fix bugs, and understand the code and its patterns better. However, when refactoring any code this affects case testing parts of the code and is very time consuming to duplicate code for multiple function use.
  After completeing the Workbook and refactoring the code I came to find that while the new code is running more efficiantly and is now easier to read. The process can still be very time consuming, when creating this kind of code. Meaning refactoring can limit our size of data and testing existing code in the function. By refactoring we now will have an easier time future debugging coding and be able to identify errors in the code itself. This could limit Steve on his ability to recreate this code and add more stocks to the dataset. I'd recommend moving to a different software this more case tester friendly and can hadle larger set of data like Python. If Steve does wish to add more stocks to the dataset.

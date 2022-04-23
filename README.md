# Stocks Analysis with VBA

## Overview of the Project

Steve, a friend, is helping his parents make an informed decision while investing in stocks of companies that work with alternative energy. Steve's parents have decided on a stock already (Ticker:*DQ*), however Steve wants to analyze the **Total Daily Volume** and the **Returns** of eleven additional green energy companies in order guide them armed with more information. In order to do this, Steve has requested help to analyze stock returns of green energy companies from 2017 and 2018.

### Purpose

The objective of this project is to use *Microsoft* Visual Basic for Applications or **VBA** to analyze stock returns for 2017 and 2018. However, we want to present a more efficient VBA script that runs through the analysis quicker. To accomplish this, a previously written script will be refactored. Based on the data and the conditional formatting of the sheet to aid visualization, Steve can help his parents make a well-informed decision on which stocks to invest in.  

## Results

Below is the analysis of twelve stocks including *DQ* for the year 2017 and 2018. The output observed does not change after refactoring the code. There are 3 components addressed in the stocks analysis data:
- the ticker name of the stock
- the total daily volume of the stock
- the percentage of yearly return (where all stocks that went up in value are highlighted in green, while the ones that went down are highlighted in red)

The outputs are as follows:

![Output-2017](https://user-images.githubusercontent.com/102441140/164943539-a166d46f-854e-4543-b087-5db209f794d1.png) ![Output-2018](https://user-images.githubusercontent.com/102441140/164943624-79349de6-f1c8-44f9-8bec-13c4538c4921.png)

### Total Daily Volume
High daily volume is generally a good sign for a stock since it shows that the stock has a lot of interest and activity, and is generally stable. But nonetheless, a low daily volume doesnt mean that the stock isn't a good stock. Tt could simply mean that the stock might be "undiscovered". For *DQ* the daily volume is pretty low when compared with the other 11 stocks in 2017. *DQ* had the lowest daily volume in 2017. But the daily volume tripled in 2018. Yet it remained as one of the stocks with the lowest daily volume is 2018.

### Percentage of Returns

From the visual conditional formatting, we can concude that overall 2017 was a good year for green energy companies. All companies, excluding *TERP* had positive returns. *DQ* while having a low daily volume had an increase of almost 200% in 2017. Thus, in 2017 at least, *DQ* was a robust stock to invest in. But the 2018 analysis reveals a grim picture. All companies except *ENPH* and *RUN* showed negative returns. *DQ* not only had the lowest daily volume but also lost 63% of its value in 2018. 

Based on this analysis, Steve's parents should be cautious when thinking about investments in the green energy stocks. Since *ENPH* and *RUN* were the only two companies in our analysis that continued to have positive returns, Steve's parents should think about maybe looking at those stocks instead of *DQ* for investment. Moreover, Steve should also potentially analyze other stocks in the dataset over a longer period of time (2017-2021) to have a clearer picture.

### Code used for output 

The table below comapres the code before and after refactoring 

Code before refactoring. |  Code after refactoring.
:------------------------------------------:| :-------------------------------------:
Original code without arrays (click to enlarge).  | Refactored code with arrays (click to enlarge).	
![original](https://user-images.githubusercontent.com/102441140/164944810-68ae6bd1-c36d-4bf2-bd47-f95d78b48c3a.png)  | ![refactored](https://user-images.githubusercontent.com/102441140/164944651-0e6691d8-e5e2-47e2-8a2e-5eebc9beadea.png)
The code is in a nested loop and is going through one ticker at a time to generate the output in the selected worksheet | Code stays in the same loop, gathers all data and stores it in arrays first. In a separate for loop the output is pulled from the arrays and then populated in the selected worksheet.  
 Execution time of the code for 2017: |  Execution time of the code for 2017:
<img width="649" alt="green_stocks_2017" src="https://user-images.githubusercontent.com/102441140/164944905-fb8e223d-32f5-47d6-b667-b788ea5ef344.png"> | <img width="649" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/102441140/164944880-9be2f3c0-7ac7-4747-9e08-3054cb1e1fbc.png">
Execution time of the code for 2018: |  Execution time of the code for 2018:
<img width="649" alt="green_stocks_2018" src="https://user-images.githubusercontent.com/102441140/164944937-98ff0ad5-dcaf-465b-9119-4d305c814d30.png"> | <img width="649" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/102441140/164944942-ac47eed1-b80f-4258-995e-99cfe3afb3db.png">
Code runs in 0.25 seconds | Code runs in 0.05 seconds ( **Almost 5 times faster** )

## Summary

**What are the advantages or disadvantages of refactoring code?**

- *Advantages of refactoring code:*
  - **efficiency:** One of the main reasons why we need to refactor code is to execute it more efficiently so that it provides the output quicker. For example, in our refactoring we stored the data in arrays which is more efficient instead of reassigning variables after each loop
  - **readability:** Often refactoring usually makes the code more *readable* as the steps are easier or more logical.
  - **improved functionality:** going over the same code a second time around also helps catch logical falacies that can lead to bugs, such as failing to account for any conditions.
  - **universal:** Often it is easier to start coding based on one specific condition or dataset, and later refactoring the code so it can be applied to multiple different conditions or other datasets.
  - **reuse others' code:** Often it is easier to refactor code from trusted websites that others have written instead of working to write something from scratch.
- *Disadvantages of refactoring code:*
  - **frustrating** depending on your familiarity with the original code, it ight be more time-consuming to refactor the code instead of just writing your own code from scratch. For example, if you are trying to refactor code of the internet for your own purposes, if you are unfamiliar with the concepts, you might not understand the code well enough to refactor it properly to suit your needs.
  - **less efficient:** unfortunately if one is unfamiliar with the code, refactoring might even make the code hard to read, inefficient or prone to bugs.

**How do these pros and cons apply to refactoring the original VBA script?**

Refactoring the original code was a good decision since it made the code run almost 5x times faster. A key component for success is meticulous commenting on the original code so that the future coder can understand the logic behind the original code. Editing the code to store the values from the for loop in arrays was a more efficient use of the computer memory which helped the code execute quicker. The addition of the arrays also made the code more universal since it will be easier to refactor this code to not only run on our datasets but others with some minor adjustment, whereas using the original code, it would be more clunky. 

As a beginner, refactoring did little to improve readability on my part since I had trouble understanding why the loop needed to use the arrays. And since I was reusing old code, I had to be careful with all the variables being used and a good portion of my time was spent catching variable mismatch errors which was frustrating. I also had to add more components to the original code which made it longer and therefore was more time-consuming. 

But if there is time and a good understanding of the code you are working with, refactoring the code can definitely improve the quality of the code and lead to a more efficient execution which saves more energy in the long run.

# stock-analysis
Performing automated analysis to stock data using VBA

  ## Overview of Project: Explain the purpose of this analysis.
  
In this module, we helped our good friend Steve, who's just graduated with his finance degree. His parents couldn't be prouder. They're so proud in fact, that they're going to be his first clients. 
Because Steve's parents are passionate about green energy, they have decided to invest all their money into DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. But Steve thinks that his parents' funds should be more diversified, so he wants to analyze several green energy stocks, in addition to DAQO stock.
Steve has given us an Excel file containing the stock data he wants you to analyze. We'll be using an extension to Excel, built to automate tasks: Visual Basic for Applications, usually referred to as VBA.

## Results: Using images and examples of your code, compare the stock performance between 2018 and 2018, as well as the execution times of the original script and the refactored script.

Using the information learnt about marcos in this module, i was able to create ticker indexes for the stocks we are going to analyze and use for loops to run through all of the rows to find information about the tickers in regards to profitability to enable our friend Steve determine which of the tickers were worth investing in for his parents in the years 2017 and 2018 as shown below.

<img width="290" alt="VBA_Challenge_2018 results" src="https://user-images.githubusercontent.com/47859209/189032904-2b824a1d-1d17-49ec-86ef-cc021ff24788.png">

<img width="860" alt="Screen Shot 2022-09-08 at 12 05 57 AM" src="https://user-images.githubusercontent.com/47859209/189033486-ae340316-706c-473a-b24c-42a30ea86065.png">

<img width="857" alt="Screen Shot 2022-09-08 at 12 06 48 AM" src="https://user-images.githubusercontent.com/47859209/189033508-100f61bd-7db5-48f2-b4b6-11038af402f4.png">

<img width="291" alt="VBA_Challenge_2017 results" src="https://user-images.githubusercontent.com/47859209/189032342-88723323-0022-4fe8-9129-6af3075b75a4.png">

An analysis using If and Then statement in conjunction with nested loops was performed using our macros to display information on the elapsed run time of our the stock data over the 2017 and 2018 years.

<img width="266" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/47859209/189033347-2ba4a695-a28f-4442-9904-6993c5a4d42b.png">

<img width="257" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/47859209/189033398-a6b1828c-1bf5-4f22-97cc-d0fcbd523b0f.png">



# Summary: In a summary statement, address the following questions.

## What are the advantages or disadvantages of refactoring code?

Advantages:
- Better comprehensibility facilitates maintenance and the extendibility of the software
- Code refactoring reduces the likelihood of errors in the future and simplifies the implementation of software functionality.
- Another potential purpose of code refactoring is performance improvement. So refactoring may enable an application to perform faster or use fewer server capacities. This is the benefit that might be really tangible for end users right after code refactoring.

Disadvantages: 
- Shoddy refactoring could introduce new errors and bugs into the code that were not previously there
- Everybody's code is different and no clear precise definition of concise code
- A disadvantage to refactoring in cohorts is that it could take a long amount of time as well.


## How do these pros and cons apply to refactoring the original VBA script?
- We ended up making our worksheet accessible and with lots more information to provide, however, we had to write more code in order to do that.

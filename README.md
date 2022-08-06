# An Analysis of Stock Data Using Excel

## Overview of Project
A VBA macro was created to enable financial analysis of stock performance.  This macro can be executed at the click of a button to analyze a desired year of data.
Because there could be a need to analyze a larger data set, this project was initiated to refactor and measure the performance of the macro.


# Results

## Performance of the Original Macro
The original financial analysis macro included a timer to enable performance measurement.  There were two datasets, one for 2017 and one for 2018.  The macro was used to analyze each dataset and the elapsed time was presented in a message box to the user.
Here is the result recorded for the 2017 dataset:

![VBA_Challenge_2017-Original Message Box](Resources/VBA_Challenge_2017-Original.png)

Here is the result recorded for the 2018 dataset:

![VBA_Challenge_2018-Original Message Box](Resources/VBA_Challenge_2018-Original.png)

## Augmentation of the Original Macro
Because the goal was to simplify the data analysis, the refactored macro also included conditional formatting of the financial data.  The original macro did not include this feature, so to ensure the comparison included all the desired functionality, the original macro was augmented with the reporting code and the timing measures were recorded.

Here is the result recorded for the 2017 dataset:

![VBA_Challenge_2017-Original_with_formatting Message Box](Resources/VBA_Challenge_2017-Original_with_formatting.png)

Here is the result recorded for the 2018 dataset:

![VBA_Challenge_2018-Original_with_formatting Message Box](Resources/VBA_Challenge_2018-Original_with_formatting.png)

## Refactoring the VBA Macro
The VBA macro was refactored to reduce the amount of rows processed per ticker symbol by creating three arrays to contain the data collected during the examination of the dataset rows, as seen in the following code snippet:   
```
    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
```

Additionally, a check was performed to determine if the current row contained the final entry for the current ticker symbol.  If this was the final data row for that ticker symbol, an `Exit Next` command was used to break out of the row processing and move on to the next ticker symbol.

Again, the macro was executed for both data sets and the timing measures were recorded.

Here is the result recorded for the 2017 dataset:

![VBA_Challenge_2017 Message Box](Resources/VBA_Challenge_2017.png)

Here is the result recorded for the 2018 dataset:

![VBA_Challenge_2018 Message Box](Resources/VBA_Challenge_2018.png)

## Analysis of VBA Macro Timing Measurements
The data collected during the execution of the macro runs was entered into Excel to compute the changes in performance and to visualize the results. This data is shown in the following table:

![VBA_Challenge-Performance_Data in Excel](Resources/VBA_Challenge-Performance_Data.png)

The Message Box results were used to compute the additional time that elapsed with the addition of the result formatting being added to the original macro.  We can see that these operations extended the original macro execution time by approximately 4.28% for the 2017 data set, and approximately 2.13% for the 2018 data set.

By examining the timing measures, we can see that the refactored macro performs significantly better.  The 2017 data set was processed in only 49.23% of the time that was required by the original macro (with results formatting).  That is an improvement of 50.77%.

Similarly, we can see that the 2018 data set was processed in only 48.96% of the time that was required by the original macro (with results formatting).  That is an improvement of 51.04%.

These results are clearly visible in the following chart:

![VBA_Challenge-Performance_Chart in Excel](Resources/VBA_Challenge-Performance_Chart.png)

# Summary
In general, there are advantages and disadvantages to refactoring code.

## General Advantages to Refactoring Code
Refactoring code can result in improvements in readability, maintainability, and performance.

## General Disadvantages to Refactoring Code
In order to perform refactoring, the developer needs to be able to understand the existing code.  This means that the practices of commenting and formatting the original code need to be consistently applied, because variation can lead to confusion and may extend the time required to perform any refactoring.  

In order to determine if performance has been improved, measurements will be required as the code is repeatibly executed.  This also takes time and effort, and can increase costs associated with the refactoring effort.

## Outcome of Refactoring the VBA Macro
For this analysis, the outcome was strongly positive because the overall performance of the macro execution was greatly increased.  The macro code was restructured which might be considered more complex (as a slight disadvantage), but the commenting and formatting both work to ensure that the code remains readable.

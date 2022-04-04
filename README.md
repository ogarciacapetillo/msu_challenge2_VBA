# VBA Challenge - Module 2

## Overview 

The analysis presented on the attached VBA_Challenge Workbook has a goal to show the performance of the stocks for the available years:
- Year of 2017
- Year of 2018

The yearly return is being used to represent if there is a gain or loss for the investment performed for year selection. 

## Results

**Stock performance** for the years under analisis are represented in the following table

| Year: 2017 | Year: 2018 |
| --- | --- |
| ![Yearly Return 2017](/Resources/Yearly_return_2017.png) | ![Yearly Return 2018](/Resources/Yearly_return_2018.png) |

It can be observed that DQ earnigns where significantly better during the 2017 year where the majority of the tickers grew, where for the year 2018 a shrink behavior can be observed. With the exception of two stocks
- ENPH
- RUN

**Macro performance** the following illustrations exhibit the performance of the refactored code versus the original version of the code

| Year | Original version | Refactored Version |
| --- | --- | --- |
| 2017 | ![2017_original_performance](/Resources/performance/2017_original_version.png) | ![2017_refactored_version](/Resources/performance/2017_refactored_version.png)|
| 2018 | ![2018_original_performance](/Resources/performance/2018_original_version.png) | ![2018_refactored_version](/Resources/performance/2018_refactored_version.png)|

A considerable improvement of the performance can be observed, due the reduction of the loop processing in the refactored code.

## Summary

As part of the summary the following questions are being answered:

**1. What are the advantages and disadvantages of the refactored code?** - It can be observed that refactored code performance is considerable faster than original version. Original code nested loop at tickers level was generating to loop all the data rows N times as per Tickers available which generated a huge difference in the performance to get processed. Once the nested loop was removed and data was used on a singular loop swipe the processing times are considerably better.  
**2. How do these pros and cons apply to refactoring the original VBA script?** - Original VBA script was not effectively collecting the required data on single loop due the lack of arrays to properly separate the data by Tickers which have the need to loop by the Tickers as the parent loop and causing to read the data more times that need it, causing an increased usage of processing resources. 
While adding the arrays to manage collected values accordingly, it was possible to perform a single loop and no longer required a nested loop to correctly capture and process the data as required. 

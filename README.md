# VBA_Challenge
### Refactor VBA code and measure performance

## Overview

The purpose of this Excel VBA code exercise is to refactor a given code solution to increase data process efficiency when handling larger amounts of investment stock data,  comparing the performance for both, original and refactored  solutions.

## Results

The original code was modified to loop through all the data one time, using an index variable to process array values as shown in this [side by side comparison](https://github.com/serpaulus/VBA_Challenge/blob/main/Side_by_Side_Code_Comparison.pdf).

After refactoring the code we ran at significantly different execution times for avaliable data sheets [2017](https://github.com/serpaulus/VBA_Challenge/blob/main/VBA_Challenge_2017.PNG) and [2018](https://github.com/serpaulus/VBA_Challenge/blob/main/VBA_Challenge_2018.PNG).

Comparing the time results, we can clearly see that refactoring the code significantly the reduces data processing time by over half in both sheets.   

## Summary

This exercise clearly shows how proper code structure and function utilization can significantly increase application performance. 

The original code, although easier to understand and write, more than doubles the processing times. The refactoring process although challenging will likely show it’s worth, especially when processing larger amounts of data.    

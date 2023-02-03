# VBA Scripting Homework: Stock Data Assignment

The purpose of this assignment is to analyze stock market data by using VBA script and using command function such as For loop, If, If Else if, and Max to sort out data and analyze it.

In this assignment, we use 2 Excel workbooks; the first one is the real stock data workbook named 'Multiple_year_stock_data.xlsx', where we have stock data for 3 years starting from 2018 to 2020. Another workbook is called 'alphabetical_testing.xlsm', which is the test workbook where we have smaller amount of data in one sheet. 

I started working on the test workbook to test our script since the run time is much faster than working on the real data workbook and once I finish writing the script and everything looks good, I copied the script to the real homework workbook to complete the assignment.

Solution
First, I gather the ticker symbol and sum up the volume for each ticker by using a For loop.  For each ticker found, I summed up the volume until I reach the last value for that ticker.  I then proceed to do the same thing with the next ticker symbol until I ran out of ticker symbol.

Second, I used the For loop to find the open value and close value for each ticker in order to calculate the yearly change and the percent change. When calculating percent change, I made sure that I did not divide by zero. If the open value is zero, I grab the next value instead.

Third, I used WorksheetFunction.Max and WorksheetFunction.Min to calculate the greatest % increase, the greatest % decrease and the greatest total volume. I also used the WorksheetFunction.Match to find the ticker that match the greatest increase, greatest decrease and greatest total volume.

Forth, I used Interior.colorIndex to format the cells in column J to display red for negative values and green for positive values. I also used percentage style to format column K to display number as percentage.

Last, after everything is done, I used For Each ws InThisWorkbook.Sheets to enable the VBS script to run every worksheet at once.

**Note: The following folders contain:**  

Image_Result: Screenshots of result or each year  
Stock Data_2018-2020: VBS scripts and Excel Workbook

© 2023_Ratima Chowadee 

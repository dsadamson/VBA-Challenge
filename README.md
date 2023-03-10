USING VBA AND EXCEL TO TRACK STOCK ACTIVITY OVER THREE YEARS

Description:

In this project, I was given an Excel workbook with three worksheets that listed a variety of stocks, their values, and their volume over three years, between 2018 and 2020. To make this data more easily digestible, I was instructed to create two summary tables for each spreadsheet of data pertaining to a given year: one to track each stock's yearly change, percent change, and volume, and another to track which stocks had the greatest percent increase in value, greatest percent decrease in volume, and greatest total volume in a given year. 

To create these tables, I converted the Excel workbook into a macro-enabled (.xlsm) file and wrote a VBA script to loop through each of the spreadsheets in the workbook. This script was programmed to create both summary tables at the beginning, where it would store values as it looped through each year's stock data. Then I instructed the VBA script to find the first opening value, which would have otherwise been excluded over later in the loop through each spreadsheet. Next, I programmed an 'if' statement, which told VBA and Excel to track when it was searching through the same stock ticker. Within this statement I told the program to add each volume value together when the stock ticker values were the same.

A majority of what the VBA script tracks is contained in the 'Else' statement, though. First, once the ticker changes, the previous ticker is recorded in the large summary table. Then, the program tracks the last closing value for that ticker, and the opening value is subtracted from the closing value to find the yearly change. This yearly change value is finally divided by the opening value to find the percent change. These values are then reset as the program tracks the next stock, and the yearly change and percent change are recorded in the summary table. As the program continues, it also places the sum of each respective stock's volume in the first summary table; fills the yearly change cells with red if a stock had a negative yearly change or green if a stock had a positive yearly change; and places the stocks with the greatest percent increase, greatest percent decrease, and greatest volume in the second summary table. 

In order to find the greatest percent increase, greatest percent decrease, and greatest volume, I ran a Worksheet.Function in VBA to find the maximum and minimum values in the percent change column, then the maximum value in the volume column. However, this meant I was searching for a particular value, which would not record the stock it was connected to. To get around this, I ran a Worksheet.Function to match the values in the second summary table to the cells in the first summary table where they were found. However, since the match function shows the number of rows before a value appears, I had to format my process to retrieve the stock ticker as ws.Cells(ws.Range("P2").Value + 1, 9), for example, so the program would retrieve the ticker value from the row it is actually in, rather than the row before it. If I were to code this task in the program again, I would try to find a more efficient way to do this, with 'if' statements rather than functions, as the functions seem to be inefficient even if they retrieve the correct value.

Installation:

Import the stock_data.vbs script into the Virtual Basic editor in Excel (or copy and paste it from Stock_data_vba_script.txt into the editor), then save it as a macro to the Multiple_year_stock_data.xlsm workbook. Run the code by pressing the play button at the top of the Virtual Basic window. Double-check whether the code ran correctly by comparing the output in your Excel worksheets to the screenshots included in the Week_2_VBA_Challenge folder.

Authors:
Daniel Adamson

Acknowledgement:
Thank you to James Miller, Danny Furman, Dean Taylor, and the Data Analytics Bootcamp study groups for providing tips and helping to debug issues with this program.

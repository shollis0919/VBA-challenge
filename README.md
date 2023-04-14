# VBA-challenge
This program VBA is used to go through stock information in 1 year and produce and display info on: price changes, percent of change on price, total stock volume, the greatest inccrease and decrease in price, greatest amount of stock volume

The program also works through each worksheet in an open spreadsheet

The program can be found in stockExchangeVBA.bas
There are also screenshots to show the expected results of the program in Excel based on the spreadsheet given for the assignment.

The following codes were produced based on info from:
https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat#syntax
https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-us&rs=en-us&ad=us


ws.Range("J" & sumTableRow).NumberFormat = "###0.##"
ws.Range("K" & sumTableRow).NumberFormat = ".##%
ws.Cells(2, 17).NumberFormat = ".##%"
ws.Cells(3, 17).NumberFormat = ".##%"
 

#psReportHelper

This is a PowerShell script module for creating an HTML report with CSS. The idea is to build on this, so that multiple PowerShell scripts can use this common repository for creating elaborate HTML reports that contains Tables with uniform colors and other details.

You can modify the CSS for a common look and feel.

You can take a look at Sample\Case1\case1.ps1 to learn how to use the module quickly

Get-Help -full for the command below for detailed help

##### Set-ReportFile - Set the file where you want to emit a table
##### Add-TableStart - Start creating the table
##### Set-ColumnWidth - Set width of columns in advance. For ex. 300, 200, 200, 300 will have 4 columns.
##### Add-TableStartRow - Start creating the row
##### Add-TableCells - Add cells in the row
##### Add-TableEndRow - End the Row
##### Add-TableEnd - End the table
##### Get-Table - View the table in host or browser
##### Add-HeadingText1 - Add the largest heading
##### Add-HeadingText2 - Add the 2nd largest heading
##### Add-HeadingText3 - Add the 3rd largest heading
##### Add-HeadingText4 - Add the 4th largest heading
##### Add-HeadingText5 - Add the 5th largest heading
##### Add-HeadingText6 - Add the 6th largest heading
##### Add-Text - Add a text with appropriate font setting and alignment
##### Add-TimeStamp - Add timestamps in multiple formats

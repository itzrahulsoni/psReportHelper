Import-Module ..\..\Common\helperReport.psm1 -Force

Set-ReportFile -Folder "D:\Documents\GitHub\psReportHelper\Report Helper\Output" -File "output.htm" -OverWrite $true
Add-TableStart -Width 700 -CellPadding 3
Set-TableColumnWidth 200, 200, 300
Add-TableStartRow
Add-TableCells -data "Name", "Address", "WebSite" -isHeader $true
Add-TableEndRow
Add-TableStartRow
Add-TableCells -data "Rahul" -isGreen $true -align "left"
Add-TableCells -data "Soni" -isRed $true -align "center"
Add-TableCells -data "Test 1" -isGreen $true -align "right"
Add-TableEndRow
Add-TableEnd
Add-HeadingText1 -Message "Hello World - 1"
Add-HeadingText2 -Message "Hello World - 2"
Add-HeadingText3 -Message "Hello World - 3"
Add-HeadingText4 -Message "Hello World - 4"
Add-HeadingText5 -Message "Hello World - 5"
Add-HeadingText6 -Message "Hello World - 6"
Add-Text -Message "Hello World" -align "center"
Add-TimeStamp -Format 9
Get-Table -OpenInBrowser $true -InsertCSS "D:\Documents\GitHub\psReportHelper\Report Helper\Common\style.css"
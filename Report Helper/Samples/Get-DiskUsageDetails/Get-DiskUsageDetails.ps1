$common = "..\..\Common"
Import-Module "$Common\helperReport.psm1" -Force
Set-ReportFile -Folder "." -File ".\Output.htm" -OverWrite $true   #Set a report file

Add-TableStart -Width 280 -BorderColor Brown -BorderWidth 3        #Start the Table
Set-TableColumnWidth 80, 100, 100                                  #Width of individual columns

Add-TableStartRow
Add-TableCells "Drive", "Free Space", "Total Space" -IsHeader $true
Add-TableEndRow

foreach($line in (Get-WmiObject win32_logicaldisk))
{
    #Adding one row of data each time
    Add-TableStartRow
    
    #Adding each cell individually
    Add-TableCells -Data $line.DeviceID

    #If the space available is less than 10% of the disk size, alert
    if($line.FreeSpace -le ($line.Size * 0.1))
    {
        Add-TableCells -Data ($line.FreeSpace / 1GB).ToString("0 GB"), ($line.Size / 1GB).ToString("0 GB") -IsRed $true -Align "right" -BorderColor Red
    }
    else
    {
        Add-TableCells -Data ($line.FreeSpace / 1GB).ToString("0 GB"), ($line.Size / 1GB).ToString("0 GB") -IsGreen $true -Align "right" -BorderColor Green
    }
    Add-TableEndRow
}
Add-TableEnd                                                       #End the Table
Get-Table -OpenInBrowser $true -InsertCSS "$common\style.css"
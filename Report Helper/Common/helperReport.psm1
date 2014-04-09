$script:file

<#
.Synopsis
    Set the Report folder and file location
.DESCRIPTION
    In order to use this module, you must set a Report folder and file name. The file name should end with htm or html. Once done, you will be able to use all the other functions.

    This is the detailed flow for creation of a table using this module.
    - 1. Start with Set-ReportFile and set the folder location along with an output file name
    - 2. Start Table creation process using Add-TableStart
    - 3. Set the width of columns here if required using Set-TableColumnWidth
    - 4. Start a Row using Add-TableStartRow
    - 5. Add Cells using Add-TableCells
    - 6. End a Row using Add-TableEndRow
    - 7. Repeat steps 4, 5, 6 until you are done adding all rows.
    - 8. End Table using Add-TableEnd

    Sample script that creates a table with 2 rows.
    
    Set-ReportFile -Folder "C:\Temp" -File "Sample.htm" -OverWrite $true
    Add-TableStart -Width 700 -CellPadding 3
    Set-TableColumnWidth 300, 300, 100
    Add-TableStartRow
    Add-TableCells -data "Name", "Address", "WebSite" -isHeader $true
    Add-TableEndRow
    Add-TableStartRow
    Add-TableCells -data "Rahul", "Soni", "www.dotnetscraps.com"
    Add-TableEndRow
    Add-TableEnd
    Get-Table -OpenInBrowser $true -InsertCSS "C:\Temp\style.css"
.EXAMPLE
    Set-ReportFile -Folder "C:\Temp" -File "output.html" -OverWrite $true
.EXAMPLE
    Set-ReportFile -Folder "C:\Temp" -File "output.html" -OverWrite $false
#>
function Set-ReportFile()
{
    param(
        [Parameter(Mandatory=$true, HelpMessage='Provide a folder name... like, C:\Temp')]
        [string]$Folder=$null, 
        [Parameter(Mandatory=$true, HelpMessage='Provide a file name ending with htm or html like, output.html')]
        [string]$File=$null,
        [Parameter(Mandatory=$true, HelpMessage='Do you want to overwrite [$true or $false]?')] 
        [boolean]$OverWrite=$true
    )
    if((Test-Path($Folder)) -eq $false)
    {
        Write-Host -ForegroundColor Red "Sorry, this folder doesn't exist!"
        return
    }
    if((Test-Path("$Folder\$File")) -eq $true)
    {
        if($OverWrite -eq $true)
        {
            Write-Host -ForegroundColor Red "This file already existed. Deleting it!"
            Remove-Item "$Folder\$File"
        }
        else
        {
            #Can't overwrite
            Write-Host -ForegroundColor Red "The file already exists in the current folder. Try -OverWrite $true"
            return            
        }
    }
    $script:file = "$Folder\$File"
}

<#
.Synopsis
    Inserts a Table start tag in the outfile file.
.DESCRIPTION
    This is the detailed flow for creation of a table using this module.
    - 1. Start with Set-ReportFile and set the folder location along with an output file name
    - 2. Start Table creation process using Add-TableStart
    - 3. Set the width of columns here if required using Set-TableColumnWidth
    - 4. Start a Row using Add-TableStartRow
    - 5. Add Cells using Add-TableCells
    - 6. End a Row using Add-TableEndRow
    - 7. Repeat steps 4, 5, 6 until you are done adding all rows.
    - 8. End Table using Add-TableEnd

    Add-TableStart will start the table creation process. You don't have to remember the table syntax.

    Sample script that creates a table with 2 rows.

    Set-ReportFile -Folder "C:\Temp" -File "Sample.htm" -OverWrite $true
    Add-TableStart -Width 700 -CellPadding 3
    Set-TableColumnWidth 300, 300, 100
    Add-TableStartRow
    Add-TableCells -data "Name", "Address", "WebSite" -isHeader $true
    Add-TableEndRow
    Add-TableStartRow
    Add-TableCells -data "Rahul", "Soni", "www.dotnetscraps.com"
    Add-TableEndRow
    Add-TableEnd
    Get-Table -OpenInBrowser $true -InsertCSS "C:\Temp\style.css"

.EXAMPLE
    To add a Table start tag use this Command directly. It will create a table with default Width = 600, cellspacing = 0, cellpadding = 0, border color = black and border width = 1
   
    Add-TableStart
.EXAMPLE
    To add a Table with custom width use...

    Add-TableStart -Width 600
.EXAMPLE
    To add a Table with custom width, cellspacing, cellpadding and border use...

    Add-TableStart -Width 800 -CellSpacing 5 -CellPadding 5 -BorderWidth 1 -BorderColor "#33CCFF"
#>
function Add-TableStart()
{
    param(
        [Parameter(HelpMessage="Width of table in pixels. [Default is 600]")]
        [int]$Width=600,
        
        [Parameter(HelpMessage="Space between the cells [Default is 0]")] 
        [int]$CellSpacing=0,
        
        [Parameter(HelpMessage="Space between the content and border of the cell [Default is 2]")] 
        [int]$CellPadding=2, 

        [Parameter(HelpMessage="Width of the border in pixels [Default is 1]")]
        [int]$BorderWidth=1,
        
        [Parameter(HelpMessage="Color of border in Hex string or Web color. You can provide a valid color name like red or green. You can also provide a hex string. [Default is Black]")] 
        [string]$BorderColor="black"
    )
    if(IsReportNull -eq $true) { return }
    Write-Host -ForegroundColor Green "Table header is added in $script:File"
    Add-Content $script:file "<table class='psTable' width='$Width' cellspacing='$CellSpacing' cellpadding='$CellPadding' border='$borderWidth' borderColor='$borderColor' >"
}

<#
.Synopsis
   Sets the width of each column of the table.
.DESCRIPTION
    This is the detailed flow for creation of a table using this module.
    - 1. Start with Set-ReportFile and set the folder location along with an output file name
    - 2. Start Table creation process using Add-TableStart
    - 3. Set the width of columns here if required using Set-TableColumnWidth
    - 4. Start a Row using Add-TableStartRow
    - 5. Add Cells using Add-TableCells
    - 6. End a Row using Add-TableEndRow
    - 7. Repeat steps 4, 5, 6 until you are done adding all rows.
    - 8. End Table using Add-TableEnd

    Set-TableColumnWidth is used to set width for multiple columns.

    Sample script that creates a table with 2 rows.
    
    Set-ReportFile -Folder "C:\Temp" -File "Sample.htm" -OverWrite $true
    Add-TableStart -Width 700 -CellPadding 3
    Set-TableColumnWidth 300, 300, 100
    Add-TableStartRow
    Add-TableCells -data "Name", "Address", "WebSite" -isHeader $true
    Add-TableEndRow
    Add-TableStartRow
    Add-TableCells -data "Rahul", "Soni", "www.dotnetscraps.com"
    Add-TableEndRow
    Add-TableEnd
    Get-Table -OpenInBrowser $true -InsertCSS "C:\Temp\style.css"

.EXAMPLE
    Set the column width for 1st, 2nd and 3rd column as 300 px, 300px, and 100px respectively.

    Set-TableColumnWidth 300, 300, 100
.EXAMPLE
    Set each column as 100 px

    Set-TableColumnWidth 100
#>
function Set-TableColumnWidth()
{
    param(
        [parameter(Mandatory=$True, helpMessage="Specify width in pixel")]
        [int[]]$width
    )
    $string = ""
    $i = 0
    foreach($w in $width)
    {
        $string += "<col width='$w px'>"
    }
    Add-Content $script:file $string
}

<#
.Synopsis
   Start a row. Insert a <tr> tag in the table.
.DESCRIPTION
    This is the detailed flow for creation of a table using this module.
    - 1. Start with Set-ReportFile and set the folder location along with an output file name
    - 2. Start Table creation process using Add-TableStart
    - 3. Set the width of columns here if required using Set-TableColumnWidth
    - 4. Start a Row using Add-TableStartRow
    - 5. Add Cells using Add-TableCells
    - 6. End a Row using Add-TableEndRow
    - 7. Repeat steps 4, 5, 6 until you are done adding all rows.
    - 8. End Table using Add-TableEnd

    Add-TableStartRow will add the start tag for the row. <tr>

    Sample script that creates a table with 2 rows.
    
    Set-ReportFile -Folder "C:\Temp" -File "Sample.htm" -OverWrite $true
    Add-TableStart -Width 700 -CellPadding 3
    Set-TableColumnWidth 300, 300, 100
    Add-TableStartRow
    Add-TableCells -data "Name", "Address", "WebSite" -isHeader $true
    Add-TableEndRow
    Add-TableStartRow
    Add-TableCells -data "Rahul", "Soni", "www.dotnetscraps.com"
    Add-TableEndRow
    Add-TableEnd
    Get-Table -OpenInBrowser $true -InsertCSS "C:\Temp\style.css"

.EXAMPLE
    Add-TableStartRow
#>
function Add-TableStartRow()
{
    Add-Content $script:file "<tr>"
}

<#
.Synopsis
    Add-TableCells adds one or multiple cell(s) at a time in a row.
.DESCRIPTION
    This is the detailed flow for creation of a table using this module.
    - 1. Start with Set-ReportFile and set the folder location along with an output file name
    - 2. Start Table creation process using Add-TableStart
    - 3. Set the width of columns here if required using Set-TableColumnWidth
    - 4. Start a Row using Add-TableStartRow
    - 5. Add Cells using Add-TableCells
    - 6. End a Row using Add-TableEndRow
    - 7. Repeat steps 4, 5, 6 until you are done adding all rows.
    - 8. End Table using Add-TableEnd

    Add-TableCells is used to add table cells from left to right.

    Sample script that creates a table with 2 rows.
    
    Set-ReportFile -Folder "C:\Temp" -File "Sample.htm" -OverWrite $true
    Add-TableStart -Width 700 -CellPadding 3
    Set-TableColumnWidth 300, 300, 100
    Add-TableStartRow
    Add-TableCells -data "Name", "Address", "WebSite" -isHeader $true
    Add-TableEndRow
    Add-TableStartRow
    Add-TableCells -data "Rahul", "Soni", "www.dotnetscraps.com"
    Add-TableEndRow
    Add-TableEnd
    Get-Table -OpenInBrowser $true -InsertCSS "C:\Temp\style.css"
.EXAMPLE
    Adds 3 cells with Name, Address & WebSite. They will all be bold by default.

    Add-TableCells -data "Name", "Address", "WebSite" -isHeader $true
.EXAMPLE
    Add 3 cells. 1st cell contains Rahul. 2nd contains Soni with green background. 3rd contains www.dotnetscraps.com with red background.

    Add-TableCells -data "Rahul"
    Add-TableCells -data "Soni" -isGreen $true
    Add-TableCells -data "www.dotnetscraps.com" -isRed $true
.EXAMPLE
    Add 4 cells. 
    1st cell contains Rahul with green background and left alignment.
    2nd cell contains Soni with red background and center alignment. 
    3rd and 4th cell contains www.dotnetscraps.com with green background and right alignment

    Add-TableCells -data "Rahul" -isGreen $true -align "left"
    Add-TableCells -data "Soni" -isRed $true -align "center"
    Add-TableCells -data "Test 1", "Test2" -isGreen $true -align "right"

#>
function Add-TableCells()
{
    param(
        [parameter(mandatory=$true, helpMessage="Specify the data that is to be added to the cell(s)")]
        [string[]]$data, 
        
        [parameter(helpMessage="Is this a header? Default is false")]
        [boolean]$isHeader=$false,

        [parameter(helpMessage="Do you want to color this cell Green? Default is false")]
        [boolean]$isGreen=$false,

        [parameter(helpMessage="Do you want to color this cell Red? Default is false")]
        [boolean]$isRed=$false,
        
        [parameter(helpMessage="Do you want to align this cell left, center or right? Default is left")] 
        [string]$align="left"
    )
    if($isHeader -eq $true)
    {
        $string = ""
        foreach($val in $data)
        {
            if($iSGreen -eq $true)
            {
                $string += "<th class=psgreen>$val</th>"
            }
            elseif($isRed -eq $true)
            {
                $string += "<th class=psred>$val</th>"
            }
            else
            {
                $string += "<th class=psnormal>$val</th>"
            }
        }
        Add-Content $file $string
    }
    else
    {
        $alignRed = "ps$align psred"
        $alignGreen = "ps$align psgreen"
        $align = "ps$align"
        $string = ""
        foreach($val in $data)
        {
            if($iSGreen -eq $true)
            {
                $string += "<td class=""$alignGreen"">$val</td>"
            }
            elseif($isRed -eq $true)
            {
                $string += "<td class=""$alignRed"">$val</td>"
            }
            else
            {
                $string += "<td class=$align>$val</td>"
            }
        }
        Add-Content $file $string
    }
}

<#
.Synopsis
   Ends a row. Insert a </tr> tag in the table.
.DESCRIPTION
    This is the detailed flow for creation of a table using this module.
    - 1. Start with Set-ReportFile and set the folder location along with an output file name
    - 2. Start Table creation process using Add-TableStart
    - 3. Set the width of columns here if required using Set-TableColumnWidth
    - 4. Start a Row using Add-TableStartRow
    - 5. Add Cells using Add-TableCells
    - 6. End a Row using Add-TableEndRow
    - 7. Repeat steps 4, 5, 6 until you are done adding all rows.
    - 8. End Table using Add-TableEnd

    Add-TableEndRow will add the end tag for the row. </tr>

    Sample script that creates a table with 2 rows.
    
    Set-ReportFile -Folder "C:\Temp" -File "Sample.htm" -OverWrite $true
    Add-TableStart -Width 700 -CellPadding 3
    Set-TableColumnWidth 300, 300, 100
    Add-TableStartRow
    Add-TableCells -data "Name", "Address", "WebSite" -isHeader $true
    Add-TableEndRow
    Add-TableStartRow
    Add-TableCells -data "Rahul", "Soni", "www.dotnetscraps.com"
    Add-TableEndRow
    Add-TableEnd
    Get-Table -OpenInBrowser $true -InsertCSS "C:\Temp\style.css"

.EXAMPLE
    Add-TableEndRow
#>
function Add-TableEndRow()
{
    Add-Content $script:file "</tr>"
}

<#
.Synopsis
   Inserts a Table end tag in the outfile file. The output file could be set using Set-ReportFile
.DESCRIPTION
    This is the detailed flow for creation of a table using this module.
    - 1. Start with Set-ReportFile and set the folder location along with an output file name
    - 2. Start Table creation process using Add-TableStart
    - 3. Set the width of columns here if required using Set-TableColumnWidth
    - 4. Start a Row using Add-TableStartRow
    - 5. Add Cells using Add-TableCells
    - 6. End a Row using Add-TableEndRow
    - 7. Repeat steps 4, 5, 6 until you are done adding all rows.
    - 8. End Table using Add-TableEnd

    Add-TableEnd simply closes the table by putting a </table> tag in the end of the table. You don't have to remember the table syntax.

    Sample script that creates a table with 2 rows.

    Set-ReportFile -Folder "C:\Temp" -File "Sample.htm" -OverWrite $true
    Add-TableStart -Width 700 -CellPadding 3
    Set-TableColumnWidth 300, 300, 100
    Add-TableStartRow
    Add-TableCells -data "Name", "Address", "WebSite" -isHeader $true
    Add-TableEndRow
    Add-TableStartRow
    Add-TableCells -data "Rahul", "Soni", "www.dotnetscraps.com"
    Add-TableEndRow
    Add-TableEnd
    Get-Table -OpenInBrowser $true -InsertCSS "C:\Temp\style.css"
.EXAMPLE
   Add-TableEnd
#>
function Add-TableEnd()
{
    Add-Content $script:file "</table>"
}

<#
.Synopsis
   Set the Report folder and file location
.DESCRIPTION
    This is the detailed flow for creation of a table using this module.
    - 1. Start with Set-ReportFile and set the folder location along with an output file name
    - 2. Start Table creation process using Add-TableStart
    - 3. Set the width of columns here if required using Set-TableColumnWidth
    - 4. Start a Row using Add-TableStartRow
    - 5. Add Cells using Add-TableCells
    - 6. End a Row using Add-TableEndRow
    - 7. Repeat steps 4, 5, 6 until you are done adding all rows.
    - 8. End Table using Add-TableEnd

    Get-Table displays the table on screen, or in browser. It is useful to see the output directly from the script.

Sample script that creates a table with 2 rows.
    Set-ReportFile -Folder "C:\Temp" -File "Sample.htm" -OverWrite $true
    Add-TableStart -Width 700 -CellPadding 3
    Set-TableColumnWidth 300, 300, 100
    Add-TableStartRow
    Add-TableCells -data "Name", "Address", "WebSite" -isHeader $true
    Add-TableEndRow
    Add-TableStartRow
    Add-TableCells -data "Rahul", "Soni", "www.dotnetscraps.com"
    Add-TableEndRow
    Add-TableEnd
    Get-Table -OpenInBrowser $true -InsertCSS "C:\Temp\style.css"
.EXAMPLE
    Display the table on screen and open a browser to show the output.

    Get-Table -OpenInBrowser $true
.EXAMPLE
    Display the table on screen and open a browser to show the output. This also includes styles, and hence the table's output would be in a color format based on your style sheet.

    Get-Table -OpenInBrowser $true -InsertCSS "C:\Temp\style.css"
#>
function Get-Table()
{
    param(
        [parameter(helpMessage="Do you want to open the output in a browser? (default = false)")]
        [boolean]$OpenInBrowser=$false, 

        [parameter(helpMessage="Path of the CSS that you want to insert. The CSS should be available at the location of your helperTable.psm1 module")]
        [string]$InsertCSS
    )
    if(IsReportNull -eq $true) { return }
    Write-Host (Get-Content($script:file))
    if($OpenInBrowser -eq $true)
    {
        if($InsertCSS -ne $null)
        {
            if(Test-Path($InsertCSS))
            {
                $tmpString = Get-Content($script:file)
                Set-Content -Path $script:file -Value ("<head><style>" + (Get-Content($insertCSS)) + "</style></head>$tmpString")
            }
            else
            {
                Write-Host "The CSS file path does not exist. Please retry!"
            }
        }
        Start-Process $script:file
    }
}

function IsReportNull()
{
    if($script:File -eq $null)
    {
        Write-Host -ForegroundColor Red "There is no report to create. Please use Set-ReportFile to set a report file."
        return $true
    }
    else
    {
        return $false
    }
}

Export-ModuleMember -Function Add-TableStart
Export-ModuleMember -Function Add-TableEnd
Export-ModuleMember -Function Add-TableStartRow
Export-ModuleMember -Function Add-TableEndRow
Export-ModuleMember -Function Add-TableCells
Export-ModuleMember -Function Get-Table
Export-ModuleMember -Function Set-TableColumnWidth
Export-ModuleMember -Function Set-ReportFile
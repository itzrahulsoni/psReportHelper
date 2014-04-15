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
        [int[]]$Width
    )
    if(IsReportNull -eq $true) { return }
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
    if(IsReportNull -eq $true) { return }
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
        [string[]]$Data, 
        
        [parameter(helpMessage="Is this a header? Default is false")]
        [boolean]$IsHeader=$false,

        [parameter(helpMessage="Do you want to color this cell Green? Default is false")]
        [boolean]$IsGreen=$false,

        [parameter(helpMessage="Do you want to color this cell Red? Default is false")]
        [boolean]$IsRed=$false,
        
        [parameter(helpMessage="Do you want to align this cell left, center or right? Default is left")] 
        [string]$Align="left"
    )
    if(IsReportNull -eq $true) { return }
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
    if(IsReportNull -eq $true) { return }
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
    if(IsReportNull -eq $true) { return }
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

<#
.Synopsis
    Adds a text as a heading level 1.
.DESCRIPTION
    There are 6 heading text level you can select from Add-HeadingText1 to Add-HeadingText6 [Larger to Smaller]

.EXAMPLE
    Adds a heading text and aligns it to the center.
    
    Add-HeadingText1 -Message "Hello World - 1"

    or 

    Add-HeadingText1 -Message "Hello World - 1" -align "center"
.EXAMPLE
    Adds a heading text and aligns it to the left.

    Add-HeadingText1 -Message "Hello World - 1" -align "left"
.EXAMPLE
    Sample to add multiple headers in descending order

    Add-HeadingText1 -Message "Hello World - 1"
    Add-HeadingText2 -Message "Hello World - 2"
    Add-HeadingText3 -Message "Hello World - 3"
    Add-HeadingText4 -Message "Hello World - 4"
    Add-HeadingText5 -Message "Hello World - 5"
    Add-HeadingText6 -Message "Hello World - 6"
#>
function Add-HeadingText1()
{
    param(
        [parameter(Mandatory=$true, helpMessage="What do you want to print?")]
        [string]$Message, 

        [parameter(helpMessage="Do you want to align this cell left, center or right? Default is center")] 
        [string]$Align="center"
    )
    if(IsReportNull -eq $true) { return }
    Add-Content $script:file "<h1 class=ps$align>$Message</h1>"
}

<#
.Synopsis
    Adds a text as a heading level 2.
.DESCRIPTION
    There are 6 heading text level you can select from Add-HeadingText1 to Add-HeadingText6 [Larger to Smaller]

.EXAMPLE
    Adds a heading text and aligns it to the center.
    
    Add-HeadingText2 -Message "Hello World - 1"

    or 

    Add-HeadingText2 -Message "Hello World - 1" -align "center"
.EXAMPLE
    Adds a heading text and aligns it to the left.

    Add-HeadingText2 -Message "Hello World - 1" -align "left"
.EXAMPLE
    Sample to add multiple headers in descending order

    Add-HeadingText1 -Message "Hello World - 1"
    Add-HeadingText2 -Message "Hello World - 2"
    Add-HeadingText3 -Message "Hello World - 3"
    Add-HeadingText4 -Message "Hello World - 4"
    Add-HeadingText5 -Message "Hello World - 5"
    Add-HeadingText6 -Message "Hello World - 6"
#>
function Add-HeadingText2()
{
    param(
        [parameter(Mandatory=$true, helpMessage="What do you want to print?")]
        [string]$Message, 

        [parameter(helpMessage="Do you want to align this cell left, center or right? Default is center")] 
        [string]$Align="center"
    )
    if(IsReportNull -eq $true) { return }
    Add-Content $script:file "<h2 class=ps$align>$Message</h2>"
}

<#
.Synopsis
    Adds a text as a heading level 3.
.DESCRIPTION
    There are 6 heading text level you can select from Add-HeadingText1 to Add-HeadingText6 [Larger to Smaller]

.EXAMPLE
    Adds a heading text and aligns it to the center.
    
    Add-HeadingText3 -Message "Hello World - 1"

    or 

    Add-HeadingText3 -Message "Hello World - 1" -align "center"
.EXAMPLE
    Adds a heading text and aligns it to the left.

    Add-HeadingText3 -Message "Hello World - 1" -align "left"
.EXAMPLE
    Sample to add multiple headers in descending order

    Add-HeadingText1 -Message "Hello World - 1"
    Add-HeadingText2 -Message "Hello World - 2"
    Add-HeadingText3 -Message "Hello World - 3"
    Add-HeadingText4 -Message "Hello World - 4"
    Add-HeadingText5 -Message "Hello World - 5"
    Add-HeadingText6 -Message "Hello World - 6"
#>
function Add-HeadingText3()
{
    param(
        [parameter(Mandatory=$true, helpMessage="What do you want to print?")]
        [string]$Message, 

        [parameter(helpMessage="Do you want to align this cell left, center or right? Default is center")] 
        [string]$Align="center"
    )
    if(IsReportNull -eq $true) { return }
    Add-Content $script:file "<h3 class=ps$align>$Message</h3>"
}

<#
.Synopsis
    Adds a text as a heading level 4.
.DESCRIPTION
    There are 6 heading text level you can select from Add-HeadingText1 to Add-HeadingText6 [Larger to Smaller]

.EXAMPLE
    Adds a heading text and aligns it to the center.
    
    Add-HeadingText4 -Message "Hello World - 1"

    or 

    Add-HeadingText4 -Message "Hello World - 1" -align "center"
.EXAMPLE
    Adds a heading text and aligns it to the left.

    Add-HeadingText4 -Message "Hello World - 1" -align "left"
.EXAMPLE
    Sample to add multiple headers in descending order

    Add-HeadingText1 -Message "Hello World - 1"
    Add-HeadingText2 -Message "Hello World - 2"
    Add-HeadingText3 -Message "Hello World - 3"
    Add-HeadingText4 -Message "Hello World - 4"
    Add-HeadingText5 -Message "Hello World - 5"
    Add-HeadingText6 -Message "Hello World - 6"
#>
function Add-HeadingText4()
{
    param(
        [parameter(Mandatory=$true, helpMessage="What do you want to print?")]
        [string]$Message, 

        [parameter(helpMessage="Do you want to align this cell left, center or right? Default is center")] 
        [string]$Align="center"
    )
    if(IsReportNull -eq $true) { return }
    Add-Content $script:file "<h4 class=ps$align>$Message</h4>"
}

<#
.Synopsis
    Adds a text as a heading level 5.
.DESCRIPTION
    There are 6 heading text level you can select from Add-HeadingText1 to Add-HeadingText6 [Larger to Smaller]

.EXAMPLE
    Adds a heading text and aligns it to the center.
    
    Add-HeadingText5 -Message "Hello World - 1"

    or 

    Add-HeadingText5 -Message "Hello World - 1" -align "center"
.EXAMPLE
    Adds a heading text and aligns it to the left.

    Add-HeadingText5 -Message "Hello World - 1" -align "left"
.EXAMPLE
    Sample to add multiple headers in descending order

    Add-HeadingText1 -Message "Hello World - 1"
    Add-HeadingText2 -Message "Hello World - 2"
    Add-HeadingText3 -Message "Hello World - 3"
    Add-HeadingText4 -Message "Hello World - 4"
    Add-HeadingText5 -Message "Hello World - 5"
    Add-HeadingText6 -Message "Hello World - 6"
#>
function Add-HeadingText5()
{
    param(
        [parameter(Mandatory=$true, helpMessage="What do you want to print?")]
        [string]$Message, 

        [parameter(helpMessage="Do you want to align this cell left, center or right? Default is center")] 
        [string]$Align="center"
    )
    if(IsReportNull -eq $true) { return }
    Add-Content $script:file "<h5 class=ps$align>$Message</h5>"
}

<#
.Synopsis
    Adds a text as a heading level 6.
.DESCRIPTION
    There are 6 heading text level you can select from Add-HeadingText1 to Add-HeadingText6 [Larger to Smaller]

.EXAMPLE
    Adds a heading text and aligns it to the center.
    
    Add-HeadingText6 -Message "Hello World - 1"

    or 

    Add-HeadingText6 -Message "Hello World - 1" -align "center"
.EXAMPLE
    Adds a heading text and aligns it to the left.

    Add-HeadingText6 -Message "Hello World - 1" -align "left"
.EXAMPLE
    Sample to add multiple headers in descending order

    Add-HeadingText1 -Message "Hello World - 1"
    Add-HeadingText2 -Message "Hello World - 2"
    Add-HeadingText3 -Message "Hello World - 3"
    Add-HeadingText4 -Message "Hello World - 4"
    Add-HeadingText5 -Message "Hello World - 5"
    Add-HeadingText6 -Message "Hello World - 6"
#>
function Add-HeadingText6()
{
    param(
        [parameter(Mandatory=$true, helpMessage="What do you want to print?")]
        [string]$Message, 

        [parameter(helpMessage="Do you want to align this cell left, center or right? Default is center")] 
        [string]$Align="center"
    )
    if(IsReportNull -eq $true) { return }
    Add-Content $script:file "<h6 class=ps$align>$Message</h6>"
}

<#
.Synopsis
    Adds a text with appropriate font setting.
.DESCRIPTION
    Add a text to the Report file. It can be aligned left, center or right. You can also provide a font-size.

.EXAMPLE
    Add a text. Internally, it would add a <p> tag to the report.

    Add-Text "Hello World!"
.EXAMPLE
    Add a text with alignment. Internally, it would add a <p> tag to the report.

    Add-Text "Hello World" -align "Left"
.EXAMPLE
    Add a text with alignment and font size.

    Add-Text "Hello World" -align "left" -fontsize "20px"
#>
function Add-Text()
{
    param(
        [parameter(Mandatory=$true, helpMessage="What do you want to print?")]
        [string]$Message, 

        [parameter(helpMessage="Do you want to align this cell left, center or right? Default is center")] 
        [string]$Align="center",

        [parameter(helpMessage="What is the size of the font in px? Default is 15px")]
        [string]$FontSize="15px"
    )
    if(IsReportNull -eq $true) { return }
    Add-Content $script:file "<p style='font-size:$fontSize;text-align:$align'>$Message</p>"
}


<#
.Synopsis
    Add a timestamp in multiple formats
.DESCRIPTION
    Add a timestamp. There are multiple formats to choose from. See examples for details.

.EXAMPLE
    Add a time stamp aligned to center with a font size of 20 px.
    Add-TimeStamp -format 3 -align "center" -fontsize "20px"

.EXAMPLE
    Add a time stamp with a format

    Add-TimeStamp           #Output - Saturday, April 12, 2014 9:58:42 AM
    Add-TimeStamp -Format 1 #Output - Saturday, April 12, 2014 - 9:58:42 AM
    Add-TimeStamp -Format 2 #Output - 12/04/2014 - 9:58:42 AM
    Add-TimeStamp -Format 3 #Output - 04/12/2014 - 9:58:42 AM
    Add-TimeStamp -Format 4 #Output - 04/12/2014
    Add-TimeStamp -Format 5 #Output - 12/04/2014
    Add-TimeStamp -Format 6 #Output - 9:58:42 AM
    Add-TimeStamp -Format 7 #Output - 09:58:42 AM
    Add-TimeStamp -Format 8 #Output - 09:58 AM
    Add-TimeStamp -Format 9 #Output - 9:58 AM
#>
function Add-TimeStamp()
{
    param(
        [parameter(helpMessage="Enter a number (0-9) for a format. See help for more details.")] 
        [int]$Format=0,

        [parameter(helpMessage="Do you want to align this cell left, center or right? Default is center")] 
        [string]$Align="center",

        [parameter(helpMessage="What is the size of the font in px? Default is 15px")]
        [string]$FontSize="15px"   
    )
    if(IsReportNull -eq $true) { return }
    switch ($format)
    {
        0 { $time = (Get-Date) }
        1 { $time = (Get-Date).ToLongDateString() + " - " + (Get-Date).ToLongTimeString() }
        2 { $time = (Get-Date).ToString("dd/MM/yyyy") + " - " + (Get-Date).ToLongTimeString() }
        3 { $time = (Get-Date).ToString("MM/dd/yyyy") + " - " + (Get-Date).ToLongTimeString() }
        4 { $time = (Get-Date).ToString("MM/dd/yyyy") }
        5 { $time = (Get-Date).ToString("dd/MM/yyyy") }
        6 { $time = (Get-Date).ToLongTimeString() }
        7 { $time = (Get-Date).ToString("hh:mm:ss tt") }
        8 { $time = (Get-Date).ToString("hh:mm tt") }
        9 { $time = (Get-Date).ToString("h:mm tt") }
        Default { $time = (Get-Date) }
    }   
    Add-Text -Message $time -align $align -fontSize $fontSize
    return $time
}

<#
.Synopsis
    Add line break(s)
.DESCRIPTION
    Adds multiple line breaks. In effect it will insert a <BR /> tag in the html document

.EXAMPLE
    Add a line break
    
    Add-LineBreak
.EXAMPLE
    Add 3 line breaks

    Add-LineBreak -Count 3
#>
function Add-LineBreak()
{
    param(
        [parameter(helpMessage="Enter the number of line breaks you want to insert.")] 
        [int]$Count=1
    )
    if(IsReportNull -eq $true) { return }
    if($Count -le 0)
    {
        Write-Host -ForegroundColor Red "You must have at least 1 line break! Try again with a number greater than or equal to 1"
        return        
    }
    Add-Text -Message ("<BR />" * $Count)
}

<#
.Synopsis
    Start a list
.DESCRIPTION
    Start an list with different list styles

.EXAMPLE
    Start an unordered list with type = disc
    
    Add-ListStart
.EXAMPLE
    Start an unordered list with different types

    Add-ListStart -Style 0  # Disc
    Add-ListStart -Style 1  # Circle
    Add-ListStart -Style 2  # Square
    Add-ListStart -Style 3  # Decimal
    Add-ListStart -Style 4  # Lower Alphabets
    Add-ListStart -Style 5  # Lower Roman numbers
    Add-ListStart -Style 6  # Upper Alphabets
    Add-ListStart -Style 7  # Upper Roman
#>
function Add-ListStart()
{
    param(
        [parameter(helpMessage="Choose 0-7 for different styles of list")]
        [int]$Style=0
    )

    if(IsReportNull -eq $true) { return }
    $listType = "ul"
    switch ($Style)
    {
        0 { Add-Content $script:file ("<$listType style='list-style-type:disc'>") }
        1 { Add-Content $script:file ("<$listType style='list-style-type:circle'>") }
        2 { Add-Content $script:file ("<$listType style='list-style-type:square'>") }
        3 { Add-Content $script:file ("<$listType style='list-style-type:decimal'>") }
        4 { Add-Content $script:file ("<$listType style='list-style-type:lower-alpha'>") }
        5 { Add-Content $script:file ("<$listType style='list-style-type:lower-roman'>") }
        6 { Add-Content $script:file ("<$listType style='list-style-type:upper-alpha'>") }
        7 { Add-Content $script:file ("<$listType style='list-style-type:upper-roman'>") }
    }
}

<#
.Synopsis
    End a list
.DESCRIPTION
    End an ordered or unordered list

.EXAMPLE
    End a list that has already started
    
    Add-ListEnd
.EXAMPLE
    End an ordered list

    Add-ListEnd -IsOrdered $true
#>
function Add-ListEnd()
{
    if(IsReportNull -eq $true) { return }
    $listType = "ul"
    Add-Content $script:file ("</$listType>")
}

<#
.Synopsis
    Add a list item to the list
.DESCRIPTION
    Add as many items as you like using Add-ListItem

.EXAMPLE
    Add a List Item with a different style, with bold items and different colors

    Add-ListStart -Style 5
    Add-ListItem -Message "Item 1"
    Add-ListItem -Message "Item 2" -IsBold $true -Color Blue
    Add-ListItem -Message "Item 3" -IsBold $true -Color Red
    Add-ListItem -Message "Item 4" -Color Green
    Add-ListItem -Message "Item 5"
    Add-ListEnd 

.EXAMPLE
    Add an array of items as List Items with a different style, with bold items and different colors

    Add-ListStart -Style 3
    Add-ListItem -Message "Item 1", "Item 2", "Item 3" -IsBold $true
    Add-ListItem -Message "Item 4", "Item 5"
    Add-ListItem -Message "Item 6" -Color Red
    Add-ListEnd 
#>
function Add-ListItem()
{
    param(
        [parameter(helpMessage="What do you want to add?")]
        [string[]]$Message,

        [parameter(helpMessage="Is this bold? Default is false")]
        [boolean]$IsBold=$false,

        [Parameter(HelpMessage="Color of text in Hex string or Web color. You can provide a valid color name like red or green. You can also provide a hex string. [Default is Black]")] 
        [string]$Color="black",

        [parameter(helpMessage="What is the size of the font in px? Default is 15px")]
        [string]$FontSize="15px"
    )

    if(IsReportNull -eq $true) { return }
    $listType = "ul"
    foreach($line in $Message)
    {
        if($IsBold)
        {
            Add-Content $script:file ("<li style='font-size:$FontSize;color:$Color'><b>$line</b></li>")
        }
        else
        {
            Add-Content $script:file ("<li style='font-size:$FontSize;color:$Color'>$line</li>")
        }
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
Export-ModuleMember -Function Add-HeadingText1
Export-ModuleMember -Function Add-HeadingText2
Export-ModuleMember -Function Add-HeadingText3
Export-ModuleMember -Function Add-HeadingText4
Export-ModuleMember -Function Add-HeadingText5
Export-ModuleMember -Function Add-HeadingText6
Export-ModuleMember -Function Add-Text
Export-ModuleMember -Function Add-TimeStamp
Export-ModuleMember -Function Add-LineBreak
Export-ModuleMember -Function Add-ListStart
Export-ModuleMember -Function Add-ListItem
Export-ModuleMember -Function Add-ListEnd
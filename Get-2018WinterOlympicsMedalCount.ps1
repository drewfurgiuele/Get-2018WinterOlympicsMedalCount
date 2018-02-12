<#
    .SYNOPSIS
    Gets the HTML contents of the NBC Olympics Page and parses it to display current leaders

    .DESCRIPTION
    This function will parse the HTML contents of the NBC Sports Olympics Medal Count page and display results. Returns objects with countries and medal counts

    .PARAMETER Summary
    Formats static output with pretty colors for medals. Because reasons.

    .NOTES 
    Author: Drew Furgiuele (@pittfurg), http://www.port1433.com

    .EXAMPLE
    ./Get-2018WinterOlympicsMedalCount.ps1
    Uses Invoke-WebRequest to parse the current medal leaderboards and return an object containing any countries with medals
    
    ./Get-2018WinterOlympicsMedalCount.ps1 - Summary
    Returns static text with color-formatted results
#>

[CmdletBinding()]
Param (
    [parameter(Mandatory=$false)] 
    [switch] $Summary
)

begin {
    $medalContent = Invoke-WebRequest -Uri http://www.nbcolympics.com/medals -TimeoutSec 10
}

process {
    $standingsTable = $medalContent.ParsedHTML.GetElementsByTagName("table") | where-object {$_.ClassName -eq "grid-table"} | select -first 1
    for ($x = 1; $x -lt $standingsTable.rows.length; $x++) {
        $Country = [PSCustomObject] @{
            Place = $standingsTable.rows[$x].cells[0].innerText.trim()
            Name = $standingsTable.rows[$x].cells[1].innerText.trim()
            Gold = $standingsTable.rows[$x].cells[2].innerText.trim()
            Silver = $standingsTable.rows[$x].cells[3].innerText.trim()
            Bronze = $standingsTable.rows[$x].cells[4].innerText.trim()
        }
        if ($Summary) {
            Write-Host $standingsTable.rows[$x].cells[0].innerText.trim() -NoNewline
            Write-Host " " -NoNewline
            Write-Host $standingsTable.rows[$x].cells[1].innerText.trim() -NoNewline
            Write-Host " *" -ForegroundColor Yellow  -NoNewline
            Write-Host $standingsTable.rows[$x].cells[2].innerText.trim() -NoNewline
            Write-Host " *" -ForegroundColor Gray  -NoNewline
            Write-Host $standingsTable.rows[$x].cells[3].innerText.trim() -NoNewline
            Write-Host " *" -ForegroundColor DarkYellow -NoNewline
            Write-Host $standingsTable.rows[$x].cells[4].innerText.trim()
        } else {
            $Country
        }
    }
}